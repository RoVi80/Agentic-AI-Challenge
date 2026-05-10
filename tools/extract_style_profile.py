"""Extract a small visual style profile from a .pptx.

Returns ~6 fields that match the `settings` block consumed by
presentation-engine/src/siro_simple_render.js. Total output stays well under
1 KB, so the agent's input context is never flooded.

Run locally for testing:
    python tools/extract_style_profile.py path/to/some_deck.pptx
"""
from ibm_watsonx_orchestrate.agent_builder.tools import tool, WXOFile

import base64
import io
import json
import posixpath
import re
import struct
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
R_EMBED_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Common scheme-color aliases. PowerPoint themes interchangeably use bg1/lt1, tx1/dk1, etc.
SCHEME_ALIASES = {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}

# Position threshold for "near bottom" — manual footers usually sit below 6.8" (in EMU)
NEAR_BOTTOM_EMU = 6_217_920  # 6.8 inches in EMU (1in = 914400)
FOOTER_MAX_FONT_HUNDREDTHS = 1200  # 12pt — anything bigger is too prominent to be a footer


def _safe_hex(value, fallback):
    """Return value as uppercase 6-char hex; fall back if value isn't valid hex."""
    if isinstance(value, str) and re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        return value.upper()
    return fallback


def _slide_num(name):
    m = re.search(r"slide(\d+)\.xml", name)
    return int(m.group(1)) if m else 0


def _candidate_template_files(z):
    """Slide master XMLs + slide layout XMLs, masters first (more likely to hold THE brand logo)."""
    files = []
    files.extend(sorted(
        n for n in z.namelist()
        if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)
    ))
    files.extend(sorted(
        n for n in z.namelist()
        if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)
    ))
    return files


def _rels_path_for(xml_path):
    """Convention: ppt/foo/bar.xml has its rels at ppt/foo/_rels/bar.xml.rels"""
    parent = posixpath.dirname(xml_path)
    name = posixpath.basename(xml_path)
    return f"{parent}/_rels/{name}.rels"


LOGO_MIN_BYTES = 1024            # skip pixel-icons / sub-1KB decoration
LOGO_MAX_BYTES = 200 * 1024      # skip photos and large decorative artwork
LOGO_MIN_ASPECT = 1.3            # skip squarish images — wordmark logos are wide


def _png_aspect(img_bytes):
    """Return width/height aspect ratio if img_bytes is a PNG; else None."""
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w, h = struct.unpack(">II", img_bytes[16:24])
        if h == 0:
            return None
        return w / h
    except struct.error:
        return None


def _image_candidates_referenced_from(z, xml_path, min_bytes=LOGO_MIN_BYTES, max_bytes=LOGO_MAX_BYTES):
    """Return [(size_bytes, base64), ...] for PNG/JPG images referenced from xml_path that
    look "wordmark-like": file size in [min_bytes, max_bytes] AND (for PNGs) aspect ratio
    >= LOGO_MIN_ASPECT. JPEGs pass through without aspect check (no easy stdlib reader)."""
    rels_path = _rels_path_for(xml_path)
    if rels_path not in z.namelist():
        return []
    try:
        with z.open(xml_path) as f:
            xml = ET.fromstring(f.read())
        embed_ids = []
        for blip in xml.findall(".//a:blip", NS):
            rid = blip.get(R_EMBED_ATTR)
            if rid and rid not in embed_ids:
                embed_ids.append(rid)
        if not embed_ids:
            return []
        with z.open(rels_path) as f:
            rels_xml = ET.fromstring(f.read())
        rid_map = {
            rel.get("Id"): rel.get("Target")
            for rel in rels_xml.findall(f"{{{RELS_NS}}}Relationship")
        }
        out = []
        for rid in embed_ids:
            target = rid_map.get(rid)
            if not target:
                continue
            target_path = posixpath.normpath(
                posixpath.join(posixpath.dirname(xml_path), target)
            )
            if target_path not in z.namelist():
                continue
            ext = posixpath.splitext(target_path.lower())[1]
            if ext not in (".png", ".jpg", ".jpeg"):
                continue
            with z.open(target_path) as f:
                img_bytes = f.read()
            size = len(img_bytes)
            if size < min_bytes or size > max_bytes:
                continue
            # Aspect check for PNGs (wordmarks are wide)
            if ext == ".png":
                aspect = _png_aspect(img_bytes)
                if aspect is not None and aspect < LOGO_MIN_ASPECT:
                    continue
            out.append((size, base64.b64encode(img_bytes).decode("ascii")))
        return out
    except (ET.ParseError, KeyError, ValueError):
        return []


def _extract_master_logo(z):
    """Find the brand logo. Search order:
    1. The TITLE slide (slide1.xml) — by convention, every branded deck shows its logo there.
    2. Slide masters — for decks where the logo is a master-level fixture.
    3. Slide layouts — last resort (sometimes logos live only on the title-slide layout).

    Within each location, pick the smallest image in the [1KB, 200KB] range. Logos are
    typically simple shapes that compress small; decorative gradients and photos are larger.
    Returns base64 string or '' if nothing matches.
    """
    title_slide = "ppt/slides/slide1.xml"
    if title_slide in z.namelist():
        candidates = _image_candidates_referenced_from(z, title_slide)
        if candidates:
            candidates.sort(key=lambda c: c[0])
            return candidates[0][1]

    for xml_path in _candidate_template_files(z):
        candidates = _image_candidates_referenced_from(z, xml_path)
        if candidates:
            candidates.sort(key=lambda c: c[0])
            return candidates[0][1]

    return ""


def _extract_master_footer(z):
    """Footers in PowerPoint live on the master/layout, not on individual slides.
    Search for a footer placeholder (type='ftr') with text content.
    """
    for xml_path in _candidate_template_files(z):
        try:
            with z.open(xml_path) as f:
                xml = ET.fromstring(f.read())
            for sp in xml.findall(".//p:sp", NS):
                ph = sp.find(".//p:nvPr/p:ph", NS)
                if ph is None or ph.get("type") != "ftr":
                    continue
                txBody = sp.find("p:txBody", NS)
                if txBody is None:
                    continue
                texts = [t.text for t in txBody.findall(".//a:t", NS) if t.text]
                full = "".join(texts).strip()
                if 2 < len(full) < 100:
                    return full
        except (ET.ParseError, KeyError):
            continue
    return ""


def _read_theme_scheme(theme_xml):
    """Map scheme tokens (dk1, lt1, accent1, ...) to 6-char hex strings."""
    out = {}
    root = ET.fromstring(theme_xml)
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return out
    for child in scheme:
        token = child.tag.split("}")[-1]
        srgb = child.find("a:srgbClr", NS)
        sysc = child.find("a:sysClr", NS)
        if srgb is not None and srgb.get("val"):
            out[token] = srgb.get("val").upper()
        elif sysc is not None and sysc.get("lastClr"):
            out[token] = sysc.get("lastClr").upper()
    return out


def _read_theme_fonts(theme_xml):
    """Map OOXML font references (+mj-lt, +mn-lt, +mj-ea, +mn-ea) to real typeface names."""
    out = {}
    root = ET.fromstring(theme_xml)
    fs = root.find(".//a:fontScheme", NS)
    if fs is None:
        return out
    for ref, path in [
        ("+mj-lt", "a:majorFont/a:latin"),
        ("+mn-lt", "a:minorFont/a:latin"),
        ("+mj-ea", "a:majorFont/a:ea"),
        ("+mn-ea", "a:minorFont/a:ea"),
    ]:
        node = fs.find(path, NS)
        if node is not None and node.get("typeface"):
            out[ref] = node.get("typeface").strip()
    return out


def _resolve_color(fill_elem, scheme):
    """Given a parent that may contain srgbClr/schemeClr/sysClr, return hex or None."""
    if fill_elem is None:
        return None
    srgb = fill_elem.find("a:srgbClr", NS)
    if srgb is not None and srgb.get("val"):
        return srgb.get("val").upper()
    sch = fill_elem.find("a:schemeClr", NS)
    if sch is not None:
        val = sch.get("val", "")
        val = SCHEME_ALIASES.get(val, val)
        return scheme.get(val)
    sysc = fill_elem.find("a:sysClr", NS)
    if sysc is not None and sysc.get("lastClr"):
        return sysc.get("lastClr").upper()
    return None


def _walk_slide(slide_xml, scheme, theme_fonts, fonts, title_colors, body_colors, bg, footers, fills):
    root = ET.fromstring(slide_xml)

    # Slide background (if explicitly set)
    bg_elem = root.find(".//p:bg/p:bgPr", NS)
    if bg_elem is not None:
        sf = bg_elem.find("a:solidFill", NS)
        if sf is not None:
            c = _resolve_color(sf, scheme)
            if c:
                bg[c] += 1

    for sp in root.findall(".//p:sp", NS):
        # Footer placeholder text (gives us footerText for free)
        ph = sp.find(".//p:nvPr/p:ph", NS)
        is_footer = ph is not None and ph.get("type") == "ftr"

        # Position check — manual footers usually live at the bottom of the slide
        spPr = sp.find("p:spPr", NS)
        is_near_bottom = False
        if spPr is not None:
            xfrm = spPr.find("a:xfrm", NS)
            if xfrm is not None:
                off = xfrm.find("a:off", NS)
                if off is not None and off.get("y"):
                    try:
                        is_near_bottom = int(off.get("y")) > NEAR_BOTTOM_EMU
                    except (ValueError, TypeError):
                        pass

        # Shape fill — feeds the accent-color heuristic
        if spPr is not None:
            sf = spPr.find("a:solidFill", NS)
            if sf is not None:
                c = _resolve_color(sf, scheme)
                if c:
                    fills[c] += 1

        # Walk text runs
        txBody = sp.find("p:txBody", NS)
        if txBody is None:
            continue

        # Track this shape's full text and max font size to detect manual-footer textboxes
        shape_text_parts = []
        shape_max_size = 0

        for para in txBody.findall("a:p", NS):
            for run in para.findall("a:r", NS):
                t = run.find("a:t", NS)
                if t is None or not t.text:
                    continue
                rPr = run.find("a:rPr", NS)

                shape_text_parts.append(t.text)
                if rPr is not None and rPr.get("sz"):
                    try:
                        shape_max_size = max(shape_max_size, int(rPr.get("sz")))
                    except (ValueError, TypeError):
                        pass

                # Footer text capture (placeholder-typed footer)
                if is_footer:
                    text = t.text.strip()
                    if 2 < len(text) < 100:
                        footers[text] += 1

                # Font face — weighted by character count, theme refs resolved
                if rPr is not None:
                    latin = rPr.find("a:latin", NS)
                    if latin is not None and latin.get("typeface"):
                        face = latin.get("typeface").strip()
                        face = theme_fonts.get(face, face)
                        fonts[face] += len(t.text)

                # Color — buckets into title vs body by font size
                color = None
                if rPr is not None:
                    sf = rPr.find("a:solidFill", NS)
                    if sf is not None:
                        color = _resolve_color(sf, scheme)
                size_hundredths = int(rPr.get("sz")) if (rPr is not None and rPr.get("sz")) else 1800
                if color:
                    if size_hundredths >= 2000:  # >= 20pt → title-ish
                        title_colors[color] += len(t.text)
                    else:
                        body_colors[color] += len(t.text)

        # After processing this shape's runs: detect manual-textbox footers
        # (small text near the bottom of the slide, but NOT the placeholder we already handled).
        if is_near_bottom and not is_footer:
            full_text = "".join(shape_text_parts).strip()
            if (2 < len(full_text) < 80
                    and 0 < shape_max_size <= FOOTER_MAX_FONT_HUNDREDTHS):
                footers[full_text] += 1


def _most_common(counter, fallback):
    if not counter:
        return fallback
    return counter.most_common(1)[0][0]


def _extract_from_bytes(file_bytes):
    """Pure function: bytes → style profile dict. Used by both the WXO tool and CLI."""
    fonts = Counter()
    title_colors = Counter()
    body_colors = Counter()
    bg = Counter()
    footers = Counter()
    fills = Counter()
    scheme = {}
    theme_fonts = {}
    logo_b64 = ""

    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        # Match only ppt/theme/themeN.xml — exclude themeOverride*.xml which has no fontScheme
        theme_files = sorted(n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n))
        if theme_files:
            with z.open(theme_files[0]) as f:
                theme_xml = f.read()
            scheme = _read_theme_scheme(theme_xml)
            theme_fonts = _read_theme_fonts(theme_xml)

        # Pull the brand logo from master/layouts (best-effort)
        logo_b64 = _extract_master_logo(z)
        # Pull footer text from master/layout placeholders (best-effort)
        master_footer = _extract_master_footer(z)

        slide_files = sorted(
            (n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")),
            key=_slide_num,
        )
        for sf_name in slide_files:
            with z.open(sf_name) as f:
                _walk_slide(f.read(), scheme, theme_fonts, fonts, title_colors, body_colors, bg, footers, fills)

    backgroundColor = _most_common(bg, scheme.get("lt1", "FFFFFF"))

    # Body color: pick most-common body text color that's NOT the background
    # (otherwise white-on-dark section slides poison the count).
    bodyColor = scheme.get("dk1", "1D1D1D")
    for color, _ in body_colors.most_common():
        if color != backgroundColor:
            bodyColor = color
            break

    # Title color: same trick as body — skip colors that match the background
    # (otherwise white titles on dark cover slides poison the count).
    titleColor = scheme.get("dk1", "1D1D1D")
    for color, _ in title_colors.most_common():
        if color != backgroundColor:
            titleColor = color
            break

    # Accent color: prefer the most-used non-background non-text shape fill;
    # fall back to theme accent1.
    accent_candidates = Counter()
    for color, count in fills.items():
        if color in {bodyColor, backgroundColor, titleColor}:
            continue
        accent_candidates[color] = count

    # Drop any unresolved theme-font tokens (e.g. "+mj-lt") so they never leak into output
    clean_fonts = Counter({k: v for k, v in fonts.items() if not k.startswith("+")})

    title_default = scheme.get("dk1", "1D1D1D")
    accent_default = scheme.get("accent1", "D0006F")

    profile = {
        "fontFace": _most_common(clean_fonts, "Arial"),
        "titleColor": _safe_hex(titleColor, title_default),
        "bodyColor": _safe_hex(bodyColor, title_default),
        "backgroundColor": _safe_hex(backgroundColor, "FFFFFF"),
        "accentColor": _safe_hex(_most_common(accent_candidates, accent_default), accent_default),
        # Slide-detected footer wins; otherwise use master/layout footer placeholder
        "footerText": _most_common(footers, master_footer),
    }
    # logo_b64 intentionally not returned — extraction proved unreliable across decks
    # (most brands draw their wordmark as text on the master, not as an image asset).
    # Helpers (_extract_master_logo, _image_candidates_referenced_from) kept around
    # for a future "explicit logo upload" feature.
    _ = logo_b64
    return profile


@tool(
    name="extract_style_profile",
    description=(
        "Extracts a small visual style profile (font, colors, footer text) from an uploaded PowerPoint file. "
        "Returns a compact JSON object suitable for use as the 'settings' block when generating a new deck "
        "via siro_simple_render. Use this when the user wants the new deck to match the look of an existing pptx."
    ),
)
def extract_style_profile(pptx_file: WXOFile) -> str:
    file_bytes = WXOFile.get_content(pptx_file)
    profile = _extract_from_bytes(file_bytes)
    return json.dumps(profile)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python tools/extract_style_profile.py <path/to/deck.pptx>")
        sys.exit(1)
    with open(sys.argv[1], "rb") as f:
        profile = _extract_from_bytes(f.read())
    print(json.dumps(profile, indent=2))
