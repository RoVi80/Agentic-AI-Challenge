"""Diagnostic: list every image inside a .pptx with its size, and which master/layout/slide
references it. Helps us pick the right logo extraction heuristic.

Run:
    python prototype/inspect_images.py "path/to/deck.pptx"
"""
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def list_xml_files(z, prefix_pattern):
    return sorted(n for n in z.namelist() if re.match(prefix_pattern, n))


def find_image_refs(z, xml_path):
    """Return list of resolved image paths referenced from this XML."""
    rels_path = f"{posixpath.dirname(xml_path)}/_rels/{posixpath.basename(xml_path)}.rels"
    if rels_path not in z.namelist():
        return []
    try:
        with z.open(xml_path) as f:
            xml = ET.fromstring(f.read())
        embed_ids = []
        for blip in xml.findall(".//a:blip", NS):
            rid = blip.get(R_EMBED)
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
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(xml_path), target))
            out.append(resolved)
        return out
    except ET.ParseError:
        return []


def main(path):
    with zipfile.ZipFile(path) as z:
        # All media images and their sizes
        all_media = sorted(
            (n, z.getinfo(n).file_size)
            for n in z.namelist()
            if n.startswith("ppt/media/") and n.lower().endswith((".png", ".jpg", ".jpeg"))
        )
        print(f"=== Images in ppt/media ({len(all_media)} total) ===")
        for path_, size in all_media:
            print(f"  {path_:<40} {size:>10,} bytes")

        # Reference graph: which XMLs reference which images
        refs = defaultdict(list)
        for xml_pattern in [
            r"ppt/slideMasters/slideMaster\d+\.xml$",
            r"ppt/slideLayouts/slideLayout\d+\.xml$",
            r"ppt/slides/slide\d+\.xml$",
        ]:
            for xml in list_xml_files(z, xml_pattern):
                for img in find_image_refs(z, xml):
                    refs[img].append(xml)

        print(f"\n=== Reference graph ===")
        for img, _ in all_media:
            ref_list = refs.get(img, [])
            label = ", ".join(r.replace("ppt/", "") for r in ref_list[:5])
            if len(ref_list) > 5:
                label += f", … (+{len(ref_list) - 5} more)"
            print(f"  {img:<40} referenced by [{len(ref_list)}]: {label or '(none)'}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python prototype/inspect_images.py <path/to/deck.pptx>")
        sys.exit(1)
    main(sys.argv[1])
