from ibm_watsonx_orchestrate.agent_builder.tools import tool, WXOFile

import io
import json
import re
import zipfile
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def _slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _emu_to_in(value):
    try:
        return round(int(value) / 914400, 4)
    except Exception:
        return None


def _get_position(shape):
    """
    Extract approximate x/y position of a shape.
    This helps us keep text blocks in a sensible visual order.
    """
    xfrm = shape.find(".//a:xfrm", NS)
    if xfrm is None:
        return {"x_in": None, "y_in": None}

    off = xfrm.find("a:off", NS)
    if off is None:
        return {"x_in": None, "y_in": None}

    return {
        "x_in": _emu_to_in(off.get("x")),
        "y_in": _emu_to_in(off.get("y")),
    }


def _get_placeholder_type(shape):
    """
    PowerPoint uses placeholders like title, subtitle, body, etc.
    This is useful for detecting slide titles.
    """
    ph = shape.find(".//p:nvPr/p:ph", NS)
    if ph is None:
        return None

    # If type is missing, PowerPoint often means body placeholder.
    return ph.get("type") or "body"


def _extract_paragraphs(shape):
    paragraphs = []

    for para in shape.findall(".//a:p", NS):
        parts = []

        for node in para:
            tag = node.tag.split("}")[-1]

            if tag == "r":
                text_node = node.find("a:t", NS)
                if text_node is not None and text_node.text:
                    parts.append(text_node.text)

            elif tag == "br":
                parts.append("\n")

        paragraph_text = "".join(parts).strip()

        if paragraph_text:
            paragraphs.append(paragraph_text)

    return paragraphs


def _extract_text_blocks(slide_root):
    text_blocks = []

    sp_tree = slide_root.find(".//p:spTree", NS)
    if sp_tree is None:
        return text_blocks

    for shape in sp_tree:
        tag = shape.tag.split("}")[-1]

        # Normal text boxes and shapes
        if tag != "sp":
            continue

        paragraphs = _extract_paragraphs(shape)

        if not paragraphs:
            continue

        c_nv_pr = shape.find(".//p:cNvPr", NS)
        shape_name = c_nv_pr.get("name") if c_nv_pr is not None else None

        placeholder = _get_placeholder_type(shape)
        position = _get_position(shape)

        text_blocks.append(
            {
                "shape_name": shape_name,
                "placeholder": placeholder,
                "position": position,
                "paragraphs": paragraphs,
                "text": "\n".join(paragraphs),
            }
        )

    # Sort roughly top-to-bottom, left-to-right
    text_blocks.sort(
        key=lambda block: (
            block["position"]["y_in"] if block["position"]["y_in"] is not None else 999,
            block["position"]["x_in"] if block["position"]["x_in"] is not None else 999,
        )
    )

    return text_blocks


def _guess_title(text_blocks):
    """
    Prefer real PowerPoint title placeholders.
    If none exist, use the first text block as fallback.
    """
    for block in text_blocks:
        if block.get("placeholder") in {"title", "ctrTitle"}:
            return block["paragraphs"][0]

    if text_blocks:
        return text_blocks[0]["paragraphs"][0]

    return ""


def _guess_subtitle(text_blocks):
    for block in text_blocks:
        if block.get("placeholder") == "subTitle":
            return block["text"]

    return ""


def _extract_body_lines(text_blocks, title, subtitle):
    lines = []

    for block in text_blocks:
        placeholder = block.get("placeholder")

        # Avoid duplicating title/subtitle in body content
        if placeholder in {"title", "ctrTitle", "subTitle"}:
            continue

        for paragraph in block["paragraphs"]:
            if paragraph != title and paragraph != subtitle:
                lines.append(paragraph)

    return lines


def _parse_slide(slide_xml_bytes, slide_number):
    root = ET.fromstring(slide_xml_bytes)

    text_blocks = _extract_text_blocks(root)

    title = _guess_title(text_blocks)
    subtitle = _guess_subtitle(text_blocks)
    body_lines = _extract_body_lines(text_blocks, title, subtitle)

    return {
        "slide_number": slide_number,
        "title": title,
        "subtitle": subtitle,
        "body_lines": body_lines,
        "text_blocks": text_blocks,
    }


@tool(
    name="siro_extract_pptx",
    description=(
        "Extracts slide titles, subtitles, and body text from an uploaded PowerPoint file. "
        "Use this when a user uploads an existing PPTX and wants to preserve the content "
        "while reformatting it into a new presentation."
    ),
)
def extract_pptx_content(pptx_file: WXOFile) -> str:
    """
    Extracts readable slide content from a PowerPoint file.

    Args:
        pptx_file (WXOFile): Uploaded .pptx file.

    Returns:
        str: JSON string containing slide titles, subtitles, body lines, and raw text blocks.
    """
    file_bytes = WXOFile.get_content(pptx_file)

    result = {
        "source_type": "pptx",
        "slide_count": 0,
        "slides": [],
    }

    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        slide_files = sorted(
            [
                name
                for name in z.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ],
            key=_slide_number,
        )

        for slide_file in slide_files:
            slide_number = _slide_number(slide_file)

            with z.open(slide_file) as f:
                slide_xml_bytes = f.read()

            result["slides"].append(
                _parse_slide(slide_xml_bytes, slide_number)
            )

    result["slide_count"] = len(result["slides"])

    return json.dumps(result, ensure_ascii=False, indent=2)
