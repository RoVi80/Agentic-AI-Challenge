"""Side-by-side comparison: Sarina's extract_pptx_formatting vs the new extract_style_profile.

Run:
    python prototype/compare_outputs.py "path/to/deck.pptx"
"""
import io
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Make tools/ importable
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from pptx_format_tool import parse_slide  # noqa: E402
from extract_style_profile import _extract_from_bytes  # noqa: E402


def _slide_num(name):
    m = re.search(r"slide(\d+)\.xml", name)
    return int(m.group(1)) if m else 0


def run_sarina(file_bytes):
    """Replicates the inner logic of Sarina's extract_pptx_formatting (without WXOFile wrapper)."""
    result = {}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        slide_files = sorted(
            (n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")),
            key=_slide_num,
        )
        for sf in slide_files:
            with z.open(sf) as f:
                result[sf] = parse_slide(ET.parse(f))
    return json.dumps(result, ensure_ascii=False, indent=2)


def run_distilled(file_bytes):
    return json.dumps(_extract_from_bytes(file_bytes), ensure_ascii=False, indent=2)


def fmt_bytes(n):
    if n >= 1024 * 1024:
        return f"{n:,} bytes ({n / (1024 * 1024):.2f} MB)"
    if n >= 1024:
        return f"{n:,} bytes ({n / 1024:.1f} KB)"
    return f"{n:,} bytes"


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python prototype/compare_outputs.py <path/to/deck.pptx>")
        sys.exit(1)

    with open(sys.argv[1], "rb") as f:
        data = f.read()

    sarina_out = run_sarina(data)
    distilled_out = run_distilled(data)

    print("=" * 70)
    print("Sarina  ::  extract_pptx_formatting (raw OOXML formatting)")
    print("=" * 70)
    print(f"Size: {fmt_bytes(len(sarina_out.encode('utf-8')))}")
    print(f"First 600 chars (truncated):")
    print(sarina_out[:600])
    print("...")
    print()
    print("=" * 70)
    print("Distilled  ::  extract_style_profile (6-field summary)")
    print("=" * 70)
    print(f"Size: {fmt_bytes(len(distilled_out.encode('utf-8')))}")
    print(f"Full output:")
    print(distilled_out)
    print()
    print("=" * 70)
    ratio = len(sarina_out) / max(len(distilled_out), 1)
    print(f"Ratio: Sarina output is ~{ratio:.0f}x larger than the distilled summary.")
    print("=" * 70)
