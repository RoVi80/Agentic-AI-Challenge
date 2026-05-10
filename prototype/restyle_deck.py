"""Per-slide parallel restyle prototype.

Run:
    setx GROQ_API_KEY "<your-key>"   # PowerShell: $env:GROQ_API_KEY="..."
    python prototype/restyle_deck.py prototype/sample_deck.json --prompt "make bullets more concise and executive-toned"
"""
import argparse
import asyncio
import json
import os
import time
from pathlib import Path

from openai import AsyncOpenAI

MODEL = "openai/gpt-oss-120b"

SYSTEM = """You restyle a single PowerPoint slide.
Input: one slide as JSON. Output: the restyled slide as a single JSON object.

Hard rules:
- Preserve all factual content (numbers, dates, names, product terms).
- Preserve the slide type field exactly.
- Preserve all keys present in the input slide.
- Keep bullets concise (max 12 words each).
- Apply the user instruction.
- Output ONLY the JSON object, no markdown, no commentary."""


async def restyle_slide(client, slide, style_guide, user_prompt, slide_idx):
    user_msg = (
        f"STYLE GUIDE:\n{json.dumps(style_guide, indent=2)}\n\n"
        f"USER INSTRUCTION:\n{user_prompt}\n\n"
        f"SLIDE (index {slide_idx}):\n{json.dumps(slide, indent=2)}"
    )
    t0 = time.perf_counter()
    resp = await client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": user_msg},
        ],
        temperature=0.3,
        response_format={"type": "json_object"},
    )
    dt = time.perf_counter() - t0
    restyled = json.loads(resp.choices[0].message.content)
    print(f"  slide {slide_idx} ({slide.get('type','?')})  restyled in {dt:.1f}s")
    return restyled


async def main(args):
    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise SystemExit("Set GROQ_API_KEY environment variable first.")

    client = AsyncOpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")

    deck = json.loads(Path(args.deck).read_text(encoding="utf-8"))
    style_guide = json.loads(Path(args.style_guide).read_text(encoding="utf-8"))

    n = len(deck["slides"])
    print(f"Restyling {n} slides in parallel against {MODEL}")
    print(f"Prompt: {args.prompt!r}\n")

    t0 = time.perf_counter()
    restyled_slides = await asyncio.gather(
        *(
            restyle_slide(client, slide, style_guide, args.prompt, i)
            for i, slide in enumerate(deck["slides"])
        )
    )
    elapsed = time.perf_counter() - t0

    new_deck = {**deck, "slides": restyled_slides}
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(new_deck, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nTotal wall time: {elapsed:.1f}s ({elapsed/n:.1f}s per slide on average)")
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("deck", help="Input deck JSON (e.g. prototype/sample_deck.json)")
    p.add_argument("--prompt", required=True, help="Restyling instruction")
    p.add_argument("--style-guide", default="knowledge/specs.json")
    p.add_argument("--out", default="prototype/output_deck.json")
    asyncio.run(main(p.parse_args()))
