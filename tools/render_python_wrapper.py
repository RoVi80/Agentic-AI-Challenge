import subprocess
import json
from pathlib import Path


def render_presentation_python_wrapped(input_json: dict):
    script_path = Path(__file__).parent / "render_janine_6.js"

    # Run Node
    result = subprocess.run(
        ["node", str(script_path)],
        input=json.dumps(input_json),
        text=True,
        capture_output=True
    )

    if result.returncode != 0:
        raise Exception(result.stderr)

    # Parse Node output
    try:
        output = json.loads(result.stdout)
    except json.JSONDecodeError:
        raise Exception(f"Invalid Node output: {result.stdout}")

    pptx_path = Path(output["output_file"])

    if not pptx_path.exists():
        raise Exception(f"PPTX file not found: {pptx_path}")

    # READ FILE AS BINARY
    file_bytes = pptx_path.read_bytes()

    # ORCHESTRATE ARTIFACT FORMAT
    return {
        "type": "file",
        "filename": pptx_path.name,
        "content": file_bytes,
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }
