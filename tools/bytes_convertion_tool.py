from ibm_watsonx_orchestrate.agent_builder.tools import tool
import base64


@tool(
    name="base64_to_bytes",
    description="A tool to convert base64 strings of files to a bytes object.",
)
def base64_to_bytes(base64_string: str) -> bytes:
    """
    Convert a base64 string of a file to a bytes object.

    Args:
        base64_string (str): A base64 string of a file.

    Returns:
        bytes: A bytes object of the file.
    """
    return base64.b64decode(base64_string)