from ..app import mcp
from ..origin_connection import execute_labtalk, get_lt_var, get_lt_str
from ..labtalk_safe import labtalk_variable, safe_labtalk_script

@mcp.tool()
def run_labtalk(script: str) -> str:
    """Execute a LabTalk script in Origin Pro.

    Use this for any Origin operation not covered by other tools.
    LabTalk is Origin's built-in scripting language.

    Args:
        script: LabTalk script to execute

    Returns:
        Success/failure message
    """
    safe_script = safe_labtalk_script(script)
    success = execute_labtalk(safe_script)
    return f"Executed {'successfully' if success else 'with errors'}: {safe_script}"

@mcp.tool()
def get_labtalk_variable(name: str) -> str:
    """Get the value of a LabTalk variable.

    Args:
        name: Variable name. Use $ suffix for strings (e.g., 'str$')

    Returns:
        Variable value as string
    """
    safe_name = labtalk_variable(name, "name")
    if safe_name.endswith("$"):
        return get_lt_str(safe_name)
    else:
        return str(get_lt_var(safe_name))
