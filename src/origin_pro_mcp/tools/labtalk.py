from ..app import mcp
from ..origin_connection import execute_labtalk, get_lt_var, get_lt_str
from ..labtalk_safe import labtalk_variable, classify_labtalk_script

@mcp.tool()
def run_labtalk(script: str, confirm: bool = False) -> str:
    """Execute a LabTalk script in Origin Pro — the universal escape hatch.

    Use this for any Origin operation not covered by other tools. LabTalk is
    Origin's built-in scripting language. Most operations (data, plotting,
    styling, save, saveAs, open, expGraph, file, etc.) run freely.

    A narrow confirm gate protects a few real command tokens that can touch
    files, the project, or the system (e.g. `doc -s`/`doc -n`, `del`/`delete`,
    `win -c`/`-cd`/`-ct`, `system.`/`system(`, `run.section`/`run -*`, `dll`,
    `dde`, `getfilename`, `getsavename`, `label -r`). Keywords inside string
    literals or comments never trigger the gate. When a gated token is present
    and `confirm` is False, the script is NOT executed and a message naming the
    token is returned; re-call with `confirm=True` to run it anyway.

    Args:
        script: LabTalk script to execute
        confirm: Set True to run a script that uses a gated command token

    Returns:
        Success/failure message, or an actionable not-executed message when a
        gated token is present and confirm is False
    """
    _ok, requires_confirm, reason = classify_labtalk_script(script)
    if requires_confirm and not confirm:
        return (
            f"NOT EXECUTED. This script uses a gated command token ('{reason}') "
            "that can affect files, the project, or the system. Review it, then "
            "re-call run_labtalk with confirm=True to run it anyway.\n"
            f"Script: {script}"
        )
    success = execute_labtalk(script)
    return f"Executed {'successfully' if success else 'with errors'}: {script}"

@mcp.tool()
def get_labtalk_variable(name: str) -> str:
    """Get the value of a LabTalk variable.

    Gotchas: numeric variables that don't exist read as 0, and variables
    declared with a type (e.g. `int x = 5`) are script-local — they vanish
    when the script ends. Use untyped assignment (`x = 5`) in run_labtalk
    if you want to read the value back later.

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
