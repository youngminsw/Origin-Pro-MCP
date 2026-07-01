import json

from ..app import mcp
from ..origin_connection import execute_labtalk, get_lt_var, get_lt_str
from ..labtalk_safe import labtalk_variable, classify_labtalk_script

@mcp.tool()
def run_labtalk(script: str, confirm: bool = False, capture: list[str] | None = None,
                timeout: float = 0.0) -> str:
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

    Observability: LabTalk's `type`/`print` output cannot be read back over
    COM, so to inspect a computed value assign it to an (untyped) LabTalk
    variable inside `script` and name it in `capture` — e.g.
    `run_labtalk("stats col(1); mean = stats.mean;", capture=["mean"])`.
    The variables' values are read back after the script runs and returned as
    JSON. Use a `$` suffix for string variables (e.g. `capture=["name$"]`).

    Args:
        script: LabTalk script to execute
        confirm: Set True to run a script that uses a gated command token
        capture: Optional list of LabTalk variable names to read back after
                 the script runs. When given, the result is a JSON object
                 {"status", "script", "values"}.
        timeout: Optional per-call dispatch budget in seconds for this one
                 script. 0 (default) uses the daemon's configured dispatch
                 timeout. A positive value bounds this call even when the
                 global dispatch timeout is off; on overrun the daemon force-
                 resets the wedged session and returns an actionable error.

    Returns:
        Success/failure message, or an actionable not-executed message when a
        gated token is present and confirm is False. When `capture` is given,
        a JSON object with the executed status and the captured variable values.
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
    status = "ok" if success else "error"
    if not capture:
        return f"Executed {'successfully' if success else 'with errors'}: {script}"
    values: dict = {}
    for raw_name in capture:
        safe_name = labtalk_variable(raw_name, "capture")
        values[safe_name] = (
            get_lt_str(safe_name) if safe_name.endswith("$") else get_lt_var(safe_name)
        )
    return json.dumps({"status": status, "script": script, "values": values})

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
