import os

from mcp.server.fastmcp import FastMCP

from .skills import SERVER_INSTRUCTIONS

mcp = FastMCP("origin-pro", log_level="ERROR", instructions=SERVER_INSTRUCTIONS)

# Cutover flag: Phase 3d FLIPPED this to True — the daemon-backed shim is now the
# default entrypoint (auto-spawn + orphan self-cleanup). Force the legacy
# in-process server with ORIGIN_PRO_MCP_USE_DAEMON=0.
USE_DAEMON_DEFAULT = True


def use_daemon() -> bool:
    """Whether ``main()`` should run the daemon-backed shim instead of the
    in-process server. Reads ``ORIGIN_PRO_MCP_USE_DAEMON`` (default OFF)."""
    flag = os.environ.get("ORIGIN_PRO_MCP_USE_DAEMON")
    if flag is None:
        return USE_DAEMON_DEFAULT
    return flag.strip().lower() in ("1", "true", "yes", "on")


def main():
    """Entry point for uvx / console_scripts."""
    if use_daemon():
        from .shim import run_stdio

        run_stdio()
        return
    from . import server  # noqa: F401 — registers all tools
    mcp.run(transport="stdio")
