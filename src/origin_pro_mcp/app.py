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


_SKIP_PLATFORM_CHECK_ENV = "ORIGIN_PRO_MCP_SKIP_PLATFORM_CHECK"


def platform_problem():
    """Why this host cannot run the server, or ``None`` if it can.

    Origin's COM automation exists only on Windows, and the server is tested on
    Windows 10/11 (NT 10.0 — which also covers Windows Server 2016 and later;
    Windows 11 reports major version 10 too). Anywhere else the server would
    start and advertise every tool, only for each call to fail, so ``main()``
    refuses to start instead. Set ``ORIGIN_PRO_MCP_SKIP_PLATFORM_CHECK=1`` to
    bypass (e.g. an unusual Windows-compatible host).

    Only the MCP entry point checks this — importing the package, the daemon
    module and the CLI stay usable everywhere (the unit tests run on Linux).
    """
    import platform
    import sys

    if os.environ.get(_SKIP_PLATFORM_CHECK_ENV, "").strip().lower() in (
            "1", "true", "yes", "on"):
        return None
    if sys.platform != "win32":
        release = platform.release().lower()
        if "microsoft" in release or "wsl" in release:
            return (
                "origin-pro-mcp only runs on Windows 10/11, and this is a Linux "
                "Python inside WSL. Origin's COM automation is Windows-only: "
                "configure the MCP client to launch the WINDOWS executable "
                "(origin-pro-mcp.exe, or uvx.exe from WSL), not a WSL Python."
            )
        return (
            f"origin-pro-mcp only runs on Windows 10/11 (detected: {sys.platform}). "
            "It drives OriginLab Origin through Windows COM automation."
        )
    try:
        ver = sys.getwindowsversion()
    except Exception:  # noqa: BLE001 — unknown -> do not block
        return None
    if ver.major < 10:
        return (
            f"origin-pro-mcp requires Windows 10 or 11 (detected Windows NT "
            f"{ver.major}.{ver.minor}, build {ver.build})."
        )
    return None


def main():
    """Entry point for uvx / console_scripts."""
    problem = platform_problem()
    if problem:
        import sys

        print(f"[origin-pro-mcp] {problem} "
              f"(set {_SKIP_PLATFORM_CHECK_ENV}=1 to bypass)", file=sys.stderr)
        sys.exit(1)
    if use_daemon():
        from .shim import run_stdio

        run_stdio()
        return
    from . import server  # noqa: F401 — registers all tools
    mcp.run(transport="stdio")


if __name__ == "__main__":
    # Run as ``python -m origin_pro_mcp.app`` this file is ``__main__``, a
    # second copy of the module: the tools register on
    # ``origin_pro_mcp.app.mcp``, not on this copy's ``mcp``. Delegate to the
    # real module so the server that runs is the one the tools are on (this
    # matters in in-process mode, ORIGIN_PRO_MCP_USE_DAEMON=0).
    from origin_pro_mcp.app import main as _package_main

    _package_main()
