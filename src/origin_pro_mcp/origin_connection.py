from .labtalk_safe import labtalk_name

_origin = None
_MAINWND_SHOW = 1

def _connection_alive(origin) -> bool:
    """Ping the cached COM proxy — stale proxies raise after Origin closes."""
    try:
        origin.Visible
    except AttributeError:
        return True  # non-COM stand-in (tests) without a Visible property
    except Exception:
        return False
    return True


def get_origin():
    global _origin
    if _origin is not None and not _connection_alive(_origin):
        _origin = None  # Origin was closed or restarted; reconnect below
    if _origin is None:
        try:
            import win32com.client
            import pywintypes
        except ModuleNotFoundError as exc:
            msg = (
                "Origin Pro COM automation requires Windows Python with pywin32. "
                "Run this MCP server on Windows, not WSL/Linux."
            )
            raise RuntimeError(msg) from exc
        try:
            _origin = win32com.client.Dispatch("Origin.ApplicationSI")
        except pywintypes.com_error as exc:
            msg = (
                "Could not connect to Origin via COM (Origin.ApplicationSI). "
                "Check that Origin/OriginPro is installed and licensed on this "
                "machine. If it is, run Origin once as administrator to "
                "re-register its Automation Server."
            )
            raise RuntimeError(msg) from exc
        try:
            _origin.Visible = _MAINWND_SHOW
        except (AttributeError, pywintypes.com_error):
            pass
    return _origin

def execute_labtalk(script: str) -> bool:
    o = get_origin()
    return o.Execute(script)

def get_lt_var(name: str) -> float:
    return get_origin().LTVar(name)

def get_lt_str(name: str) -> str:
    return get_origin().LTStr(name)

def workbook_names() -> list:
    """Names of all open workbooks."""
    pages = get_origin().WorksheetPages
    return [pages.Item(i).Name for i in range(pages.Count)]

def graph_names() -> list:
    """Names of all open graph windows."""
    pages = get_origin().GraphPages
    return [pages.Item(i).Name for i in range(pages.Count)]

def require_worksheet(book: str, sheet: str) -> str:
    """Return the [book]sheet reference, or raise with the open workbooks."""
    target = f"[{book}]{sheet}"
    if get_origin().FindWorksheet(target) is None:
        books = ", ".join(workbook_names()) or "(none)"
        msg = f"Worksheet {target} not found. Open workbooks: {books}."
        raise ValueError(msg)
    return target

def require_graph(graph_name: str):
    """Return the graph's Layer1, or raise with the list of open graphs."""
    layer = get_origin().FindGraphLayer(f"[{graph_name}]Layer1")
    if layer is None:
        graphs = ", ".join(graph_names()) or "(none)"
        msg = f"Graph '{graph_name}' not found. Open graphs: {graphs}."
        raise ValueError(msg)
    return layer

def activate_window(name: str, field: str = "window") -> None:
    """Activate a window by name, or raise with the list of open windows.

    Origin's `win -a` silently returns False for unknown windows, and any
    follow-up command would then hit whatever window happens to be active —
    so every tool must go through this guard before window-scoped commands.
    """
    safe_name = labtalk_name(name, field)
    if not execute_labtalk(f"win -a {safe_name};"):
        books = ", ".join(workbook_names()) or "(none)"
        graphs = ", ".join(graph_names()) or "(none)"
        msg = (
            f"Window '{safe_name}' not found. "
            f"Open workbooks: {books}. Open graphs: {graphs}."
        )
        raise ValueError(msg)
