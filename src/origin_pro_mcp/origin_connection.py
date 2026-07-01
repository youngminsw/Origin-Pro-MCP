import threading
import time

from .labtalk_safe import labtalk_name

# Per-thread COM proxy store. Each thread (e.g. a daemon session worker) holds
# its own Origin instance, so the cached proxy never leaks across threads.
_state = threading.local()
_MAINWND_SHOW = 1

def _connection_alive(origin) -> bool:
    """Ping the cached COM proxy — stale proxies raise after Origin closes.

    Retries once on a transient error so a momentarily-busy (but still alive)
    Origin isn't mistaken for dead — a false negative would spawn a redundant
    instance and orphan the live one."""
    for attempt in range(2):
        try:
            origin.Visible
            return True
        except AttributeError:
            return True  # non-COM stand-in (tests) without a Visible property
        except Exception:
            if attempt == 0:
                time.sleep(0.25)
                continue
            return False
    return False


def _close_instance(inst) -> None:
    """Best-effort close of an Origin instance (COM). Used before relaunching a
    dead/stale proxy so a still-alive process isn't left orphaned."""
    for closer in ("Exit", "Close"):
        fn = getattr(inst, closer, None)
        if callable(fn):
            try:
                fn()
            except Exception:
                pass
            break


def set_session_origin(origin, factory=None) -> None:
    """Bind a COM proxy to the current thread (daemon worker / test injection).

    ``factory`` (daemon sessions) is the zero-arg callable that launched this
    session's OWN isolated instance. If the proxy later dies (e.g. the user
    closes the window), ``get_origin`` re-runs it to relaunch a fresh isolated
    instance — instead of falling back to the shared ``ApplicationSI``, which
    could silently attach the agent to a DIFFERENT or user-opened Origin and
    modify the user's work.
    """
    _state.origin = origin
    if factory is not None:
        _state.factory = factory


def clear_session_origin() -> None:
    """Drop the current thread's proxy and factory (full session teardown)."""
    for attr in ("origin", "factory", "project_path"):
        if hasattr(_state, attr):
            delattr(_state, attr)


def remember_project_path(path) -> None:
    """Record the last project this thread loaded/saved, so a relaunch after a
    crash (dead COM proxy) can auto-reopen it. Falsy value forgets it (e.g.
    after ``new_project``)."""
    if path:
        _state.project_path = path
    elif hasattr(_state, "project_path"):
        del _state.project_path


def get_remembered_project_path():
    return getattr(_state, "project_path", None)


def _reopen_remembered_project(origin) -> None:
    """Best-effort: reopen the remembered on-disk project into a freshly
    relaunched instance. A crash/close still drops unsaved edits, but the last
    saved project comes back instead of an empty session. Fully guarded."""
    import os

    path = getattr(_state, "project_path", None)
    if not path or not os.path.isfile(path):
        return
    try:
        origin.Load(path)
    except Exception:
        pass


def get_origin():
    origin = getattr(_state, "origin", None)
    if origin is not None and not _connection_alive(origin):
        # Close the stale instance first: if the process is alive-but-wedged,
        # this stops the relaunch below from orphaning it. (A truly-dead proxy
        # just no-ops here.)
        _close_instance(origin)
        origin = None  # Origin was closed or restarted; reconnect below
        if hasattr(_state, "origin"):
            del _state.origin  # keep the factory so a daemon session can relaunch
    if origin is None:
        factory = getattr(_state, "factory", None)
        if factory is not None:
            # Daemon session: relaunch THIS session's own isolated instance.
            # Never fall through to ApplicationSI — that could hijack the user's
            # open Origin or another session's instance.
            origin = factory()
            _state.origin = origin
            # Recovery: if this thread had a project loaded before the proxy
            # died, reopen it into the fresh instance (best-effort). On the
            # first launch nothing is remembered yet, so this is a no-op.
            _reopen_remembered_project(origin)
            return origin
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
            origin = win32com.client.Dispatch("Origin.ApplicationSI")
        except pywintypes.com_error as exc:
            msg = (
                "Could not connect to Origin via COM (Origin.ApplicationSI). "
                "Check that Origin/OriginPro is installed and licensed on this "
                "machine. If it is, run Origin once as administrator to "
                "re-register its Automation Server."
            )
            raise RuntimeError(msg) from exc
        try:
            origin.Visible = _MAINWND_SHOW
        except (AttributeError, pywintypes.com_error):
            pass
        set_session_origin(origin)
    return origin

def execute_labtalk(script: str) -> bool:
    o = get_origin()
    return o.Execute(script)

def get_lt_var(name: str) -> float:
    return get_origin().LTVar(name)

def get_lt_str(name: str) -> str:
    return get_origin().LTStr(name)

# ASCII record separator — cannot appear in an Origin window/sheet name, so it
# is a safe delimiter for the LabTalk-built name list.
_ENUM_DELIM = "\x1e"


def safe_page_names(pages) -> list:
    """Top-level names of a COM page collection, isolating a bad/corrupt entry
    so one unreadable window can't abort the whole enumeration (which, on a
    heavy project, is what can wedge/crash the COM bridge)."""
    names: list = []
    try:
        count = pages.Count
    except Exception:
        return names
    for i in range(count):
        try:
            names.append(pages.Item(i).Name)
        except Exception:
            continue
    return names


def sheet_names(book_name: str) -> list:
    """Sheet (layer) names of a workbook via LabTalk's internal ``layer$(k)``
    loop — never the deep ``page.Layers.Item(j).Name`` COM traversal, which
    instantiates a proxy per sheet and can HARD-CRASH Origin on heavy projects.
    Crash-safe (verified on Origin 2020). Returns [] on any failure."""
    o = get_origin()
    try:
        o.Execute(
            f'string _opm_sh$="";win -a {book_name};'
            f'for(int _opmk=1;_opmk<=page.nlayers;_opmk++)'
            f'{{_opm_sh$=_opm_sh$+layer$(_opmk).name$+"{_ENUM_DELIM}";}}'
        )
        raw = o.LTStr("_opm_sh$") or ""
        return [s for s in raw.split(_ENUM_DELIM) if s]
    except Exception:
        return []


def workbook_names() -> list:
    """Names of all open workbooks (per-item isolated)."""
    return safe_page_names(get_origin().WorksheetPages)


def graph_names() -> list:
    """Names of all open graph windows (per-item isolated)."""
    return safe_page_names(get_origin().GraphPages)


def matrix_names() -> list:
    """Names of all open matrix books (per-item isolated)."""
    return safe_page_names(get_origin().MatrixPages)

def require_matrix(book: str, sheet: str = "MSheet1") -> str:
    """Return the [book]sheet matrix reference, or raise with open matrices."""
    target = f"[{book}]{sheet}"
    if get_origin().FindMatrixSheet(target) is None:
        mats = ", ".join(matrix_names()) or "(none)"
        msg = f"Matrix {target} not found. Open matrices: {mats}."
        raise ValueError(msg)
    return target

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
