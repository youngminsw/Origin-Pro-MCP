import os
import shutil
import sys
import time

from ..app import mcp
from ..origin_connection import (
    _close_instance,
    _connection_alive,
    activate_window,
    check_project_collision,
    drop_session_origin,
    execute_labtalk,
    get_origin,
    get_remembered_project_path,
    graph_names,
    peek_session_origin,
    remember_project_path,
    require_graph,
    session_origin_is_shared,
)
from ..labtalk_safe import (
    allowed_write_roots, labtalk_choice, labtalk_name, require_writable_path,
    windows_path, writable_path,
)
from .graph import EXPORT_FORMATS, export_graph_to_file

PROJECT_EXTENSIONS = {".opj", ".opju"}

# A real .opju is kilobytes; a blank/empty project is tiny. Used to guard
# against N5: an empty project silently overwriting a real one.
_MIN_REAL_PROJECT_BYTES = 2048


def _project_page_count(o) -> int:
    """Total open windows (workbooks + graphs + matrices), or -1 if it can't be
    read (unknown -> never used to justify a destructive action)."""
    try:
        return (o.WorksheetPages.Count + o.GraphPages.Count + o.MatrixPages.Count)
    except Exception:
        return -1


def _current_project_location(o):
    """Where an in-place ``Save("")`` would write: the remembered load/save
    path, else Origin's ``%X`` (project folder). None if unknown."""
    path = get_remembered_project_path()
    if path:
        return path
    try:
        folder = o.LTStr("%X")
    except Exception:
        folder = ""
    return folder or None


def _require_current_location_writable(o, action: str) -> None:
    """With ORIGIN_PRO_MCP_ALLOWED_ROOTS set, refuse an in-place save of a
    project whose file is outside the allowed folders (e.g. an original on
    OneDrive) — or whose location cannot be determined."""
    loc = _current_project_location(o)
    if loc is None:
        if not allowed_write_roots():  # unrestricted -> nothing to check
            return
        msg = (f"Refusing to {action}: the project's file location is unknown and "
               "writes are restricted to ORIGIN_PRO_MCP_ALLOWED_ROOTS. Use "
               "save_project(file_path=...) inside an allowed folder instead.")
        raise ValueError(msg)
    require_writable_path(loc, "the project's current file")


def _backup_existing(path: str):
    """Copy an existing non-empty project file to ``<path>.bak`` before it may be
    overwritten (N5 safety net). Returns the backup path, or None."""
    try:
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            bak = path + ".bak"
            shutil.copy2(path, bak)
            return bak
    except OSError:
        pass
    return None

@mcp.tool()
def new_project() -> str:
    """Create a new empty Origin project (closes current without saving)."""
    o = get_origin()
    o.NewProject()
    remember_project_path(None)  # a fresh project has no on-disk path to recover
    return "New project created"

@mcp.tool()
def save_project(file_path: str = "") -> str:
    """Save the current Origin project.

    Args:
        file_path: Output path for the .opju file (Windows or WSL style).
                   ".opju" is appended when no extension is given.
                   If empty, saves to the project's current location.

    Returns:
        Save confirmation with path
    """
    o = get_origin()
    if file_path:
        path = writable_path(file_path, "file_path")
        ext = os.path.splitext(path)[1].lower()
        if not ext:
            path = f"{path}.opju"
        elif ext not in PROJECT_EXTENSIONS:
            msg = f"file_path must end in .opj or .opju, got '{ext}'."
            raise ValueError(msg)
        # N5 guard: NEVER let an empty project silently overwrite a real one.
        pages = _project_page_count(o)
        if (pages == 0 and os.path.isfile(path)
                and os.path.getsize(path) > _MIN_REAL_PROJECT_BYTES):
            msg = (
                f"Refusing to save: the current Origin project is EMPTY (no "
                f"worksheets/graphs/matrices) but '{path}' already holds a real "
                f"project ({os.path.getsize(path)} bytes). Saving would destroy "
                f"it. Load the project first, or save to a different path."
            )
            raise ValueError(msg)
        out_dir = os.path.dirname(path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        bak = _backup_existing(path)  # safety net before overwriting
        if not o.Save(path):
            msg = f"Origin could not save the project to {path}."
            if bak:
                msg += f" A pre-save backup was kept at {bak}."
            raise ValueError(msg)
        remember_project_path(path)
        return f"Project saved to: {path}" + (f" (backup: {bak})" if bak else "")
    # No path -> save to the project's current location.
    if _project_page_count(o) == 0:
        msg = (
            "Refusing to save: the current Origin project is empty; saving to its "
            "existing location could overwrite real data. Pass an explicit "
            "file_path if you really mean to save an empty project elsewhere."
        )
        raise ValueError(msg)
    _require_current_location_writable(o, "save in place")
    if not o.Save(""):
        msg = (
            "Origin could not save: the project has no file location yet. "
            "Pass file_path to save it somewhere."
        )
        raise ValueError(msg)
    return "Project saved"

@mcp.tool()
def close_origin(save: bool = False, confirm: bool = False) -> str:
    """Close THIS session's Origin window.

    Only the Origin instance this MCP session launched is closed — never
    another session's, and never a window the user opened themselves (attach
    mode / in-process mode) unless ``confirm=True``. Use this instead of
    ``run_labtalk("exit")``, which can reach the wrong instance after a daemon
    restart. The next tool call starts a fresh, empty Origin.

    Args:
        save: True saves the project to its current file first (fails if it
              has none — call save_project(file_path=...) first). False
              (default) DISCARDS unsaved changes without a prompt.
        confirm: Required to close an Origin the user opened themselves.

    Returns:
        What was closed and whether changes were saved or discarded.
    """
    o = peek_session_origin()
    if o is None:
        drop_session_origin()
        return "No Origin window is open for this session; nothing to close."
    if session_origin_is_shared() and not confirm:
        msg = (
            "Refusing to close: this session is using an Origin window the user "
            "opened (attach mode or in-process mode), so closing it would close "
            "THEIR Origin. Re-call with confirm=True only if the user explicitly "
            "asked to close it."
        )
        raise ValueError(msg)
    if save:
        if _project_page_count(o) == 0:
            msg = ("Refusing to save: the project is empty. Close with save=False, "
                   "or load/create content first.")
            raise ValueError(msg)
        _require_current_location_writable(o, "save before closing")
        if not o.Save(""):
            msg = ("Origin could not save: the project has no file location. Call "
                   "save_project(file_path=...) first, or close with save=False.")
            raise ValueError(msg)
        outcome = "changes saved"
    else:
        # Mark the project clean so Exit closes without a "Save changes?" modal
        # (which would block the session until someone clicks it).
        try:
            o.Execute("doc -s;")
        except Exception:  # noqa: BLE001, S110 - best-effort COM guard
            pass
        outcome = "unsaved changes discarded"
    # The PID is only trusted for an isolated instance this session launched;
    # an attached/shared Origin is never force-killed.
    pid = None if session_origin_is_shared() else _session_origin_pid()
    _close_instance(o)
    # Exit returns before the process is gone, and on Origin 2018 an Exit sent
    # to a -Embedding instance does not always take — so confirm by PID (or, if
    # the PID is unknown, by the COM proxy dying).
    closed = _wait_closed(o, pid, timeout=10.0)
    forced = False
    if not closed and pid is not None:
        try:
            _terminate_pid(pid)
        except Exception:  # noqa: BLE001, S110 - best-effort COM guard
            pass
        closed = _wait_closed(o, pid, timeout=3.0)
        forced = closed
    _forget_session_origin_pid()
    drop_session_origin()
    if not closed:
        return (f"Exit sent to this session's Origin ({outcome}), but it has not "
                "exited yet — a dialog may be open in that window.")
    if forced:
        return (f"Closed this session's Origin ({outcome}); it ignored Exit, so its "
                f"process (PID {pid}) was terminated.")
    return f"Closed this session's Origin ({outcome})."


def _daemon_pid_stores() -> list:
    """The daemon's thread-local PID store(s).

    The daemon runs as ``python -m origin_pro_mcp.daemon``, so inside that
    process its module is ``__main__``; ``import origin_pro_mcp.daemon`` would
    load a SECOND, empty copy. Check both, never importing a new copy.
    """
    stores = []
    for name in ("__main__", "origin_pro_mcp.daemon"):
        mod = sys.modules.get(name)
        tls = getattr(mod, "_real_pid_tls", None) if mod is not None else None
        if tls is not None and all(tls is not s for s in stores):
            stores.append(tls)
    return stores


def _session_origin_pid():
    """PID of the Origin this daemon worker launched (captured by the daemon's
    factory on this thread), or None (in-process mode / capture failed)."""
    for tls in _daemon_pid_stores():
        pid = getattr(tls, "pid", None)
        if isinstance(pid, int) and pid > 0:
            return pid
    return None


def _forget_session_origin_pid() -> None:
    """Clear the captured PID so a later force-kill can never hit a reused PID."""
    for tls in _daemon_pid_stores():
        try:
            tls.pid = None
        except Exception:  # noqa: BLE001, S110 - best-effort COM guard
            pass


def _terminate_pid(pid: int) -> None:
    for name in ("__main__", "origin_pro_mcp.daemon"):
        fn = getattr(sys.modules.get(name), "default_terminate_process", None)
        if callable(fn):
            fn(pid)
            return


def _pid_alive(pid: int) -> bool:
    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    import ctypes

    SYNCHRONIZE = 0x00100000
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    WAIT_TIMEOUT = 0x102
    k32 = ctypes.windll.kernel32
    handle = k32.OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return False
    try:
        return k32.WaitForSingleObject(handle, 0) == WAIT_TIMEOUT
    finally:
        k32.CloseHandle(handle)


def _wait_closed(o, pid, timeout: float) -> bool:
    deadline = time.monotonic() + timeout
    while True:
        alive = _pid_alive(pid) if pid is not None else _connection_alive(o)
        if not alive:
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.25)

@mcp.tool()
def load_project(file_path: str) -> str:
    """Open an Origin project file. Replaces the current project.

    Args:
        file_path: Path to a .opj or .opju file (Windows or WSL style)

    Returns:
        Success message
    """
    o = get_origin()
    path = windows_path(file_path, "file_path")
    # Check before calling Load: a failed Load can still discard the
    # currently open project, so a typo must not reach Origin.
    if not os.path.isfile(path):
        msg = f"Project file not found: {path}"
        raise ValueError(msg)
    # Keep one pre-load backup of the target (N5 safety net). o.Load is
    # read-only, but a backup guarantees the file is recoverable no matter what
    # a later save/autosave does in a crash-recovery sequence.
    _backup_existing(path)
    if not o.Load(path):
        msg = f"Origin could not open the project file: {path}"
        raise ValueError(msg)
    # Verify the load actually took: Origin sometimes returns success while the
    # attached instance stays empty (the "success but list_worksheets empty"
    # case). Surface it as a failure so the caller can retry, and don't
    # remember an empty session as the recovery target.
    try:
        pages = (o.WorksheetPages.Count + o.GraphPages.Count
                 + o.MatrixPages.Count)
    except Exception:
        pages = -1  # can't introspect (odd COM build) -> don't false-fail
    if pages == 0:
        msg = (
            f"Origin reported success but the loaded project is empty "
            f"(no worksheets/graphs/matrices). This is the flaky empty-load; "
            f"retry load_project('{file_path}')."
        )
        raise ValueError(msg)
    remember_project_path(path)
    message = f"Loaded project: {path}"
    # Warn (never block) if another live Origin still holds this same project —
    # saving from both would clobber. Pure string append from the session ledger.
    warning = check_project_collision(path)
    if warning:
        message += "\n" + warning
    return message

@mcp.tool()
def export_all_graphs(
    output_dir: str,
    format: str = "png",
    width: int = 0,
) -> str:
    """Export every graph in the project to image files (one per graph).

    Uses the same clipboard-free export as export_graph (Origin's expGraph
    X-Function), so the user's clipboard is preserved.

    Args:
        output_dir: Output directory (Windows or WSL style). Created if missing.
        format: Output format. Raster: png, jpg, tif, bmp. Vector: pdf, eps, emf
        width: Output raster pixel WIDTH. 0 (default) = auto ~1200px wide. Any
            value > 0 is applied to every exported graph (height follows each
            graph's aspect ratio; Origin's expGraph cannot set an independent
            height). Ignored for vector formats (resolution-independent).

    Returns:
        Per-graph list of exported files
    """
    from .graph import _export_via_expgraph

    safe_format = labtalk_choice(format.lower(), EXPORT_FORMATS, "format")
    out_dir = writable_path(output_dir, "output_dir")
    os.makedirs(out_dir, exist_ok=True)

    names = graph_names()
    if not names:
        return "No graphs found in the current project."

    exported = []
    failed = []
    for name in names:
        path = os.path.join(out_dir, f"{name}.{safe_format}")
        try:
            if width > 0:
                exported.append(_export_via_expgraph(
                    name, path, width=width, format=safe_format
                ))
            else:
                exported.append(export_graph_to_file(name, path, format=safe_format))
        except (ValueError, RuntimeError) as exc:
            failed.append(f"{name}: {exc}")

    lines = [f"Exported {len(exported)} of {len(names)} graphs to {out_dir}:"]
    lines.extend(f"  {p}" for p in exported)
    if failed:
        lines.append("Failed:")
        lines.extend(f"  {f}" for f in failed)
    return "\n".join(lines)


@mcp.tool()
def save_graph_template(
    graph_name: str,
    file_path: str,
    old_format: bool = False
) -> str:
    """Save a graph as a reusable Origin template (.otpu/.otp).

    Args:
        graph_name: Graph to save as a template
        file_path: Output path (Windows or WSL style). The extension is
                   forced to .otpu (or .otp when old_format=True).
        old_format: Save the legacy .otp format (loadable in Origin 2017
                    and earlier) instead of the modern .otpu

    Returns:
        Path to the saved template
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    path = writable_path(file_path, "file_path")
    ext = ".otp" if old_format else ".otpu"
    root = os.path.splitext(path)[0]
    path = root + ext
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    flag = "-tj" if old_format else "-t"
    if not execute_labtalk(f'save {flag} {safe_graph} "{path}";'):
        msg = f"Origin could not save the template for {safe_graph}."
        raise ValueError(msg)
    if not os.path.exists(path):
        msg = f"Template save reported success but {path} was not created."
        raise ValueError(msg)
    return f"Saved template: {path} ({os.path.getsize(path)} bytes)"
