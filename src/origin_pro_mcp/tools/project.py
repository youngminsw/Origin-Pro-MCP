import os
import shutil

from ..app import mcp
from ..origin_connection import (
    activate_window,
    check_project_collision,
    execute_labtalk,
    get_origin,
    graph_names,
    remember_project_path,
    require_graph,
)
from ..labtalk_safe import labtalk_choice, labtalk_name, windows_path
from .graph import EXPORT_IMAGE_FORMATS, export_graph_to_file

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
        path = windows_path(file_path, "file_path")
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
    if not o.Save(""):
        msg = (
            "Origin could not save: the project has no file location yet. "
            "Pass file_path to save it somewhere."
        )
        raise ValueError(msg)
    return "Project saved"

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
    dpi: int = 300,
    width: int = 800,
    height: int = 600
) -> str:
    """Export every graph in the project to image files (one per graph).

    Uses the same clipboard-free export as export_graph (Origin's expGraph
    X-Function), so the user's clipboard is preserved. Each graph is exported
    at ~1200px wide with the aspect ratio kept.

    Args:
        output_dir: Output directory (Windows or WSL style). Created if missing.
        format: Image format: png, jpg, tif, bmp
        dpi: Unused (kept for API compatibility; ~1200px wide is used)
        width: Unused (kept for API compatibility)
        height: Unused (kept for API compatibility)

    Returns:
        Per-graph list of exported files
    """
    safe_format = labtalk_choice(format.lower(), EXPORT_IMAGE_FORMATS, "format")
    out_dir = windows_path(output_dir, "output_dir")
    os.makedirs(out_dir, exist_ok=True)

    names = graph_names()
    if not names:
        return "No graphs found in the current project."

    exported = []
    failed = []
    for name in names:
        path = os.path.join(out_dir, f"{name}.{safe_format}")
        try:
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
    path = windows_path(file_path, "file_path")
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
