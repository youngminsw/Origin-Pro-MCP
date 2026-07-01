import os

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_origin,
    graph_names,
    remember_project_path,
    require_graph,
)
from ..labtalk_safe import labtalk_choice, labtalk_name, windows_path
from .graph import EXPORT_IMAGE_FORMATS, export_graph_to_file

PROJECT_EXTENSIONS = {".opj", ".opju"}

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
        out_dir = os.path.dirname(path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        if not o.Save(path):
            msg = f"Origin could not save the project to {path}."
            raise ValueError(msg)
        remember_project_path(path)
        return f"Project saved to: {path}"
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
    return f"Loaded project: {path}"

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
