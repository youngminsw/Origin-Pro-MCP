import os
import time

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_origin,
    require_graph,
    require_worksheet,
)
from ..labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_string,
    positive_column,
    windows_path,
)

PLOT_TYPES = {
    "scatter": 201,
    "line": 200,
    "line+symbol": 202,
    "column": 203,
    "bar": 204,
    "area": 205,
    "histogram": 206,
    "box": 207,
    "contour": 208,
    "3d_scatter": 209,
    "3d_surface": 210,
    "pie": 212,
    "bubble": 228,
}

# Clipboard-based export (CopyPage) only produces raster images
EXPORT_IMAGE_FORMATS = {"png", "jpg", "tif", "bmp"}


def _designate_error_column(book: str, sheet: str, col: int) -> None:
    """Mark a column as Y Error so styling tools can recognize error-bar
    plots later. Must use the active-sheet `wks` form — the sheet-qualified
    `[Book]Sheet!col(n).type = ...` is silently ignored on Origin 2020."""
    execute_labtalk(f"win -a {book};")
    execute_labtalk(f'page.active$ = "{sheet}";')
    execute_labtalk(f"wks.col{col}.type = 3;")


def export_graph_to_file(graph_name: str, file_path: str) -> str:
    """Export one graph via CopyPage + clipboard. Shared by export tools.

    Returns the output path. Raises ValueError with a friendly message on
    failure. Note: this overwrites the Windows clipboard contents.
    """
    try:
        from PIL import ImageGrab
    except ImportError as exc:
        msg = "Pillow (PIL) is required for export. Install with: pip install Pillow"
        raise RuntimeError(msg) from exc

    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    activate_window(safe_graph_name, "graph_name")

    out_dir = os.path.dirname(file_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # CopyPage: format=4 (BMP to clipboard), then save via Pillow.
    # expGraph does not produce files over COM on Origin 2020.
    if not o.CopyPage(safe_graph_name, 4, 96, 24):
        msg = f"Origin could not copy graph '{safe_graph_name}' to the clipboard."
        raise ValueError(msg)

    # Clipboard write is asynchronous — poll instead of a fixed sleep.
    img = None
    deadline = time.monotonic() + 5.0
    while time.monotonic() < deadline:
        time.sleep(0.2)
        img = ImageGrab.grabclipboard()
        if img is not None:
            break
    if img is None:
        msg = (
            f"Export failed: no image appeared on the clipboard for "
            f"'{safe_graph_name}'. Check that the Origin window is not "
            "minimized and that no other app is using the clipboard."
        )
        raise ValueError(msg)

    try:
        img.save(file_path)
    except (ValueError, OSError) as exc:
        msg = f"Could not save image to {file_path}: {exc}"
        raise ValueError(msg) from exc
    if not os.path.exists(file_path):
        msg = f"Export failed: {file_path} was not created."
        raise ValueError(msg)
    return file_path


@mcp.tool()
def create_graph(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    plot_type: str = "scatter",
    y_error_col: int = 0,
    title: str = ""
) -> str:
    """Create a graph from worksheet data.

    Args:
        graph_name: Name for the graph window
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column number (1-based)
        y_col: Y column number (1-based)
        plot_type: scatter, line, line+symbol, column, bar, area, histogram,
                   box, contour, pie, bubble
        y_error_col: Optional Y error column number (1-based, 0=none)
        title: Optional graph title

    Returns:
        Created graph name (may differ from graph_name if it was taken)
    """
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_plot_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    ptype = PLOT_TYPES[safe_plot_type]
    require_worksheet(safe_book, safe_sheet)

    data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
    if y_error_col > 0:
        safe_error_col = positive_column(y_error_col, "y_error_col")
        _designate_error_column(safe_book, safe_sheet, safe_error_col)
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_error_col})"

    name = o.CreatePage(3, safe_graph_name, "origin")

    if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{name}]Layer1;"):
        # Remove the empty graph page we just created
        execute_labtalk(f"win -cd {name};")
        msg = (
            f"Could not plot {data_ref}. Check that columns "
            f"{safe_x_col} and {safe_y_col} exist and contain data, and that "
            f"plot_type '{safe_plot_type}' suits this data."
        )
        raise ValueError(msg)

    if title:
        execute_labtalk(f"label -n title -s {labtalk_string(title, 'title')}; title.x = 50; title.y = 95;")

    return f"Created graph: {name} ({safe_plot_type})"

@mcp.tool()
def add_plot_to_graph(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    plot_type: str = "scatter",
    y_error_col: int = 0
) -> str:
    """Add another data series to an existing graph.

    Args:
        graph_name: Existing graph name
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column number (1-based)
        y_col: Y column number (1-based)
        plot_type: Plot type (scatter, line, line+symbol, etc.)
        y_error_col: Optional Y error column (1-based, 0=none)

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_plot_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    ptype = PLOT_TYPES[safe_plot_type]
    require_graph(safe_graph_name)
    require_worksheet(safe_book, safe_sheet)
    data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
    if y_error_col > 0:
        safe_error_col = positive_column(y_error_col, "y_error_col")
        _designate_error_column(safe_book, safe_sheet, safe_error_col)
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_error_col})"

    if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{safe_graph_name}]Layer1;"):
        msg = (
            f"Could not add plot {data_ref} to {safe_graph_name}. Check that "
            f"columns {safe_x_col} and {safe_y_col} exist and contain data."
        )
        raise ValueError(msg)
    return f"Added {safe_plot_type} plot to {safe_graph_name}"

@mcp.tool()
def set_axis_labels(
    graph_name: str,
    x_label: str = "",
    y_label: str = "",
    title: str = ""
) -> str:
    """Set axis labels and title for a graph.

    Args:
        graph_name: Graph name
        x_label: X axis label
        y_label: Y axis label
        title: Graph title

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    activate_window(safe_graph_name, "graph_name")
    if x_label:
        execute_labtalk(f"xb.text$ = {labtalk_string(x_label, 'x_label')};")
    if y_label:
        execute_labtalk(f"yl.text$ = {labtalk_string(y_label, 'y_label')};")
    if title:
        execute_labtalk(f"label -n title -s {labtalk_string(title, 'title')}; title.x = 50; title.y = 95;")
    return f"Updated labels for {safe_graph_name}"

@mcp.tool()
def set_axis_range(
    graph_name: str,
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None
) -> str:
    """Set axis range for a graph.

    Args:
        graph_name: Graph name
        x_min: X axis minimum (None=auto)
        x_max: X axis maximum (None=auto)
        y_min: Y axis minimum (None=auto)
        y_max: Y axis maximum (None=auto)

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    activate_window(safe_graph_name, "graph_name")
    if x_min is not None:
        execute_labtalk(f"layer.x.from = {x_min};")
    if x_max is not None:
        execute_labtalk(f"layer.x.to = {x_max};")
    if y_min is not None:
        execute_labtalk(f"layer.y.from = {y_min};")
    if y_max is not None:
        execute_labtalk(f"layer.y.to = {y_max};")
    return f"Set axis range for {safe_graph_name}"

@mcp.tool()
def export_graph(
    graph_name: str,
    file_path: str,
    format: str = "png",
    width: int = 600,
    height: int = 400,
    dpi: int = 300
) -> str:
    """Export a graph to an image file.

    Uses Origin's CopyPage + clipboard, so the Windows clipboard contents
    are replaced during export. Image size is determined by the Origin
    page (width/height/dpi are accepted but have no effect).

    Args:
        graph_name: Graph name to export
        file_path: Output path (Windows or WSL style, e.g.
                   C:\\Users\\me\\fig1.png or /mnt/c/Users/me/fig1.png).
                   Missing directories are created.
        format: Image format: png, jpg, tif, bmp. Used as the file
                extension when file_path has none.
        width: Unused (kept for API compatibility; size determined by Origin page)
        height: Unused (kept for API compatibility)
        dpi: Unused (kept for API compatibility)

    Returns:
        Path to exported file
    """
    safe_format = labtalk_choice(format.lower(), EXPORT_IMAGE_FORMATS, "format")
    path = windows_path(file_path, "file_path")
    if not os.path.splitext(path)[1]:
        path = f"{path}.{safe_format}"

    out = export_graph_to_file(graph_name, path)
    size = os.path.getsize(out)
    return f"Exported to: {out} ({size} bytes)"
