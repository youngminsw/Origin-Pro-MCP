import json
from ..app import mcp
from ..origin_connection import get_origin, execute_labtalk
from ..labtalk_safe import labtalk_choice, labtalk_name, labtalk_string, positive_column

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
        Created graph name
    """
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_plot_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    ptype = PLOT_TYPES[safe_plot_type]
    name = o.CreatePage(3, safe_graph_name, "origin")

    data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
    if y_error_col > 0:
        safe_error_col = positive_column(y_error_col, "y_error_col")
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_error_col})"

    execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{name}]Layer1;")

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
    data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
    if y_error_col > 0:
        safe_error_col = positive_column(y_error_col, "y_error_col")
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_error_col})"

    execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{safe_graph_name}]Layer1;")
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
    execute_labtalk(f"win -a {safe_graph_name};")
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
    x_min: float = None,
    x_max: float = None,
    y_min: float = None,
    y_max: float = None
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
    execute_labtalk(f"win -a {safe_graph_name};")
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

    Args:
        graph_name: Graph name to export
        file_path: Full Windows output path (e.g., C:\\Users\\fig1.png)
        format: Image format: png, jpg, tif, bmp (clipboard-based export)
        width: Unused (kept for API compatibility; size determined by Origin page)
        height: Unused (kept for API compatibility)
        dpi: Unused (kept for API compatibility)

    Returns:
        Path to exported file
    """
    try:
        from PIL import ImageGrab
    except ImportError:
        return "Export failed: Pillow (PIL) is required. Install with: pip install Pillow"

    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    execute_labtalk(f"win -a {safe_graph_name};")

    # CopyPage: format=4 (BMP to clipboard), then save via Pillow
    try:
        o.CopyPage(safe_graph_name, 4, 96, 24)
    except Exception as e:
        return f"CopyPage failed: {e}"

    import time
    time.sleep(0.5)  # wait for clipboard

    img = ImageGrab.grabclipboard()
    if img is None:
        return f"Export failed: could not grab clipboard image for {safe_graph_name}"

    img.save(file_path)
    return f"Exported to: {file_path}"
