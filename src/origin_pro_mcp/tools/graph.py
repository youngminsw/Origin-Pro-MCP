import os
import time
import uuid

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_origin,
    graph_names,
    require_graph,
    require_worksheet,
)
from ..labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_string,
    positive_column,
    positive_int,
    windows_path,
)

# Plot type IDs verified against OriginLab's "Plot Type IDs" reference and
# tested live on Origin Pro 2020. The earlier values for area/bar/box/
# histogram/pie/contour/3d were wrong (off-by-template), producing empty
# or mis-typed graphs.
PLOT_TYPES = {
    "scatter": 201,
    "line": 200,
    "line+symbol": 202,
    "column": 203,
    "bar": 215,          # horizontal bar (204 is Area)
    "area": 204,
    "pie": 225,
    "histogram": 219,    # Y range
    "box": 206,          # Y range
    "contour": 243,      # XYZ range
    "3d_scatter": 240,   # XYZ range (OpenGL, owns its graph)
}

# Take a single Y column instead of an X,Y pair.
_Y_ONLY_TYPES = {"histogram", "box"}
# Need X, Y, Z columns and are drawn with plotxyz.
_XYZ_TYPES = {"contour", "3d_scatter"}
# OpenGL 3D types that must own their graph window — plotting them into a
# pre-made 2D page collapses them to an empty/flat projection.
_OWN_GRAPH_TYPES = {"3d_scatter"}

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
    z_col: int = 0,
    title: str = ""
) -> str:
    """Create a graph from worksheet data.

    Args:
        graph_name: Name for the graph window
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column number (1-based). Ignored for box/histogram.
        y_col: Y column number (1-based)
        plot_type: scatter, line, line+symbol, column, bar, area, pie,
                   box, histogram, contour, 3d_scatter
        y_error_col: Optional Y error column (1-based, 0=none). XY plots only.
        z_col: Z column (1-based). REQUIRED for contour and 3d_scatter.
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

    def _set_title():
        if title:
            execute_labtalk(
                f"label -n title -s {labtalk_string(title, 'title')}; "
                "title.x = 50; title.y = 95;"
            )

    # --- XYZ plots: contour (2D) and 3d_scatter (OpenGL), via plotxyz ---
    if safe_plot_type in _XYZ_TYPES:
        if z_col < 1:
            msg = f"plot_type '{safe_plot_type}' requires z_col (1-based Z column)."
            raise ValueError(msg)
        safe_z_col = positive_column(z_col, "z_col")
        xyz = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_z_col})"
        if safe_plot_type in _OWN_GRAPH_TYPES:
            # Activate the source sheet (a non-graph window) so plotxyz
            # creates a fresh OpenGL 3D graph instead of merging into an
            # active 2D graph and collapsing to a flat projection.
            execute_labtalk(f"win -a {safe_book};")
            before = set(graph_names())
            if not execute_labtalk(f"plotxyz iz:={xyz} plot:={ptype};"):
                msg = (
                    f"Could not create a 3D plot from {xyz}. Check that the "
                    "columns contain numeric data."
                )
                raise ValueError(msg)
            new = set(graph_names()) - before
            name = new.pop() if new else o.LTStr("page.name$")
            if name != safe_graph_name and execute_labtalk(
                f"win -r {name} {safe_graph_name};"
            ):
                name = safe_graph_name
        else:
            name = o.CreatePage(3, safe_graph_name, "origin")
            if not execute_labtalk(
                f"plotxyz iz:={xyz} plot:={ptype} ogl:=[{name}]Layer1;"
            ):
                execute_labtalk(f"win -cd {name};")
                msg = (
                    f"Could not plot {xyz}. Check that columns "
                    f"{safe_x_col}, {safe_y_col}, {safe_z_col} contain data."
                )
                raise ValueError(msg)
        _set_title()
        return f"Created graph: {name} ({safe_plot_type})"

    # --- 2D plots (plotxy) ---
    if safe_plot_type in _Y_ONLY_TYPES:
        # Box/histogram need the source column designated as Y and the
        # sheet active, or Origin draws an empty layer.
        activate_window(safe_book, "data_book")
        execute_labtalk(f'page.active$ = "{safe_sheet}"; wks.col{safe_y_col}.type = 1;')
        data_ref = f"[{safe_book}]{safe_sheet}!col({safe_y_col})"
    else:
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
        if y_error_col > 0:
            safe_error_col = positive_column(y_error_col, "y_error_col")
            _designate_error_column(safe_book, safe_sheet, safe_error_col)
            data_ref = (
                f"[{safe_book}]{safe_sheet}!"
                f"({safe_x_col},{safe_y_col},{safe_error_col})"
            )

    name = o.CreatePage(3, safe_graph_name, "origin")
    if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{name}]Layer1;"):
        execute_labtalk(f"win -cd {name};")
        msg = (
            f"Could not plot {data_ref}. Check that columns exist and contain "
            f"data, and that plot_type '{safe_plot_type}' suits this data."
        )
        raise ValueError(msg)

    _set_title()
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
    if safe_plot_type in _Y_ONLY_TYPES or safe_plot_type in _XYZ_TYPES:
        msg = (
            f"add_plot_to_graph supports only X,Y plot types; "
            f"'{safe_plot_type}' must be created with create_graph."
        )
        raise ValueError(msg)
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


# Origin axis scale type codes (layer.x/y.type).
_AXIS_SCALES = {"linear": 1, "log10": 2, "ln": 8, "log2": 9}


@mcp.tool()
def set_axis_scale(graph_name: str, axis: str = "y", scale: str = "log10") -> str:
    """Set an axis to linear or a logarithmic scale.

    Args:
        graph_name: Graph name
        axis: "x" or "y"
        scale: linear, log10, ln, or log2

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_axis = labtalk_choice(axis.lower(), {"x", "y"}, "axis")
    safe_scale = labtalk_choice(scale.lower(), _AXIS_SCALES, "scale")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    if not execute_labtalk(f"layer.{safe_axis}.type = {_AXIS_SCALES[safe_scale]};"):
        msg = f"Could not set {safe_axis} axis of {safe_graph} to {safe_scale}."
        raise ValueError(msg)
    return f"Set {safe_axis} axis of {safe_graph} to {safe_scale} scale"


@mcp.tool()
def add_second_y_axis(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    plot_type: str = "line+symbol"
) -> str:
    """Add a right-side Y axis layer and plot a second dataset on it.

    Args:
        graph_name: Existing graph
        data_book, data_sheet: Source data
        x_col, y_col: Columns for the right-axis series (1-based)
        plot_type: scatter, line, line+symbol, column, bar, area

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    safe_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    if safe_type in _Y_ONLY_TYPES or safe_type in _XYZ_TYPES:
        msg = f"add_second_y_axis supports only X,Y plot types, not '{safe_type}'."
        raise ValueError(msg)
    ptype = PLOT_TYPES[safe_type]
    require_graph(safe_graph)
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_graph, "graph_name")
    if not execute_labtalk("layadd type:=righty;"):
        msg = f"Could not add a right-Y layer to {safe_graph}."
        raise ValueError(msg)
    layer = int(get_origin().LTVar("page.nlayers"))
    data_ref = f"[{safe_book}]{safe_sheet}!({sx},{sy})"
    if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{safe_graph}]{layer}!;"):
        msg = f"Added the right-Y layer but could not plot {data_ref} on it."
        raise ValueError(msg)
    return f"Added right-Y axis (layer {layer}) to {safe_graph} with {safe_type} plot"


@mcp.tool()
def add_layer(graph_name: str, layer_type: str = "right-y") -> str:
    """Add a new layer (panel/axis) to a graph.

    Args:
        graph_name: Graph name
        layer_type: right-y, top-x, inset, or independent

    Returns:
        Success message naming the new layer index
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    type_map = {
        "right-y": "righty",
        "top-x": "topx",
        "inset": "inset",
        "independent": "independent",
    }
    safe_layer_type = labtalk_choice(layer_type.lower(), type_map, "layer_type")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    if not execute_labtalk(f"layadd type:={type_map[safe_layer_type]};"):
        msg = f"Could not add a {safe_layer_type} layer to {safe_graph}."
        raise ValueError(msg)
    layer = int(get_origin().LTVar("page.nlayers"))
    return f"Added {safe_layer_type} layer (layer {layer}) to {safe_graph}"


@mcp.tool()
def add_reference_line(
    graph_name: str,
    orientation: str,
    value: float
) -> str:
    """Draw a horizontal or vertical reference line at a data value.

    Args:
        graph_name: Graph name
        orientation: "horizontal" (constant Y) or "vertical" (constant X)
        value: Axis value where the line is drawn

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_orientation = labtalk_choice(
        orientation.lower(), {"horizontal", "vertical"}, "orientation"
    )
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    flag = "-h" if safe_orientation == "horizontal" else "-v"
    if not execute_labtalk(f"draw -l {flag} {float(value)};"):
        msg = f"Could not draw a {safe_orientation} reference line on {safe_graph}."
        raise ValueError(msg)
    return f"Drew {safe_orientation} reference line at {value} on {safe_graph}"


@mcp.tool()
def add_text_annotation(
    graph_name: str,
    text: str,
    x: float,
    y: float,
    name: str = "anno"
) -> str:
    """Add a text label to a graph at data coordinates.

    Args:
        graph_name: Graph name
        text: Annotation text (no quotes, line breaks, or ';')
        x: X position in data coordinates
        y: Y position in data coordinates
        name: Internal object name (letters/numbers/underscore)

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_name = labtalk_name(name, "name")
    if any(ch in text for ch in ('"', "\r", "\n", ";")) or not text.strip():
        msg = "text cannot be empty or contain quotes, line breaks, or ';'."
        raise ValueError(msg)
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    if not execute_labtalk(f"label -p {float(x)} {float(y)} -n {safe_name} {text};"):
        msg = f"Could not add the annotation to {safe_graph}."
        raise ValueError(msg)
    return f"Added annotation '{text}' at ({x}, {y}) on {safe_graph}"


@mcp.tool()
def export_graph_sized(
    graph_name: str,
    file_path: str,
    width: int = 1200,
    height: int = 0,
    format: str = "png"
) -> str:
    """Export a graph to an image at a chosen pixel size (expGraph).

    Unlike export_graph (clipboard, page-size only), this controls the
    output pixel width/height directly.

    Args:
        graph_name: Graph to export
        file_path: Output path (Windows or WSL style)
        width: Image width in pixels (default 1200)
        height: Image height in pixels (0 = keep aspect ratio)
        format: png, jpg, tif, or bmp

    Returns:
        Path and pixel/byte size of the exported file
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_format = labtalk_choice(format.lower(), EXPORT_IMAGE_FORMATS, "format")
    safe_width = positive_int(width, "width")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    path = windows_path(file_path, "file_path")
    if not os.path.splitext(path)[1]:
        path = f"{path}.{safe_format}"
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    fname = os.path.splitext(os.path.basename(path))[0]
    size_opts = f"tr1.unit:=2 tr1.width:={safe_width}"
    if height > 0:
        size_opts += f" tr1.height:={positive_int(height, 'height')}"
    cmd = (
        f'expGraph type:={safe_format} path:="{out_dir}" filename:="{fname}" '
        f"overwrite:=replace {size_opts};"
    )
    if not execute_labtalk(cmd):
        msg = f"Origin could not export {safe_graph} to {path}."
        raise ValueError(msg)
    if not os.path.exists(path):
        msg = f"Export failed: {path} was not created."
        raise ValueError(msg)
    return f"Exported {safe_graph} to {path} ({safe_width}px wide, {os.path.getsize(path)} bytes)"


@mcp.tool()
def apply_color_map(graph_name: str, palette: str = "Fire") -> str:
    """Apply a color palette to a contour/heatmap/surface graph.

    Args:
        graph_name: Graph name (must hold a colormapped plot)
        palette: Palette name (a built-in Origin .pal), e.g. Fire,
            Rainbow, GrayScale, Maple, Thermometer, Temperature

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_palette = labtalk_name(palette, "palette")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    if not execute_labtalk(f"layer.cmap.load({safe_palette}.pal); layer.cmap.updateScale();"):
        msg = f"Could not apply palette '{safe_palette}' to {safe_graph}."
        raise ValueError(msg)
    return f"Applied '{safe_palette}' palette to {safe_graph}"


@mcp.tool()
def set_colormap_levels(graph_name: str, z_min: float, z_max: float) -> str:
    """Set the Z range (color scale levels) of a colormapped graph.

    Args:
        graph_name: Graph name (contour/heatmap/surface)
        z_min: Minimum Z for the color scale
        z_max: Maximum Z for the color scale

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    if z_max <= z_min:
        msg = "z_max must be greater than z_min."
        raise ValueError(msg)
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    if not execute_labtalk(
        f"layer.cmap.zmin = {float(z_min)}; layer.cmap.zmax = {float(z_max)}; "
        "layer.cmap.SetLevels(); layer.cmap.updateScale();"
    ):
        msg = f"Could not set colormap levels on {safe_graph}."
        raise ValueError(msg)
    return f"Set colormap Z range to [{z_min}, {z_max}] on {safe_graph}"


@mcp.tool()
def add_line(
    graph_name: str,
    x1: float,
    y1: float,
    x2: float,
    y2: float
) -> str:
    """Draw a straight line between two data points on a graph.

    Useful for guides, connectors, and trend indicators. For a line with
    an arrowhead, use add_arrow instead.

    Args:
        graph_name: Graph name
        x1, y1: Start point in data coordinates
        x2, y2: End point in data coordinates

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    coords = f"{{{float(x1)},{float(y1)},{float(x2)},{float(y2)}}}"
    if not execute_labtalk(f"draw -l {coords};"):
        msg = f"Could not draw a line on {safe_graph}."
        raise ValueError(msg)
    return f"Drew line ({x1},{y1})->({x2},{y2}) on {safe_graph}"


@mcp.tool()
def add_arrow(
    graph_name: str,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    double_headed: bool = False,
    head_size: int = 10
) -> str:
    """Draw an arrow from (x1,y1) to (x2,y2) at data coordinates.

    The arrowhead sits at the (x2,y2) end (and at the start too when
    double_headed). Set double_headed=False for a single-ended arrow.

    Args:
        graph_name: Graph name
        x1, y1: Tail (start) point in data coordinates
        x2, y2: Head (end) point in data coordinates
        double_headed: Put an arrowhead on both ends
        head_size: Arrowhead size in points (default 10)

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_size = positive_int(head_size, "head_size")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    name = "arr" + uuid.uuid4().hex[:8]
    coords = f"{{{float(x1)},{float(y1)},{float(x2)},{float(y2)}}}"
    if not execute_labtalk(f"draw -n {name} -l {coords};"):
        msg = f"Could not draw an arrow on {safe_graph}."
        raise ValueError(msg)
    width = max(safe_size * 2 // 3, 3)
    script = (
        f"{name}.arrowEndShape = 1; "
        f"{name}.arrowEndLength = {safe_size}; {name}.arrowEndWidth = {width};"
    )
    if double_headed:
        script += (
            f" {name}.arrowBeginShape = 1; "
            f"{name}.arrowBeginLength = {safe_size}; {name}.arrowBeginWidth = {width};"
        )
    execute_labtalk(script)
    ends = "double-headed" if double_headed else "single-headed"
    return f"Drew {ends} arrow ({x1},{y1})->({x2},{y2}) on {safe_graph}"
