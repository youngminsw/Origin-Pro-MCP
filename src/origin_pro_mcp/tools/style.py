from ..app import mcp
from ..origin_connection import execute_labtalk
from ..labtalk_safe import labtalk_choice, labtalk_name, labtalk_string, positive_int
from .style_helpers import (
    get_plot_names,
    graph_layer_execute,
    position_legend,
    set_legend_entries,
)

# Symbol shapes for different datasets
SYMBOL_SHAPES = {1: 2, 2: 3, 3: 1, 4: 4, 5: 5, 6: 6}  # circle, triangle-up, square, diamond, triangle-down, hexagon

# Default color order (colorblind-safe)
DEFAULT_COLORS = ["blue", "red", "green", "orange", "purple", "cyan"]

COLOR_MAP = {
    "black": 1, "red": 2, "green": 3, "blue": 4,
    "cyan": 5, "magenta": 6, "yellow": 7, "orange": 19,
    "purple": 13, "gray": 8, "grey": 8,
}


@mcp.tool()
def set_plot_style(
    graph_name: str,
    plot_index: int = 1,
    line_width: float = 1.5,
    symbol_size: int = 8,
    symbol_shape: int = 0,
    color: str = ""
) -> str:
    """Set line/symbol style for a data plot.

    Args:
        graph_name: Graph name
        plot_index: Plot index (1-based, order of data added)
        line_width: Line width in points (default 1.5)
        symbol_size: Symbol size (3-20, default 8)
        symbol_shape: 0=auto, 1=square, 2=circle, 3=triangle-up,
                      4=diamond, 5=triangle-down, 6=hexagon
        color: Color name (black, red, blue, green, orange, purple, cyan, magenta)

    Returns:
        Success message
    """
    import time
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    execute_labtalk(f"win -a {safe_graph_name};")
    plot_names = get_plot_names(safe_graph_name)

    idx = plot_index - 1
    if idx >= len(plot_names):
        return f"Plot index {plot_index} not found. Available: {plot_names}"

    pname = plot_names[idx]
    shape = symbol_shape if symbol_shape > 0 else SYMBOL_SHAPES.get(plot_index, 2)

    # Each set command needs a small delay to avoid Origin COM rendering races
    if color:
        safe_color = labtalk_choice(color.lower(), COLOR_MAP, "color")
        c = COLOR_MAP[safe_color]
        execute_labtalk(f"set {pname} -c {c};")
        time.sleep(0.2)

    lw = int(line_width * 10)
    execute_labtalk(f"set {pname} -w {lw};")
    time.sleep(0.2)

    execute_labtalk(f"set {pname} -k {shape};")
    time.sleep(0.2)

    execute_labtalk(f"set {pname} -z {symbol_size};")
    time.sleep(0.2)

    # Ensure line style is solid (don't use -kf 1 which makes hollow)
    execute_labtalk(f"set {pname} -l 1;")

    return f"Updated style for plot {plot_index} ({pname}) in {safe_graph_name}"


@mcp.tool()
def apply_publication_style(
    graph_name: str,
    x_label: str = "",
    y_label: str = "",
    x_min: float = None,
    x_max: float = None,
    y_min: float = None,
    y_max: float = None,
    legend_entries: str = "",
    legend_position: str = "top-right"
) -> str:
    """Apply complete publication styling to a graph in ONE call.

    Sets font (bold), colors, ticks, frame, and legend all at once.
    Designed to minimize token usage — call this once instead of many separate tools.

    Args:
        graph_name: Graph name
        x_label: X axis label with units, e.g. "Temperature (K)"
        y_label: Y axis label with units, e.g. "Absorbance (a.u.)"
        x_min: X axis minimum (None=auto)
        x_max: X axis maximum (None=auto)
        y_min: Y axis minimum (None=auto)
        y_max: Y axis maximum (None=auto)
        legend_entries: Comma-separated legend entries, e.g. "Sample A,Sample B"
        legend_position: top-left, top-right, bottom-left, bottom-right

    Returns:
        Summary of applied styling
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    execute_labtalk(f"win -a {safe_graph_name};")
    plot_names = get_plot_names(safe_graph_name)

    # 1. Axis titles — bold, Arial 28pt
    if x_label:
        x_title = labtalk_string("\\b(" + x_label + ")", "x_label")
        execute_labtalk(f'xb.text$ = {x_title}; xb.fsize = 28; xb.font$ = "Arial";')
    if y_label:
        y_title = labtalk_string("\\b(" + y_label + ")", "y_label")
        execute_labtalk(f'yl.text$ = {y_title}; yl.fsize = 28; yl.font$ = "Arial";')

    # 2. Tick labels — bold, Arial 22pt
    graph_layer_execute(graph_name, 'layer.x.label.pt = 22; layer.y.label.pt = 22;')
    graph_layer_execute(graph_name, 'layer.x.label.bold = 1; layer.y.label.bold = 1;')

    # 3. Axis range
    range_cmds = []
    if x_min is not None:
        range_cmds.append(f"layer.x.from = {x_min};")
    if x_max is not None:
        range_cmds.append(f"layer.x.to = {x_max};")
    if y_min is not None:
        range_cmds.append(f"layer.y.from = {y_min};")
    if y_max is not None:
        range_cmds.append(f"layer.y.to = {y_max};")
    if range_cmds:
        graph_layer_execute(safe_graph_name, " ".join(range_cmds))

    # 4. Ticks — inward, minor ticks on, proper lengths
    graph_layer_execute(safe_graph_name,
        "layer.x.ticks = 1; layer.y.ticks = 1; "
        "layer.x.minor = 4; layer.y.minor = 4; "
        "layer.x.majorLen = 8; layer.y.majorLen = 8;"
    )

    # 5. Frame — closed (4 sides), thick
    graph_layer_execute(safe_graph_name,
        "layer.x.opposite = 1; layer.y.opposite = 1; "
        "layer.x.thickness = 2; layer.y.thickness = 2;"
    )

    # 6. Remove grid lines
    graph_layer_execute(safe_graph_name,
        "layer.x.grid = 0; layer.y.grid = 0; "
        "layer.x.minorGrid = 0; layer.y.minorGrid = 0;"
    )

    # 7. Auto-style each plot with colorblind-safe colors + distinct symbols
    # Each command needs delay to avoid Origin COM rendering races
    import time
    for i, pname in enumerate(plot_names):
        color_name = DEFAULT_COLORS[i % len(DEFAULT_COLORS)]
        c = COLOR_MAP[color_name]
        shape = SYMBOL_SHAPES.get(i + 1, 2)

        execute_labtalk(f"set {pname} -c {c};")
        time.sleep(0.2)
        execute_labtalk(f"set {pname} -w 20;")
        time.sleep(0.2)
        execute_labtalk(f"set {pname} -k {shape};")
        time.sleep(0.2)
        execute_labtalk(f"set {pname} -z 10;")
        time.sleep(0.2)
        execute_labtalk(f"set {pname} -l 1;")  # solid line style
        time.sleep(0.2)

    # 8. Legend — reconstruct, then customize
    execute_labtalk(f"win -a {safe_graph_name}; legend -r;")
    time.sleep(0.3)

    if legend_entries:
        entry_list = [e.strip() for e in legend_entries.split(",")]
        set_legend_entries(safe_graph_name, entry_list)
        time.sleep(0.3)

    execute_labtalk('legend.fsize = 20; legend.font$ = "Arial";')
    position_legend(safe_graph_name, legend_position)

    styled = len(plot_names)
    return (
        f"Publication style applied to {safe_graph_name}: "
        f"{styled} plots styled, Arial bold labels, "
        f"inward ticks, closed frame, no grid"
    )


@mcp.tool()
def set_graph_font(
    graph_name: str,
    font_name: str = "Arial",
    font_size: int = 24,
    target: str = "all"
) -> str:
    """Set font for graph elements.

    Args:
        graph_name: Graph name
        font_name: Font family (e.g., Arial)
        font_size: Font size in points (default 24)
        target: "all", "axes", "title", "legend", "tick"

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_font_name = labtalk_string(font_name, "font_name")
    safe_font_size = positive_int(font_size, "font_size")
    safe_target = labtalk_choice(target, {"all", "axes", "title", "legend", "tick"}, "target")
    execute_labtalk(f"win -a {safe_graph_name};")

    if safe_target in ("all", "axes"):
        execute_labtalk(f"xb.font$ = {safe_font_name}; xb.fsize = {safe_font_size};")
        execute_labtalk(f"yl.font$ = {safe_font_name}; yl.fsize = {safe_font_size};")

    if safe_target in ("all", "tick"):
        tick_size = max(safe_font_size - 4, 16)
        graph_layer_execute(safe_graph_name, f"layer.x.label.pt = {tick_size}; layer.y.label.pt = {tick_size};")
        graph_layer_execute(safe_graph_name, "layer.x.label.bold = 1; layer.y.label.bold = 1;")

    if safe_target in ("all", "legend"):
        execute_labtalk(f"legend.font$ = {safe_font_name}; legend.fsize = {max(safe_font_size - 4, 16)};")

    if safe_target == "title":
        execute_labtalk(f"title.font$ = {safe_font_name}; title.fsize = {safe_font_size};")

    return f"Set font {font_name} {safe_font_size}pt on {safe_target} for {safe_graph_name}"


@mcp.tool()
def set_legend(
    graph_name: str,
    visible: bool = True,
    position: str = "top-right",
    entries: str = ""
) -> str:
    """Configure graph legend.

    Args:
        graph_name: Graph name
        visible: Show or hide legend
        position: top-left, top-right, bottom-left, bottom-right
        entries: Comma-separated custom legend entries

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_position = labtalk_choice(position, {"top-left", "top-right", "bottom-left", "bottom-right"}, "position")
    execute_labtalk(f"win -a {safe_graph_name};")

    if not visible:
        execute_labtalk("legend.show = 0;")
        return f"Legend hidden for {safe_graph_name}"

    execute_labtalk("legend -r;")

    if entries:
        import time
        time.sleep(0.3)
        entry_list = [e.strip() for e in entries.split(",")]
        set_legend_entries(safe_graph_name, entry_list)

    position_legend(safe_graph_name, safe_position)

    return f"Updated legend for {safe_graph_name}"


@mcp.tool()
def set_tick_style(
    graph_name: str,
    tick_direction: str = "in",
    major_length: int = 8,
    minor_count: int = 4,
    show_minor: bool = True
) -> str:
    """Set tick mark style.

    Args:
        graph_name: Graph name
        tick_direction: "in", "out", or "both"
        major_length: Major tick length in points (default 8)
        minor_count: Number of minor ticks between major ticks (default 4)
        show_minor: Whether to show minor ticks

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    dir_map = {"in": 1, "out": 2, "both": 3}
    safe_tick_direction = labtalk_choice(tick_direction, dir_map, "tick_direction")
    d = dir_map[safe_tick_direction]

    minor = minor_count if show_minor else 0

    graph_layer_execute(safe_graph_name,
        f"layer.x.ticks = {d}; layer.y.ticks = {d}; "
        f"layer.x.minor = {minor}; layer.y.minor = {minor}; "
        f"layer.x.majorLen = {major_length}; layer.y.majorLen = {major_length};"
    )

    return f"Updated tick style for {safe_graph_name}"
