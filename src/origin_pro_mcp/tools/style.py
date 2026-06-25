from ..app import mcp
from ..origin_connection import activate_window, execute_labtalk, get_lt_var, require_graph
from ..labtalk_safe import labtalk_choice, labtalk_name, labtalk_string, positive_int
from .style_helpers import (
    find_plot_column,
    get_plot_info,
    graph_layer_execute,
    place_legend_avoiding_data,
    set_legend_entries,
)

# Symbol shapes for different datasets
SYMBOL_SHAPES = {1: 2, 2: 3, 3: 1, 4: 4, 5: 5, 6: 6}  # circle, triangle-up, square, diamond, triangle-down, hexagon

# Default palette: muted/pastel RGB tones (no pure primaries — easier on
# the eyes, survives grayscale, colorblind-distinguishable). Applied via
# LabTalk color(r,g,b), verified working on Origin Pro 2020.
PASTEL_RGB = [
    (93, 143, 179),   # soft steel blue
    (204, 102, 119),  # muted rose
    (68, 170, 153),   # muted teal
    (221, 170, 102),  # soft amber
    (153, 136, 187),  # soft purple
    (119, 170, 187),  # gray cyan
]

COLOR_MAP = {
    "black": 1, "red": 2, "green": 3, "blue": 4,
    "cyan": 5, "magenta": 6, "yellow": 7, "orange": 19,
    "purple": 13, "gray": 8, "grey": 8,
}

# LabTalk `set -w` line-width units are tiny (~200 units per point,
# calibrated visually on Origin 2020) — NOT points*10.
_WIDTH_UNITS_PER_POINT = 200

# Publication-style line/symbol defaults for apply_publication_style,
# tuned visually on Origin 2020 for journal figures. Edit these to match
# your lab's taste — bigger symbols and heavier error bars read better in
# print than Origin's thin defaults.
_PUB_LINE_WIDTH_PT = 3.0        # data series line width (points)
_PUB_SYMBOL_SIZE = 14           # data symbol size
# Error-bar geometry. `set -erw` is in POINTS (unlike the data line's
# `-w`, which is ~200 units/pt), and `-erwc` is the cap (whisker) width.
# Both verified on Origin 2020 — earlier "blow-ups" were a units mistake
# (passing -w-scale values into -erw), not an Origin limitation.
_PUB_ERROR_BAR_WIDTH_PT = 2.5            # error-bar line width (points, via -erw)
_PUB_ERROR_CAP_WIDTH = _PUB_SYMBOL_SIZE - 2   # cap width, matched to symbol size


def _rgb(color: tuple) -> str:
    return f"color({color[0]},{color[1]},{color[2]})"


def _plot_has_symbols(plot_name: str) -> bool:
    """True when a plot renders symbols (scatter, line+symbol, area).

    Read via `get -k`: symbol plots report >= 1, column/bar/line report 0.
    Symbol commands (-k/-z) must never reach bar-type plots — they
    corrupt the bars' pattern (observed on Origin 2020).
    """
    if not execute_labtalk(f"__mcpk = 0; get {plot_name} -k __mcpk;"):
        return False
    return get_lt_var("__mcpk") >= 1


def _nice_increment(low: float, high: float):
    """A round tick increment giving roughly 5 major intervals, or None."""
    import math
    span = abs(high - low)
    if span <= 0:
        return None
    exponent = math.floor(math.log10(span / 5))
    best = None
    for mantissa in (1, 2, 2.5, 5, 10):
        inc = mantissa * 10 ** exponent
        intervals = span / inc
        if 3 <= intervals <= 8:
            score = abs(intervals - 5)
            if best is None or score < best[0]:
                best = (score, inc)
    return best[1] if best else None


@mcp.tool()
def set_plot_style(
    graph_name: str,
    plot_index: int = 1,
    line_width: float = 2.5,
    symbol_size: int = 8,
    symbol_shape: int = 0,
    color: str = ""
) -> str:
    """Set line/symbol style for a data plot.

    Args:
        graph_name: Graph name
        plot_index: Data series index (1-based, order the datasets were
                    added; error-bar plots are not counted)
        line_width: Line width in points (default 2.5)
        symbol_size: Symbol size (3-20, default 8)
        symbol_shape: 0=auto, 1=square, 2=circle, 3=triangle-up,
                      4=diamond, 5=triangle-down, 6=hexagon
        color: Color name (black, red, blue, green, orange, purple, cyan, magenta)

    Returns:
        Success message
    """
    import time
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    activate_window(safe_graph_name, "graph_name")
    data_plots = [p["name"] for p in get_plot_info(safe_graph_name) if not p["is_error"]]

    idx = plot_index - 1
    if idx >= len(data_plots):
        return f"Plot index {plot_index} not found. Available data plots: {data_plots}"

    pname = data_plots[idx]
    shape = symbol_shape if symbol_shape > 0 else SYMBOL_SHAPES.get(plot_index, 2)

    # Each set command needs a small delay to avoid Origin COM rendering races
    if color:
        safe_color = labtalk_choice(color.lower(), COLOR_MAP, "color")
        c = COLOR_MAP[safe_color]
        execute_labtalk(f"set {pname} -c {c};")
        time.sleep(0.2)

    lw = int(line_width * _WIDTH_UNITS_PER_POINT)
    execute_labtalk(f"set {pname} -w {lw};")
    time.sleep(0.2)

    if _plot_has_symbols(pname):
        execute_labtalk(f"set {pname} -k {shape};")
        time.sleep(0.2)
        execute_labtalk(f"set {pname} -z {symbol_size};")
    elif color:
        # bar/column-type plots: color the fill too
        execute_labtalk(f"set {pname} -cf {c};")

    return f"Updated style for plot {plot_index} ({pname}) in {safe_graph_name}"


@mcp.tool()
def apply_publication_style(
    graph_name: str,
    x_label: str = "",
    y_label: str = "",
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None,
    legend_entries: str = "",
    legend_position: str = "top-right"
) -> str:
    """Apply complete publication styling to a graph in ONE call.

    Sets bold Arial labels, a muted pastel color palette, 2.5 pt lines,
    readable tick spacing, inward ticks, closed frame, and a borderless
    bold legend all at once. Designed to minimize token usage — call this
    once instead of many separate tools.

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
    safe_legend_position = labtalk_choice(
        legend_position,
        {"top-left", "top-right", "bottom-left", "bottom-right"},
        "legend_position",
    )
    require_graph(safe_graph_name)
    activate_window(safe_graph_name, "graph_name")
    plot_infos = get_plot_info(safe_graph_name)

    # 1. Axis titles — bold, Arial 28pt
    if x_label:
        x_title = labtalk_string("\\b(" + x_label + ")", "x_label")
        execute_labtalk(f'xb.text$ = {x_title}; xb.fsize = 28; xb.font$ = "Arial";')
    if y_label:
        y_title = labtalk_string("\\b(" + y_label + ")", "y_label")
        execute_labtalk(f'yl.text$ = {y_title}; yl.fsize = 28; yl.font$ = "Arial";')

    # 2. Tick labels — bold, Arial 22pt
    graph_layer_execute(safe_graph_name, 'layer.x.label.pt = 22; layer.y.label.pt = 22;')
    graph_layer_execute(safe_graph_name, 'layer.x.label.bold = 1; layer.y.label.bold = 1;')

    # 3. Axis range + readable tick spacing (~5 major intervals)
    range_cmds = []
    if x_min is not None:
        range_cmds.append(f"layer.x.from = {x_min};")
    if x_max is not None:
        range_cmds.append(f"layer.x.to = {x_max};")
    if y_min is not None:
        range_cmds.append(f"layer.y.from = {y_min};")
    if y_max is not None:
        range_cmds.append(f"layer.y.to = {y_max};")
    if x_min is not None and x_max is not None:
        inc = _nice_increment(x_min, x_max)
        if inc is not None:
            range_cmds.append(f"layer.x.inc = {inc};")
    if y_min is not None and y_max is not None:
        inc = _nice_increment(y_min, y_max)
        if inc is not None:
            range_cmds.append(f"layer.y.inc = {inc};")
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

    # 7. Auto-style each data plot with a muted pastel palette + distinct
    # symbols. Error-bar plots only get the color of their data plot —
    # symbol/line commands would redraw them as connected lines.
    # Each command needs delay to avoid Origin COM rendering races.
    import time
    data_index = 0
    current_color = _rgb(PASTEL_RGB[0])
    for info in plot_infos:
        pname = info["name"]
        if info["is_error"]:
            execute_labtalk(f"set {pname} -c {current_color};")
            time.sleep(0.2)
            execute_labtalk(f"set {pname} -erw {_PUB_ERROR_BAR_WIDTH_PT};")  # error-bar line width (pt)
            time.sleep(0.2)
            execute_labtalk(f"set {pname} -erwc {_PUB_ERROR_CAP_WIDTH};")  # cap width ~ symbol
            time.sleep(0.2)
            continue
        current_color = _rgb(PASTEL_RGB[data_index % len(PASTEL_RGB)])
        shape = SYMBOL_SHAPES.get(data_index + 1, 2)
        data_index += 1

        execute_labtalk(f"set {pname} -c {current_color};")
        time.sleep(0.2)
        execute_labtalk(f"set {pname} -w {int(_PUB_LINE_WIDTH_PT * _WIDTH_UNITS_PER_POINT)};")  # data lines
        time.sleep(0.2)
        if _plot_has_symbols(pname):
            execute_labtalk(f"set {pname} -k {shape};")
            time.sleep(0.2)
            execute_labtalk(f"set {pname} -z {_PUB_SYMBOL_SIZE};")
            time.sleep(0.2)
        else:
            # bar/column/area-type plots: color the fill instead
            execute_labtalk(f"set {pname} -cf {current_color};")
            time.sleep(0.2)

    # 8. Legend — reconstruct, then customize: bold entries, no border
    execute_labtalk(f"win -a {safe_graph_name}; legend -r;")
    time.sleep(0.3)

    if legend_entries:
        entry_list = [f"\\b({e.strip()})" for e in legend_entries.split(",")]
        set_legend_entries(safe_graph_name, entry_list)
        time.sleep(0.3)
    else:
        # Bold the auto legend (from column long names) for consistency
        renamed = False
        for info in plot_infos:
            if info["is_error"]:
                continue
            col = find_plot_column(info["name"])
            if col is not None and col.LongName and not col.LongName.startswith("\\b("):
                col.LongName = f"\\b({col.LongName})"
                renamed = True
        if renamed:
            execute_labtalk(f"win -a {safe_graph_name}; legend -r;")
            time.sleep(0.3)

    execute_labtalk('legend.fsize = 20; legend.font$ = "Arial"; legend.background = 0;')
    placement = place_legend_avoiding_data(safe_graph_name, safe_legend_position)

    moved_out = " — moved outside the frame to avoid the data" if placement.startswith("outside") else ""
    return (
        f"Publication style applied to {safe_graph_name}: "
        f"{data_index} data plots styled (pastel palette, {_PUB_LINE_WIDTH_PT} pt lines), "
        f"Arial bold labels, inward ticks, closed frame, no grid, "
        f"borderless bold legend ({placement}){moved_out}"
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
    require_graph(safe_graph_name)
    activate_window(safe_graph_name, "graph_name")

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
    require_graph(safe_graph_name)
    activate_window(safe_graph_name, "graph_name")

    if not visible:
        execute_labtalk("legend.show = 0;")
        return f"Legend hidden for {safe_graph_name}"

    execute_labtalk("legend -r;")

    if entries:
        import time
        time.sleep(0.3)
        entry_list = [f"\\b({e.strip()})" for e in entries.split(",")]
        set_legend_entries(safe_graph_name, entry_list)

    execute_labtalk("legend.background = 0;")
    placement = place_legend_avoiding_data(safe_graph_name, safe_position)

    moved_out = " (moved outside the frame to avoid the data)" if placement.startswith("outside") else ""
    return f"Updated legend for {safe_graph_name}: placed {placement}{moved_out}"


def _set_tick_style_impl(
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
    require_graph(safe_graph_name)
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
