from ..app import mcp
from ..origin_connection import (
    activate_window, execute_labtalk, get_lt_str, get_lt_var,
    require_graph,
)
from ..labtalk_safe import (
    labtalk_choice, labtalk_name, labtalk_string, positive_int,
    validate_text_escapes,
)
from .style_helpers import (
    _collect_xy,
    find_plot_column,
    get_plot_info,
    graph_layer_execute,
    place_legend_avoiding_data,
    require_data_plots,
    set_legend_entries,
    settle_new_plots,
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
    """A round tick increment giving ~4 major intervals (4-6 ticks), or None.

    Caps the count at 5 intervals (so at most 6 major ticks) to keep the axis
    uncrowded, and scores toward 4 intervals.
    """
    import math
    span = abs(high - low)
    if span <= 0:
        return None
    exponent = math.floor(math.log10(span / 4))
    best = None
    for mantissa in (1, 2, 2.5, 5, 10):
        inc = mantissa * 10 ** exponent
        intervals = span / inc
        if 3 <= intervals <= 5:
            score = abs(intervals - 4)
            if best is None or score < best[0]:
                best = (score, inc)
    return best[1] if best else None


def _tight_axis_cmds(
    graph_name: str,
    x_min: float | None,
    x_max: float | None,
    y_min: float | None,
    y_max: float | None,
) -> list[str]:
    """LabTalk commands that tighten each axis to the data (best-effort).

    For each axis the effective [lo, hi] is the explicit min/max when given,
    else the data extent read via ``_collect_xy``. ``from``/``to`` are set
    TIGHT (no padding, so the curve touches the axes; if the data starts at 0
    the axis starts at 0) plus a capped tick increment. Never raises: when the
    data can't be read (or an axis has no usable points) that axis is left on
    Origin's auto range.
    """
    try:
        xs, ys = _collect_xy(graph_name)
    except Exception:
        xs, ys = [], []
    x_vals = [x for x in xs if x is not None]

    cmds: list[str] = []
    for prefix, lo_opt, hi_opt, values in (
        ("x", x_min, x_max, x_vals),
        ("y", y_min, y_max, ys),
    ):
        lo = lo_opt if lo_opt is not None else (min(values) if values else None)
        hi = hi_opt if hi_opt is not None else (max(values) if values else None)
        if lo is None or hi is None:
            continue
        cmds.append(f"layer.{prefix}.from = {lo}; layer.{prefix}.to = {hi};")
        inc = _nice_increment(lo, hi)
        if inc is not None:
            cmds.append(f"layer.{prefix}.inc = {inc};")
    return cmds


@mcp.tool()
def set_plot_style(
    graph_name: str,
    plot_index: int = 1,
    line_width: float = 2.5,
    symbol_size: int = 8,
    symbol_shape: int = 0,
    color: str = "",
    rgb: str = "",
    open_symbol: bool = False,
) -> str:
    """Set line/symbol style for a data plot.

    Args:
        graph_name: Graph name
        plot_index: Data series index (1-based, order the datasets were
                    added; error-bar plots are not counted)
        line_width: Line width in points (default 2.5)
        symbol_size: Symbol size (default 8; not validated — Origin accepts any
                     positive value, but roughly 3-20 is the readable range)
        symbol_shape: 0=auto, 1=square, 2=circle, 3=triangle-up,
                      4=diamond, 5=triangle-down, 6=hexagon
        color: Color name (black, red, green, blue, cyan, magenta, yellow,
               orange, purple, gray/grey)
        rgb: Explicit "r,g,b" (each 0-255), e.g. "128,0,200", for per-curve
             rainbow/gradient colors that named colors can't express. Overrides
             `color` when given.
        open_symbol: True = open/hollow marker interior (publication standard),
                     False = solid fill (LabTalk `set -kf 1` vs `0`).

    NOTE: color applies only on UNGROUPED plots. create_graph + add_plot_to_graph
    build ungrouped plots (fine). For a grouped multi-curve plot (loaded project
    or a single multi-Y plot), the group's color increment overrides this — call
    ungroup_plots(graph_name) first. (line_width/symbol still apply while grouped.)

    Returns:
        Success message
    """
    import time
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    # require_data_plots activates the page and takes a fresh handle first, then
    # RAISES an actionable error if the layer still exposes no data plots (the
    # loaded-from-.opju freeze) — so this never silently no-ops on a graph whose
    # plots Origin won't reveal over COM.
    infos, data_plots = require_data_plots(safe_graph_name)

    idx = plot_index - 1
    if idx < 0 or idx >= len(data_plots):
        valid = f"1-{len(data_plots)}" if data_plots else "(none)"
        msg = (
            f"Plot index {plot_index} not found on {safe_graph_name}. "
            f"Valid range: {valid}. Data plots: {data_plots}"
        )
        raise ValueError(msg)

    pname = data_plots[idx]
    shape = symbol_shape if symbol_shape > 0 else SYMBOL_SHAPES.get(plot_index, 2)
    # Target the EXACT plot by its DATASET NAME (`set <name> ...`) run on the
    # graph's Layer1 COM object (graph_layer_execute -> gl.Execute). Verified on
    # Origin 2020 (isolated instance, export pixel-checked) to color each plot of
    # an ungrouped multi-curve graph independently. This deliberately avoids two
    # dead ends: `layer -s <N>; set %C` only ever selects plot 1 (N>=2 no-ops),
    # and global execute_labtalk needs `win -a`, which fails on .opju-loaded
    # graphs and froze all styling there. gl.Execute + the dataset name needs
    # neither an active window nor plot selection.
    def _set(spec: str) -> None:
        graph_layer_execute(safe_graph_name, f"set {pname} {spec};")

    # Resolve the color expression: explicit rgb overrides a named color.
    c = None
    if rgb:
        r, g, b = _parse_rgb(rgb)
        c = f"color({r},{g},{b})"
    elif color:
        c = COLOR_MAP[labtalk_choice(color.lower(), COLOR_MAP, "color")]

    # All flags for this plot are sent as ONE `set` command (LabTalk supports
    # multiple -flags per call) so there is only one COM-render settle to wait
    # for, instead of one per flag.
    specs = []
    if c is not None:
        specs.append(f"-c {c}")
    lw = int(line_width * _WIDTH_UNITS_PER_POINT)
    specs.append(f"-w {lw}")

    if _plot_has_symbols(pname):
        specs.append(f"-k {shape}")
        specs.append(f"-z {symbol_size}")
        # Symbol interior: 1 = Open (hollow), 0 = Solid (verified on Origin 2020).
        specs.append(f"-kf {1 if open_symbol else 0}")
    elif c is not None:
        # bar/column-type plots: color the fill too
        specs.append(f"-cf {c}")

    _set(" ".join(specs))
    time.sleep(0.2)

    grouping_note = ""
    if len(data_plots) > 1:
        # Can't reliably read Origin's group state over COM on this build, so
        # this is a static caveat (not a live check) whenever more than one
        # data plot exists — grouping only matters when there's more than one.
        grouping_note = (
            "; NOTE: if these plots are GROUPED, per-plot colors may be "
            "overridden by the group — call ungroup_plots first if colors "
            "don't look right"
        )
    return (
        f"Updated style for plot {plot_index} ({pname}) in "
        f"{safe_graph_name}{grouping_note}"
    )




def _rgb_component(value: int, field: str) -> int:
    v = int(value)
    if v < 0 or v > 255:
        raise ValueError(f"{field} must be 0-255, got {value}.")
    return v


def _parse_rgb(rgb: str):
    """Parse an 'r,g,b' string into a validated (r, g, b) tuple (each 0-255)."""
    parts = [p.strip() for p in str(rgb).split(",")]
    if len(parts) != 3:
        raise ValueError(f"rgb must be 'r,g,b' (three 0-255 values), got '{rgb}'.")
    try:
        vals = [int(p) for p in parts]
    except ValueError:
        raise ValueError(f"rgb components must be integers, got '{rgb}'.")
    return tuple(_rgb_component(v, "rgb") for v in vals)


@mcp.tool()
def ungroup_plots(graph_name: str, plot_type: str = "line") -> str:
    """Break a plot group so each curve can be colored independently.

    Origin 2020 has NO working ungroup command: `layer -g` only GROUPS (there is
    no ungroup toggle), verified exhaustively. While plots are grouped, the
    group's color increment overrides per-curve `set_plot_style` color. The only
    reliable fix (spike-verified on Origin 2020) is to remove the grouped plots
    and re-plot each source dataset on its own — separate plots are NOT grouped.
    This preserves the data and the X axis but RESETS each curve to `plot_type`
    (re-style afterward with set_plot_style). Best used before styling.

    Args:
        graph_name: Graph name.
        plot_type: Plot type for the rebuilt curves: line, scatter, line+symbol.

    Returns:
        Confirmation with the rebuilt plot count.
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    codes = {"line": 200, "scatter": 201, "line+symbol": 202}
    ptype = codes[labtalk_choice(plot_type, codes, "plot_type")]
    # Activates + fresh handle, then RAISES if the layer exposes no plots (the
    # loaded-graph freeze) rather than reporting a no-op ungroup as success.
    infos, data_names = require_data_plots(safe_graph)
    has_error = any(p["is_error"] for p in infos)
    # Remove every plot, then re-plot each DATA dataset on its own. Run on the
    # Layer1 COM object (gl.Execute) so it works without an active window (win -a
    # fails on .opju-loaded graphs). `plotxy iy:=<dataset>` reuses the sheet's X
    # designation, preserving the original X.
    for p in infos:
        graph_layer_execute(safe_graph, f"layer -e {p['name']};")
    rebuilt = 0
    for name in data_names:
        if graph_layer_execute(
            safe_graph,
            f"plotxy iy:={name} plot:={ptype} ogl:=[{safe_graph}]Layer1;",
        ):
            rebuilt += 1
    if rebuilt:
        settle_new_plots(safe_graph, expected_min_plots=rebuilt)
    note = (" Error-bar plots were dropped — re-add them with set_error_bars."
            if has_error else "")
    return (f"Ungrouped {safe_graph}: rebuilt {rebuilt} independent {plot_type} "
            f"plot(s) on Layer1. Color each with set_plot_style.{note}")


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

    NOTE: per-curve COLORS apply only on UNGROUPED plots. Graphs built with
    create_graph + add_plot_to_graph are ungrouped (fine). A grouped multi-curve
    plot (loaded from a project, or a single multi-Y plot) shares one color
    increment that overrides per-curve colors, so the palette will NOT apply —
    call ungroup_plots(graph_name) first, then this. Axes/frame/labels apply
    either way.

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
        validate_text_escapes(x_label, "x_label")
        x_title = labtalk_string("\\b(" + x_label + ")", "x_label")
        execute_labtalk(f'xb.text$ = {x_title}; xb.fsize = 28; xb.font$ = "Arial";')
    if y_label:
        validate_text_escapes(y_label, "y_label")
        y_title = labtalk_string("\\b(" + y_label + ")", "y_label")
        execute_labtalk(f'yl.text$ = {y_title}; yl.fsize = 28; yl.font$ = "Arial";')

    # 2. Tick labels — bold, Arial 22pt
    graph_layer_execute(safe_graph_name, 'layer.x.label.pt = 22; layer.y.label.pt = 22;')
    graph_layer_execute(safe_graph_name, 'layer.x.label.bold = 1; layer.y.label.bold = 1;')

    # 3. Axis range — TIGHT to the data by default (no empty gap before the
    # first or after the last point; if the data starts at 0 the axis starts
    # at 0) + readable tick spacing (~4 major intervals). Explicit min/max
    # override the data extent. Best-effort: unreadable data leaves the axis
    # on Origin's auto range.
    range_cmds = _tight_axis_cmds(safe_graph_name, x_min, x_max, y_min, y_max)
    if range_cmds:
        graph_layer_execute(safe_graph_name, " ".join(range_cmds))

    # 4. Ticks — inward, minor ticks on, proper lengths
    graph_layer_execute(safe_graph_name,
        "layer.x.ticks = 1; layer.y.ticks = 1; "
        "layer.x.minor = 1; layer.y.minor = 1; "
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
    # Each plot's flags are sent as ONE `set` command (LabTalk supports
    # multiple -flags per call) with a single post-command settle, instead of
    # one Execute + sleep per flag — cuts an 8-plot graph from ~24 sleeps to 8.
    import time
    data_index = 0
    current_color = _rgb(PASTEL_RGB[0])
    for info in plot_infos:
        pname = info["name"]
        if info["is_error"]:
            execute_labtalk(
                f"set {pname} -c {current_color} "
                f"-erw {_PUB_ERROR_BAR_WIDTH_PT} -erwc {_PUB_ERROR_CAP_WIDTH};"
            )
            time.sleep(0.2)
            continue
        current_color = _rgb(PASTEL_RGB[data_index % len(PASTEL_RGB)])
        shape = SYMBOL_SHAPES.get(data_index + 1, 2)
        data_index += 1

        line_width_units = int(_PUB_LINE_WIDTH_PT * _WIDTH_UNITS_PER_POINT)
        if _plot_has_symbols(pname):
            execute_labtalk(
                f"set {pname} -c {current_color} -w {line_width_units} "
                f"-k {shape} -z {_PUB_SYMBOL_SIZE};"
            )
        else:
            # bar/column/area-type plots: color the fill instead
            execute_labtalk(
                f"set {pname} -c {current_color} -w {line_width_units} "
                f"-cf {current_color};"
            )
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
    if data_index == 0:
        # Axes/frame/labels above still applied, but no data plot was found to
        # style — do not let that read as full success. On a graph loaded from
        # a .opju whose layer stays frozen even after activation, per-curve
        # styling is impossible; say so instead of implying the curves were
        # styled.
        return (
            f"Publication style applied to {safe_graph_name}: axes, frame and "
            f"labels set, but NO data plots were found to style. If this graph "
            f"was loaded from a project file (.opju) and its curves look "
            f"unstyled, Origin is freezing its plot list over COM — recreate "
            f"the graph in-session (create_graph / plotxy)."
        )
    grouping_note = ""
    if data_index > 1:
        # Can't reliably read Origin's group state over COM on this build, so
        # this is a static caveat (not a live check) whenever more than one
        # data plot was styled — grouping only matters when there's more than
        # one plot, and a grouped layer overrides the per-plot palette below.
        grouping_note = (
            "; NOTE: if these plots are GROUPED, per-plot colors may be "
            "overridden by the group — call ungroup_plots first if the "
            "palette doesn't look right"
        )
    return (
        f"Publication style applied to {safe_graph_name}: "
        f"{data_index} data plots styled (pastel palette, {_PUB_LINE_WIDTH_PT} pt lines), "
        f"Arial bold labels, inward ticks, closed frame, no grid, "
        f"borderless bold legend ({placement}){moved_out}{grouping_note}"
    )


def _bold_text_object(obj: str) -> None:
    """Bold an axis-title / graph-title text object by wrapping its current
    text in `\\b(...)`. Origin 2020 exposes no `.bold` on these objects, so
    this rich-text markup is the only reliable route. Idempotent (won't
    double-wrap); a no-op when the text can't be read."""
    current = get_lt_str(f"{obj}.text$")
    if not current or current.startswith("\\b(") or '"' in current:
        return
    execute_labtalk(f'{obj}.text$ = {labtalk_string(chr(92) + "b(" + current + ")", obj)};')


@mcp.tool()
def set_graph_font(
    graph_name: str,
    font_name: str = "Arial",
    font_size: int = 24,
    target: str = "all",
    bold: bool = False,
) -> str:
    """Set font for graph elements.

    Args:
        graph_name: Graph name
        font_name: Font family (e.g., Arial)
        font_size: Font size in points (default 24)
        target: "all", "axes", "title", "legend", "tick"
        bold: When True, also bold the targeted element(s). Tick labels are
              always bold (Origin publication default); this bolds axis titles
              and the graph title via `\\b(...)` markup for the chosen target.

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_font_name = labtalk_string(font_name, "font_name")
    safe_font_size = positive_int(font_size, "font_size")
    safe_target = labtalk_choice(target, {"all", "axes", "title", "legend", "tick"}, "target")
    require_graph(safe_graph_name)
    activate_window(safe_graph_name, "graph_name")

    # Bold: Origin 2020 has NO `.bold` property on the axis-title/title text
    # objects (only tick labels expose `layer.x.label.bold`). Bolding a title
    # is done by wrapping its text in the `\b(...)` rich-text markup.
    if safe_target in ("all", "axes"):
        if not execute_labtalk(f"xb.font$ = {safe_font_name}; xb.fsize = {safe_font_size};"):
            msg = f"Could not set the x-axis title font on {safe_graph_name}."
            raise ValueError(msg)
        if not execute_labtalk(f"yl.font$ = {safe_font_name}; yl.fsize = {safe_font_size};"):
            msg = f"Could not set the y-axis title font on {safe_graph_name}."
            raise ValueError(msg)
        if bold:
            _bold_text_object("xb")
            _bold_text_object("yl")

    if safe_target in ("all", "tick"):
        tick_size = max(safe_font_size - 4, 16)
        if not graph_layer_execute(safe_graph_name, f"layer.x.label.pt = {tick_size}; layer.y.label.pt = {tick_size};"):
            msg = f"Could not set tick label font size on {safe_graph_name}."
            raise ValueError(msg)
        if not graph_layer_execute(safe_graph_name, "layer.x.label.bold = 1; layer.y.label.bold = 1;"):
            msg = f"Could not bold tick labels on {safe_graph_name}."
            raise ValueError(msg)

    if safe_target in ("all", "legend"):
        if not execute_labtalk(f"legend.font$ = {safe_font_name}; legend.fsize = {max(safe_font_size - 4, 16)};"):
            msg = f"Could not set the legend font on {safe_graph_name}."
            raise ValueError(msg)

    if safe_target == "title":
        if not execute_labtalk(f"title.font$ = {safe_font_name}; title.fsize = {safe_font_size};"):
            msg = f"Could not set the title font on {safe_graph_name}."
            raise ValueError(msg)
        if bold:
            _bold_text_object("title")

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
        if not execute_labtalk("legend.show = 0;"):
            msg = f"Could not hide the legend on {safe_graph_name}."
            raise ValueError(msg)
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
    axis: str = "both",
    tick_direction: str = "in",
    major_length: int = 8,
    minor_count: int = 4,
    show_minor: bool = True
) -> str:
    """Set tick mark style.

    Args:
        graph_name: Graph name
        axis: "x", "y", or "both" (default "both")
        tick_direction: "in", "out", or "both"
        major_length: Major tick length in points (default 8)
        minor_count: Number of minor ticks between major ticks (default 4)
        show_minor: Whether to show minor ticks

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    safe_axis = labtalk_choice(axis.lower(), {"x", "y", "both"}, "axis")
    dir_map = {"in": 1, "out": 2, "both": 3}
    safe_tick_direction = labtalk_choice(tick_direction, dir_map, "tick_direction")
    d = dir_map[safe_tick_direction]

    minor = minor_count if show_minor else 0

    axes = ["x", "y"] if safe_axis == "both" else [safe_axis]
    cmds = " ".join(
        f"layer.{a}.ticks = {d}; layer.{a}.minor = {minor}; "
        f"layer.{a}.majorLen = {major_length};"
        for a in axes
    )
    if not graph_layer_execute(safe_graph_name, cmds):
        msg = f"Could not update {safe_axis} tick style for {safe_graph_name}."
        raise ValueError(msg)

    return f"Updated {safe_axis} tick style for {safe_graph_name}"


# layer.axis.label.numFormat codes (verified against OriginLab docs).
_TICK_LABEL_FORMATS = {"decimal": 1, "scientific": 2, "engineering": 3}


@mcp.tool()
def set_tick_labels(
    graph_name: str,
    axis: str = "both",
    format: str | None = None,
    bold: bool | None = None,
    decimal_places: int | None = None,
) -> str:
    """Format an axis's tick labels: numeric format, bold, decimal places.

    For a log axis Origin already renders ticks as powers of ten (10^n) by
    default — set the scale with axis(op="scale", scale="log10"); there is no
    separate "powers10" tick format on this build. On a linear axis, use
    format="scientific" for the 1E3 style or "decimal" for plain numbers.

    Args:
        graph_name: Graph name
        axis: "x", "y", or "both" (default)
        format: "decimal", "scientific", or "engineering" (None = leave)
        bold: True/False to bold/unbold tick labels (None = leave)
        decimal_places: Number of decimals to show, or -1 for auto (None = leave)

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_axis = labtalk_choice(axis.lower(), {"x", "y", "both"}, "axis")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    if format is None and bold is None and decimal_places is None:
        msg = "Provide at least one of format, bold, or decimal_places."
        raise ValueError(msg)
    axes = ["x", "y"] if safe_axis == "both" else [safe_axis]
    changed = []
    cmds = []
    for a in axes:
        if format is not None:
            safe_fmt = labtalk_choice(format.lower(), _TICK_LABEL_FORMATS, "format")
            cmds.append(f"layer.{a}.label.numFormat = {_TICK_LABEL_FORMATS[safe_fmt]};")
        if bold is not None:
            cmds.append(f"layer.{a}.label.bold = {1 if bold else 0};")
        if decimal_places is not None:
            cmds.append(f"layer.{a}.label.decPlaces = {int(decimal_places)};")
    if format is not None:
        changed.append("format")
    if bold is not None:
        changed.append("bold")
    if decimal_places is not None:
        changed.append("decimal_places")
    if not graph_layer_execute(safe_graph, " ".join(cmds)):
        msg = f"Could not update tick labels ({', '.join(changed)}) for {safe_graph}."
        raise ValueError(msg)
    return f"Updated {safe_axis} tick labels for {safe_graph}: {', '.join(changed)}"
