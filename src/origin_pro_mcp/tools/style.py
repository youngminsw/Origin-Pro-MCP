from ..app import mcp
import time

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

# Per-plot-index default symbol shape, used when symbol_shape is 0/None (auto).
# Codes are Origin 2020's LabTalk `set <ds> -k <n>` symbol-kind table, RE-VERIFIED
# live (Task 5, single-flag `-k <n>` only, no `-kf`/`-z` combined, on a settled
# page): 1=square, 2=circle, 3=triangle-up, 4=triangle-down, 5=diamond, 6=plus,
# 7=x/cross, 8=asterisk. (9=dash, 10=vertical-bar, 11/12=render as literal "1"/
# "A" glyphs, not real marker shapes — excluded here as not useful defaults.)
# The earlier table (1=circle,2=triangle-up,3=square,4=diamond,5=triangle-down,
# 6=hexagon) was WRONG — a pre-settle-fix, flag-combination-corrupted reading;
# do not resurrect it.
SYMBOL_SHAPES = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6}  # square, circle, tri-up, tri-down, diamond, plus

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


def _set_axis_title_verified(
    graph_name: str, obj: str, label_text: str, font_size: int = 28
) -> bool:
    """Write an axis-title text object (``xb`` bottom-X / ``yl`` left-Y), bold
    Arial, then CONFIRM it landed by reading ``<obj>.text$`` back.

    Origin's active-window text-object writes can silently hit the wrong window
    or no-op on the new-page settle race — the exact bug that left
    apply_publication_style's titles reading back as the ``%(?X)`` column-name
    placeholder while the tool claimed "Arial bold labels" (usability agent,
    3/3). LabTalk returns success either way, so success is gated on the
    read-back actually containing the label words, with an activate + settle
    retry. Returns True when verified, False otherwise (caller must NOT claim
    the label was set)."""
    safe = labtalk_string("\\b(" + label_text + ")", obj)
    for attempt in range(3):
        activate_window(graph_name, "graph_name")
        execute_labtalk(
            f'{obj}.text$ = {safe}; {obj}.fsize = {font_size}; {obj}.font$ = "Arial";'
        )
        got = get_lt_str(f"{obj}.text$") or ""
        if label_text in got:
            return True
        if attempt < 2:
            time.sleep(0.3)
    return False


@mcp.tool()
def set_plot_style(
    graph_name: str,
    plot_index: int = 1,
    line_width: float | None = None,
    symbol_size: int | None = None,
    symbol_shape: int | None = None,
    color: str = "",
    rgb: str = "",
    open_symbol: bool | None = None,
    error_bar_width: float | None = None,
    error_cap_width: float | None = None,
) -> str:
    """Set line/symbol/error-bar style for a data plot. Only the aspects you
    pass are touched — everything else on that curve is left exactly as it
    was.

    Args:
        graph_name: Graph name
        plot_index: Data series index (1-based, order the datasets were
                    added; error-bar plots are not counted)
        line_width: Line width in points (None = leave unchanged)
        symbol_size: Symbol size (None = leave unchanged; not validated —
                     Origin accepts any positive value, roughly 3-20 is the
                     readable range)
        symbol_shape: 0=auto (uses SYMBOL_SHAPES's per-plot-index default), else
                      the LabTalk `-k` code, RE-VERIFIED live on Origin 2020:
                      1=square, 2=circle, 3=triangle-up, 4=triangle-down,
                      5=diamond, 6=plus, 7=x/cross, 8=asterisk. (9-12 render as
                      a dash/vertical-bar/literal digit-or-letter glyph, not
                      useful marker shapes — avoid them.) None = leave unchanged.
        color: Color name (black, red, green, blue, cyan, magenta, yellow,
               orange, purple, gray/grey). "" = leave unchanged.
        rgb: Explicit "r,g,b" (each 0-255), e.g. "128,0,200", for per-curve
             rainbow/gradient colors that named colors can't express. Overrides
             `color` when given. "" = leave unchanged.
        open_symbol: True = open/hollow marker interior (publication standard),
                     False = solid fill (LabTalk `set -kf 1` vs `0`). None =
                     leave unchanged.
        error_bar_width: Error-bar LINE width in POINTS (LabTalk `-erw`, NOT
                     the same unit scale as `line_width`'s `-w`). None = leave
                     unchanged. Requires the plot to have an adjacent error
                     bar (from y_error_col or set_error_bars).
        error_cap_width: Error-bar CAP (whisker) width (LabTalk `-erwc`). None
                     = leave unchanged. Same adjacency requirement as above.

    GOTCHA (probe-verified, Origin 2020): every LabTalk flag below is sent as
    its OWN `set <ds> -flag val;` call — combining e.g. `-c` and `-cf` (or
    `-k`+`-kf`+`-z`) into ONE `set` command silently resets the plot to BLACK
    or blanks the symbol. Never batch flags yourself via run_labtalk either.

    NOTE on grouping: per-curve color/width/symbol/fill ALL apply only on
    UNGROUPED plots (probe-confirmed — the earlier claim that line_width/
    symbol still apply while grouped was WRONG). create_graph +
    add_plot_to_graph build ungrouped plots (fine). A grouped multi-curve plot
    (loaded project, or a single multi-Y plotxy) shares one color/style
    increment that overrides ALL of this — call ungroup_plots(graph_name)
    first, then re-style.

    Returns:
        Success message. Raises ValueError if no aspect was given, or if
        plot_index is out of range, or if error_bar_width/error_cap_width is
        requested but the plot has no error bars.
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

    # Resolve the color expression: explicit rgb overrides a named color.
    c = None
    if rgb:
        r, g, b = _parse_rgb(rgb)
        c = f"color({r},{g},{b})"
    elif color:
        c = COLOR_MAP[labtalk_choice(color.lower(), COLOR_MAP, "color")]

    touch_symbols = (
        symbol_shape is not None or symbol_size is not None or open_symbol is not None
    )
    nothing_requested = (
        c is None and line_width is None and not touch_symbols
        and error_bar_width is None and error_cap_width is None
    )
    if nothing_requested:
        msg = (
            "set_plot_style: nothing to change — pass at least one of "
            "line_width, symbol_size, symbol_shape, color, rgb, open_symbol, "
            "error_bar_width, or error_cap_width."
        )
        raise ValueError(msg)

    # Target the EXACT plot by its DATASET NAME (`set <name> ...`) run on the
    # graph's Layer1 COM object (graph_layer_execute -> gl.Execute). Verified on
    # Origin 2020 (isolated instance, export pixel-checked) to color each plot of
    # an ungrouped multi-curve graph independently. This deliberately avoids two
    # dead ends: `layer -s <N>; set %C` only ever selects plot 1 (N>=2 no-ops),
    # and global execute_labtalk needs `win -a`, which fails on .opju-loaded
    # graphs and froze all styling there. gl.Execute + the dataset name needs
    # neither an active window nor plot selection.
    def _set(dataset: str, spec: str) -> None:
        graph_layer_execute(safe_graph_name, f"set {dataset} {spec};")

    # P8 HARD RULE (probe-verified): one flag per `set` call. Combining `-c` +
    # `-cf` in one command wipes the plot to black; `-k`+`-kf`+`-z` combined
    # can blank the symbol entirely. Never join these into one string.
    flags: list[str] = []
    if c is not None:
        flags.append(f"-c {c}")
    if line_width is not None:
        flags.append(f"-w {int(line_width * _WIDTH_UNITS_PER_POINT)}")
    has_symbols = _plot_has_symbols(pname)
    if touch_symbols and has_symbols:
        if symbol_shape is not None:
            shape = symbol_shape if symbol_shape > 0 else SYMBOL_SHAPES.get(plot_index, 2)
            flags.append(f"-k {shape}")
        if symbol_size is not None:
            flags.append(f"-z {symbol_size}")
        if open_symbol is not None:
            # Symbol interior: 1 = Open (hollow), 0 = Solid (Origin 2020).
            flags.append(f"-kf {1 if open_symbol else 0}")
    if c is not None:
        # Symbol interior fill / bar fill follows the curve color — sent as
        # its OWN call (P8), never combined with the `-c` flag above.
        flags.append(f"-cf {c}")

    for flag in flags:
        _set(pname, flag)
    if flags:
        time.sleep(0.2)

    error_note = ""
    if error_bar_width is not None or error_cap_width is not None:
        error_note = _apply_error_bar_style(
            safe_graph_name, infos, pname, error_bar_width, error_cap_width
        )

    grouping_note = ""
    if len(data_plots) > 1:
        # Can't reliably read Origin's group state over COM on this build, so
        # this is a static caveat (not a live check) whenever more than one
        # data plot exists — grouping only matters when there's more than one.
        grouping_note = (
            "; NOTE: if these plots are GROUPED, per-plot color/width/symbol "
            "changes may be overridden by the group — call ungroup_plots "
            "first if the style doesn't look right"
        )
    return (
        f"Updated style for plot {plot_index} ({pname}) in "
        f"{safe_graph_name}{error_note}{grouping_note}"
    )


def _apply_error_bar_style(
    safe_graph_name: str,
    infos: list,
    pname: str,
    error_bar_width: float | None,
    error_cap_width: float | None,
) -> str:
    """Apply -erw/-erwc to the error plot(s) adjacent to `pname` in `infos`
    order (P6-confirmed: error plots are adjacent to their data plot for both
    the y_error_col and set_error_bars construction routes). Falls back to
    ALL error plots on the layer when none are adjacent. Raises ValueError
    when the graph has no error plots at all. Returns a note to append to the
    caller's return message."""
    pos = next((i for i, p in enumerate(infos) if p["name"] == pname), None)
    adjacent = []
    if pos is not None:
        i = pos + 1
        while i < len(infos) and infos[i]["is_error"]:
            adjacent.append(infos[i]["name"])
            i += 1

    all_errors = [p["name"] for p in infos if p["is_error"]]
    if not all_errors:
        msg = (
            f"error_bar_width/error_cap_width requested but {safe_graph_name} "
            f"plot {pname} has no error bars — attach them first with "
            "set_error_bars or create_graph(..., y_error_col=...)."
        )
        raise ValueError(msg)

    targets = adjacent if adjacent else all_errors
    for err_name in targets:
        if error_bar_width is not None:
            graph_layer_execute(safe_graph_name, f"set {err_name} -erw {error_bar_width};")
        if error_cap_width is not None:
            graph_layer_execute(safe_graph_name, f"set {err_name} -erwc {error_cap_width};")
    import time
    time.sleep(0.2)

    if adjacent:
        return ""
    return (
        f"; NOTE: no error plot was directly adjacent to {pname} in plot "
        f"order, so error_bar_width/error_cap_width were applied to ALL "
        f"error plots on the layer ({', '.join(targets)})"
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
    removed = len(infos)
    for p in infos:
        graph_layer_execute(safe_graph, f"layer -e {p['name']};")
    rebuilt = 0
    failed_names = []
    for name in data_names:
        if graph_layer_execute(
            safe_graph,
            f"plotxy iy:={name} plot:={ptype} ogl:=[{safe_graph}]Layer1;",
        ):
            rebuilt += 1
        else:
            failed_names.append(name)
    if rebuilt:
        settle_new_plots(safe_graph, expected_min_plots=rebuilt)

    # Ungroup removes EVERY plot before re-plotting, so a rebuild failure leaves
    # the graph with fewer curves than it started with — that is never a
    # "success". Be honest about how many of the known datasets came back.
    if rebuilt == 0:
        msg = (
            f"Ungroup FAILED on {safe_graph}: removed {removed} plot(s) but could "
            f"NOT re-plot any of the {len(data_names)} data series "
            f"({', '.join(data_names)}). Layer1 may now be empty — recreate the "
            f"plots with create_graph / add_plot_to_graph."
        )
        raise ValueError(msg)
    if rebuilt < len(data_names):
        return (
            f"Ungroup PARTIAL on {safe_graph}: rebuilt {rebuilt} of "
            f"{len(data_names)} data series as independent {plot_type} plots, but "
            f"FAILED to re-plot {', '.join(failed_names)} — those series are now "
            f"MISSING from the graph (recreate them with add_plot_to_graph). Color "
            f"the rebuilt ones with set_plot_style."
        )
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

    NOTE: per-curve COLOR, WIDTH, and SYMBOL all apply only on UNGROUPED plots
    (probe-confirmed — a grouped plot's shared style increment overrides color
    AND width AND symbol, not just color). Graphs built with create_graph +
    add_plot_to_graph are ungrouped (fine). A grouped multi-curve plot (loaded
    from a project, or a single multi-Y plotxy) will NOT take the palette —
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

    # 1. Axis titles — bold, Arial 28pt. Written through the read-back-gated
    # helper: the old bare active-window writes silently dropped the labels on
    # the settle race (titles stayed the %(?X) column-name placeholder) while
    # the return still claimed "Arial bold labels". label_unverified collects
    # any axis whose title could not be confirmed so the return tells the truth.
    label_unverified: list[str] = []
    if x_label:
        validate_text_escapes(x_label, "x_label")
        if not _set_axis_title_verified(safe_graph_name, "xb", x_label):
            label_unverified.append("x")
    if y_label:
        validate_text_escapes(y_label, "y_label")
        if not _set_axis_title_verified(safe_graph_name, "yl", y_label):
            label_unverified.append("y")

    # Step returns are now checked: a FRAME or TICK failure raises (those are
    # structural — a "publication style" without a frame is misleading), while a
    # purely cosmetic failure (tick-label font, grid removal) is noted in the
    # return rather than aborting the whole call.
    cosmetic_notes: list[str] = []

    # 2. Tick labels — bold, Arial 22pt (cosmetic)
    if not graph_layer_execute(safe_graph_name, 'layer.x.label.pt = 22; layer.y.label.pt = 22;'):
        cosmetic_notes.append("tick-label size")
    if not graph_layer_execute(safe_graph_name, 'layer.x.label.bold = 1; layer.y.label.bold = 1;'):
        cosmetic_notes.append("tick-label bold")

    # 3. Axis range — TIGHT to the data by default (no empty gap before the
    # first or after the last point; if the data starts at 0 the axis starts
    # at 0) + readable tick spacing (~4 major intervals). Explicit min/max
    # override the data extent. Best-effort: unreadable data leaves the axis
    # on Origin's auto range.
    range_cmds = _tight_axis_cmds(safe_graph_name, x_min, x_max, y_min, y_max)
    if range_cmds:
        graph_layer_execute(safe_graph_name, " ".join(range_cmds))

    # 4. Ticks — inward, minor ticks on, proper lengths (structural → raise)
    if not graph_layer_execute(safe_graph_name,
        "layer.x.ticks = 1; layer.y.ticks = 1; "
        "layer.x.minor = 1; layer.y.minor = 1; "
        "layer.x.majorLen = 8; layer.y.majorLen = 8;"
    ):
        msg = f"Could not set inward ticks on {safe_graph_name}."
        raise ValueError(msg)

    # 5. Frame — closed (4 sides), thick (structural → raise)
    if not graph_layer_execute(safe_graph_name,
        "layer.x.opposite = 1; layer.y.opposite = 1; "
        "layer.x.thickness = 2; layer.y.thickness = 2;"
    ):
        msg = f"Could not close/thicken the frame on {safe_graph_name}."
        raise ValueError(msg)

    # 6. Remove grid lines (cosmetic)
    if not graph_layer_execute(safe_graph_name,
        "layer.x.grid = 0; layer.y.grid = 0; "
        "layer.x.minorGrid = 0; layer.y.minorGrid = 0;"
    ):
        cosmetic_notes.append("grid removal")

    # 7. Auto-style each data plot with a muted pastel palette + distinct
    # symbols. Error-bar plots only get the color of their data plot —
    # symbol/line commands would redraw them as connected lines.
    # P8 HARD RULE (probe-verified): every flag is its OWN `set <ds> -flag
    # val;` call. The previous code batched multiple `-flag`s into ONE `set`
    # command as a "settle" optimization — that batching is what silently
    # wiped colors to black on Origin 2020 (combining -c/-cf, or -k/-kf/-z,
    # in one command corrupts the plot). One settle after the LAST flag for
    # each plot is enough; never combine flags across a single `set` call.
    import time
    data_index = 0
    current_color = _rgb(PASTEL_RGB[0])
    for info in plot_infos:
        pname = info["name"]
        if info["is_error"]:
            graph_layer_execute(safe_graph_name, f"set {pname} -c {current_color};")
            graph_layer_execute(safe_graph_name, f"set {pname} -erw {_PUB_ERROR_BAR_WIDTH_PT};")
            graph_layer_execute(safe_graph_name, f"set {pname} -erwc {_PUB_ERROR_CAP_WIDTH};")
            time.sleep(0.2)
            continue
        current_color = _rgb(PASTEL_RGB[data_index % len(PASTEL_RGB)])
        shape = SYMBOL_SHAPES.get(data_index + 1, 2)
        data_index += 1

        line_width_units = int(_PUB_LINE_WIDTH_PT * _WIDTH_UNITS_PER_POINT)
        graph_layer_execute(safe_graph_name, f"set {pname} -c {current_color};")
        graph_layer_execute(safe_graph_name, f"set {pname} -w {line_width_units};")
        if _plot_has_symbols(pname):
            graph_layer_execute(safe_graph_name, f"set {pname} -k {shape};")
            graph_layer_execute(safe_graph_name, f"set {pname} -z {_PUB_SYMBOL_SIZE};")
        else:
            # bar/column/area-type plots: color the fill too, as its own call.
            graph_layer_execute(safe_graph_name, f"set {pname} -cf {current_color};")
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
    requested_labels = [a for a, lbl in (("x", x_label), ("y", y_label)) if lbl]
    if not requested_labels:
        labels_phrase = "Arial bold labels"
    elif label_unverified:
        labels_phrase = (
            f"Arial bold ticks (WARNING: the {', '.join(label_unverified)}-axis "
            f"label(s) could NOT be verified — Origin read them back unchanged, "
            f"so they were most likely dropped; set them again with "
            f"axis(op='labels'))"
        )
    else:
        labels_phrase = "Arial bold labels (verified)"
    cosmetic_note = (
        f"; NOTE: these cosmetic steps did not take: {', '.join(cosmetic_notes)}"
        if cosmetic_notes else ""
    )
    return (
        f"Publication style applied to {safe_graph_name}: "
        f"{data_index} data plots styled (pastel palette, {_PUB_LINE_WIDTH_PT} pt lines), "
        f"{labels_phrase}, inward ticks, closed frame, no grid, "
        f"borderless bold legend ({placement}){moved_out}{grouping_note}{cosmetic_note}"
    )


def _read_text_fsize(obj: str):
    """Read a text-object font size (e.g. ``xb.fsize``) back, or None if it
    can't be read. Used to confirm a title/legend font write actually landed —
    live-probed reliable on graphs loaded from a .opju once the page is
    activated (activate_window, which every caller does first)."""
    if not execute_labtalk(f"__mcp_fs = {obj}.fsize;"):
        return None
    try:
        return float(get_lt_var("__mcp_fs"))
    except Exception:
        return None


def _verify_text_fsize(obj: str, expected: int, label: str, graph_name: str) -> None:
    """Raise if ``obj``'s font size read back to a DIFFERENT non-zero value —
    turning a silent wrong-window / frozen no-op into a loud failure. A 0/None
    read-back (unreadable, e.g. test doubles) is skipped so this never
    false-alarms; live it reads the real size back."""
    got = _read_text_fsize(obj)
    if got and abs(got - expected) > 0.5:
        msg = (
            f"{label} font did not take on {graph_name} (set {expected}, read "
            f"back {got}). If this graph was loaded from a .opju, Origin may be "
            f"freezing its text objects over COM — recreate it in-session."
        )
        raise ValueError(msg)


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
        bold: Bold the targeted element(s). Axis titles and the graph title are
              bolded via `\\b(...)` markup; tick labels honor this flag directly
              (bold=False now leaves them un-bold, instead of the previous
              always-bold behavior).

    Note: for the "tick" (and "all") target, tick labels are sized at
    max(font_size - 4, 16) — a step smaller than the axis titles, the usual
    publication proportion; pass a larger font_size to enlarge them.

    On a graph LOADED from a .opju the axis-title/legend font writes are now
    read back (the size is confirmed) — live-verified reliable once the page is
    activated, and it raises instead of reporting a false success if the write
    is silently frozen.

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
        # Read-back gate: LabTalk returns success even if the write hit the wrong
        # window / a frozen loaded-graph text object, so confirm the size landed.
        _verify_text_fsize("xb", safe_font_size, "x-axis title", safe_graph_name)
        _verify_text_fsize("yl", safe_font_size, "y-axis title", safe_graph_name)

    if safe_target in ("all", "tick"):
        tick_size = max(safe_font_size - 4, 16)
        if not graph_layer_execute(safe_graph_name, f"layer.x.label.pt = {tick_size}; layer.y.label.pt = {tick_size};"):
            msg = f"Could not set tick label font size on {safe_graph_name}."
            raise ValueError(msg)
        tick_bold = 1 if bold else 0
        if not graph_layer_execute(
            safe_graph_name,
            f"layer.x.label.bold = {tick_bold}; layer.y.label.bold = {tick_bold};",
        ):
            msg = f"Could not set tick-label bold on {safe_graph_name}."
            raise ValueError(msg)

    if safe_target in ("all", "legend"):
        legend_size = max(safe_font_size - 4, 16)
        if not execute_labtalk(f"legend.font$ = {safe_font_name}; legend.fsize = {legend_size};"):
            msg = f"Could not set the legend font on {safe_graph_name}."
            raise ValueError(msg)
        _verify_text_fsize("legend", legend_size, "legend", safe_graph_name)

    if safe_target == "title":
        if not execute_labtalk(f"title.font$ = {safe_font_name}; title.fsize = {safe_font_size};"):
            msg = f"Could not set the title font on {safe_graph_name}."
            raise ValueError(msg)
        if bold:
            _bold_text_object("title")
        _verify_text_fsize("title", safe_font_size, "graph title", safe_graph_name)

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

    On a graph loaded from a .opju the legend rebuild is read back (the legend
    object's presence is confirmed) — live-verified reliable once the page is
    activated; a missing legend returns a WARNING rather than a clean success.

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
    # Read-back gate: confirm the legend object is readable after the rebuild.
    # legend -r reports success even when it silently did nothing on a frozen
    # loaded-graph, so an UNREADABLE legend must not read as a clean success
    # (live-probed reliable on a reloaded .opju, so this only fires on a genuine
    # freeze where the property can't be read back at all).
    if _read_text_fsize("legend") is None:
        return (
            f"Updated legend for {safe_graph_name}: placed {placement}{moved_out} "
            f"(WARNING: could not confirm the legend rendered — if it is missing, "
            f"this graph may be a frozen loaded .opju; rebuild it in-session)"
        )
    return f"Updated legend for {safe_graph_name}: placed {placement}{moved_out}"


# axis -> LabTalk axis-property prefix. "top"/"right" target ONLY the
# opposite-side border axis (x2/y2) — e.g. to strip its tick MARKS without
# touching the bottom/left axis's marks or number labels.
_TICK_AXIS_PREFIXES = {"x": ["x"], "y": ["y"], "both": ["x", "y"], "top": ["x2"], "right": ["y2"]}


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
        axis: "x", "y", "both" (default), "top", or "right" ("top"/"right"
              target only the opposite-side border axis, x2/y2)
        tick_direction: "in", "out", "both", or "none" ("none" removes that
              side's tick MARKS when present, while always leaving its number
              labels intact — this uses `layer.<ax>.ticks = 0`, NEVER
              `layer.<ax>.majorTicks`, which probe-confirmed wipes ALL axes'
              number labels on Origin 2020. A default closed frame may already
              show few/no marks on the opposite side, so the visible change can
              be subtle)
        major_length: Major tick length in points (default 8)
        minor_count: Number of minor ticks between major ticks (default 4)
        show_minor: Whether to show minor ticks

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    safe_axis = labtalk_choice(
        axis.lower(), {"x", "y", "both", "top", "right"}, "axis"
    )
    dir_map = {"in": 1, "out": 2, "both": 3, "none": 0}
    safe_tick_direction = labtalk_choice(tick_direction, dir_map, "tick_direction")
    d = dir_map[safe_tick_direction]

    minor = minor_count if show_minor else 0

    axes = _TICK_AXIS_PREFIXES[safe_axis]
    # NEVER emit `majorTicks` here — probe-confirmed on Origin 2020 to wipe
    # the NUMBER LABELS of all four axes. `ticks` is the label-safe knob for
    # tick-mark direction/removal.
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

# The perpendicular "distance from the axis" knob differs by axis orientation:
# the x (bottom) axis moves its labels VERTICALLY, the y (left) axis moves them
# HORIZONTALLY. Both are `layer.<ax>.label.offset<H|V>`, in % of the tick-label
# font size (the GUI's Axis dialog ▸ Tick Labels ▸ Format ▸ "Offset in % Point
# Size"), default 0. Probe-verified on Origin 2020: POSITIVE pulls the labels
# TOWARD the axis (shrinking the axis→label gap), negative pushes them away.
_TICK_LABEL_OFFSET_PROP = {"x": "offsetV", "y": "offsetH"}


@mcp.tool()
def set_tick_labels(
    graph_name: str,
    axis: str = "both",
    format: str | None = None,
    bold: bool | None = None,
    decimal_places: int | None = None,
    offset_pct: int | None = None,
) -> str:
    """Format an axis's tick labels: numeric format, bold, decimal places, offset.

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
        offset_pct: Perpendicular distance of the tick labels from the axis, in
            % of the tick-label font size (the GUI's "Offset in % Point Size").
            POSITIVE pulls the labels TOWARD the axis (smaller gap); negative
            pushes them away; 0 is Origin's default. Applied to each axis's
            perpendicular direction (x → vertical, y → horizontal), so it is the
            single knob for the axis→tick-label gap. None = leave untouched. Use
            a positive value to tighten Origin's default gap toward a
            matplotlib-style look.

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_axis = labtalk_choice(axis.lower(), {"x", "y", "both"}, "axis")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    if format is None and bold is None and decimal_places is None and offset_pct is None:
        msg = "Provide at least one of format, bold, decimal_places, or offset_pct."
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
        if offset_pct is not None:
            prop = _TICK_LABEL_OFFSET_PROP[a]
            cmds.append(f"layer.{a}.label.{prop} = {int(offset_pct)};")
    if format is not None:
        changed.append("format")
    if bold is not None:
        changed.append("bold")
    if decimal_places is not None:
        changed.append("decimal_places")
    if offset_pct is not None:
        changed.append("offset")
    if not graph_layer_execute(safe_graph, " ".join(cmds)):
        msg = f"Could not update tick labels ({', '.join(changed)}) for {safe_graph}."
        raise ValueError(msg)
    return f"Updated {safe_axis} tick labels for {safe_graph}: {', '.join(changed)}"
