import json
from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_lt_var,
    get_origin,
    require_graph,
    require_worksheet,
    sheet_names,
)
from ..labtalk_safe import labtalk_choice, labtalk_name, positive_column

# Tool-facing name -> (Origin FDF name, parameter node names in the NLFit
# result tree). FDF names and parameter names verified live on Origin Pro
# 2020 with ground-truth data (see git history); reading an unknown tree
# node silently returns a stale value, so only verified names are listed.
FITTING_FUNCTIONS = {
    "line": (None, None),  # handled by fitlr, not NLFit
    "poly2": ("Parabola", ("A", "B", "C")),
    "poly3": ("Cubic", ("A", "B", "C", "D")),
    "poly4": ("poly4", ("A0", "A1", "A2", "A3", "A4")),
    "poly5": ("poly5", ("A0", "A1", "A2", "A3", "A4", "A5")),
    "exp1": ("ExpDec1", ("y0", "A1", "t1")),
    "exp2": ("ExpDec2", ("y0", "A1", "t1", "A2", "t2")),
    "expgrow1": ("expgrow1", ("y0", "A1", "t1", "x0")),
    "expdecay1": ("expdecay1", ("y0", "A1", "t1", "x0")),
    "gauss": ("gauss", ("y0", "xc", "w", "A")),
    "lorentz": ("lorentz", ("y0", "xc", "w", "A")),
    "voigt": ("voigt", ("y0", "xc", "A", "wG", "wL")),
    # power's exponent is not readable from the result tree on Origin 2020
    "power": ("power", None),
    "lognormal": ("lognormal", ("y0", "xc", "w", "A")),
    "logistic": ("logistic", ("A1", "A2", "x0", "p")),
    "boltzmann": ("boltzmann", ("A1", "A2", "x0", "dx")),
    "hill": ("hill", ("Vmax", "k", "n")),
    "sine": ("sine", ("y0", "xc", "w", "A")),
}

# Result-tree statistic nodes (verified on Origin Pro 2020)
_STAT_NODES = {
    "r_squared": "cod",
    "sum_sq_residuals": "ssr",
    "reduced_chi_sq": "chisqr",
    "dof": "dof",
}

_TREE = "__mcpfit"

# NLFit version of the linear fit, used when the fit curve must be drawn.
# Origin's built-in "Line" FDF is y = A + B*x (A=intercept, B=slope), the
# same convention fitlr uses for fitlr.a/fitlr.b below — mapped back to
# "intercept"/"slope" so callers see identical parameter keys regardless
# of which fit path (fitlr vs NLFit) was used.
_LINE_NLFIT = ("Line", ("A", "B"))
_LINE_PARAM_MAP = {"A": "intercept", "B": "slope"}


# Origin's tiny "missing/unset" numeric sentinel (-1.23456789e-300): a
# non-converged replica fit freezes params at their init values and reads its
# std errors back as this value. Same magnitude heuristic used in labtalk.py /
# matrix.py — deep in the denormal regime, far below any real fit quantity.
_MISSING_SENTINEL_ABS_MAX = 1e-290


def _is_missing(value) -> bool:
    return isinstance(value, (int, float)) and 0 < abs(value) < _MISSING_SENTINEL_ABS_MAX


def _read_tree_value(node: str):
    """Read one node of the NLFit result tree, or None if unreadable."""
    if not execute_labtalk(f"__mcpread = {_TREE}.{node};"):
        return None
    return get_lt_var("__mcpread")


def _read_param_entry(node: str) -> dict:
    """{"value", "std_error"} for one result-tree parameter node, each key
    present only when it reads back as a real number (the missing-value
    sentinel of a frozen/non-converged fit is dropped, like the "power"
    unreadable precedent)."""
    entry = {}
    value = _read_tree_value(node)
    if value is not None and not _is_missing(value):
        entry["value"] = value
    std_error = _read_tree_value(f"e_{node}")
    if std_error is not None and not _is_missing(std_error):
        entry["std_error"] = std_error
    return entry


def _columns_by_longname(book: str, sheet: str):
    """(1-based index, LongName) for every column of [book]sheet, or [] if the
    sheet cannot be resolved. Used to locate the 'Cumulative Fit Peak' /
    'Fit Peak k' columns of a multi-peak fit-curve sheet by NAME (their column
    order is not assumed)."""
    ws = get_origin().FindWorksheet(f"[{book}]{sheet}")
    if ws is None:
        return []
    return [(c + 1, ws.Columns.Item(c).LongName) for c in range(ws.Columns.Count)]


def _plot_multipeak_fit(book, curve_sheet, target_graph, function, peaks):
    """Draw a multi-peak fit onto ``target_graph``: the 'Cumulative Fit Peak'
    envelope as the bold fit line, plus each 'Fit Peak k' component as a thin
    line (best-effort). Columns are found by LongName. Returns the ``fit_curve``
    dict, or a note dict when the cumulative column is absent/undrawable."""
    import time

    from .style_helpers import get_plot_names, place_legend_avoiding_data

    cols = _columns_by_longname(book, curve_sheet)
    cum_idx = next((i for i, ln in cols if ln == "Cumulative Fit Peak"), None)
    comp_idx = [i for i, ln in cols if ln.startswith("Fit Peak ")]
    if cum_idx is None:
        return {
            "fit_curve_note": (
                f"Multi-peak fit succeeded but [{book}]{curve_sheet} had no "
                "'Cumulative Fit Peak' column, so nothing was drawn."
            )
        }

    curve_ws = get_origin().FindWorksheet(f"[{book}]{curve_sheet}")
    if curve_ws is not None and cum_idx <= curve_ws.Columns.Count:
        curve_ws.Columns.Item(cum_idx - 1).LongName = f"\\b({function.capitalize()} fit)"

    if not execute_labtalk(
        f"plotxy iy:=[{book}]{curve_sheet}!(1,{cum_idx}) plot:=200 "
        f"ogl:=[{target_graph}]Layer1;"
    ):
        return {
            "fit_curve_note": (
                f"Could not draw the cumulative fit of [{book}]{curve_sheet} on "
                f"{target_graph}."
            )
        }

    # `set` resolves plot names against the ACTIVE window; plotxy ogl:= does not
    # activate the graph. Style the cumulative envelope brick-red 2 pt. P8 HARD
    # RULE: one flag per `set` call (batching corrupts the plot to black).
    execute_labtalk(f"win -a {target_graph};")
    curve_plots = get_plot_names(target_graph)
    if curve_plots:
        fit_plot = curve_plots[-1]
        execute_labtalk(f"set {fit_plot} -c color(170,68,80);")
        execute_labtalk(f"set {fit_plot} -w 400;")  # 2 pt
        time.sleep(0.2)

    # Per-peak component curves: thin muted grey, best-effort (a failure to draw
    # a component never fails the fit — the cumulative line is the deliverable).
    components_drawn = 0
    for order, idx in enumerate(comp_idx, start=1):
        if idx <= curve_ws.Columns.Count:
            curve_ws.Columns.Item(idx - 1).LongName = f"Peak {order}"
        if execute_labtalk(
            f"plotxy iy:=[{book}]{curve_sheet}!(1,{idx}) plot:=200 "
            f"ogl:=[{target_graph}]Layer1;"
        ):
            execute_labtalk(f"win -a {target_graph};")
            names = get_plot_names(target_graph)
            if names:
                execute_labtalk(f"set {names[-1]} -c color(150,150,150);")
                execute_labtalk(f"set {names[-1]} -w 100;")  # 0.5 pt
            components_drawn += 1
    time.sleep(0.2)

    # Adding plots rebuilds the legend with new entries — re-anchor it inside.
    execute_labtalk(f"win -a {target_graph}; legend -r;")
    time.sleep(0.3)
    place_legend_avoiding_data(target_graph)
    return {
        "fit_curve": {
            "sheet": f"[{book}]{curve_sheet}",
            "plotted_on": target_graph,
            "cumulative_column": cum_idx,
            "components_drawn": components_drawn,
        }
    }


def _book_sheet_names(book: str) -> set:
    """Sheet names of a workbook via the shared crash-safe LabTalk enumeration
    (never the deep page.Layers COM traversal that can crash heavy projects)."""
    return set(sheet_names(book))


def _rows_in_x_range(book: str, sheet: str, x_col: int, x_min, x_max):
    """1-based (first, last) row indices whose X value falls in [x_min, x_max]
    (either bound may be None = open). Used to build an NLFit row-subrange
    ``!(x,y)[i1:i2]`` — probe-verified to restrict the fit to those rows.
    Raises if no rows qualify. Assumes X is monotonic (spectra are); a
    contiguous [min..max] block is taken."""
    data = get_origin().GetWorksheet(f"[{book}]{sheet}")
    if not isinstance(data, (list, tuple)):
        msg = f"Could not read X column {x_col} of [{book}]{sheet} for the fit range."
        raise ValueError(msg)
    rows = []
    for r, row in enumerate(data, start=1):
        if x_col - 1 < len(row):
            v = row[x_col - 1]
            if isinstance(v, (int, float)) and abs(v) < 1e100:
                if (x_min is None or v >= x_min) and (x_max is None or v <= x_max):
                    rows.append(r)
    if not rows:
        lo = "-inf" if x_min is None else x_min
        hi = "+inf" if x_max is None else x_max
        msg = (
            f"No data rows of [{book}]{sheet} have X in [{lo}, {hi}]; "
            "the fit range excludes all points."
        )
        raise ValueError(msg)
    return min(rows), max(rows)


@mcp.tool()
def curve_fit(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    function: str = "line",
    y_error_col: int = 0,
    plot_on_graph: str = "",
    x_min: float | None = None,
    x_max: float | None = None,
    peaks: int = 1,
    peak_centers: str = ""
) -> str:
    """Perform curve fitting on worksheet data.

    Args:
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column number (1-based)
        y_col: Y column number (1-based)
        function: Fitting function. Built-in options:
                  line, poly2-5, exp1, exp2, expgrow1, expdecay1,
                  gauss, lorentz, voigt, power, lognormal, logistic,
                  boltzmann, hill, sine
        y_error_col: Y error column (1-based, 0=none)
        plot_on_graph: Optional name of an existing graph — the fitted
                       curve is drawn on it as a line (paper style:
                       data symbols + fit line). Also keeps the fit
                       report sheets in the workbook.
        x_min: Optional lower X bound — restrict the fit to rows whose X is
               >= x_min. Use with x_max to fit a single peak/region and
               exclude a baseline or a neighboring feature outside the range.
        x_max: Optional upper X bound — restrict the fit to rows whose X is
               <= x_max. Either bound may be given alone. The restriction is
               applied as a contiguous row block (X is assumed monotonic, as
               spectra are); a "line" fit with a range uses the NLFit engine.
        peaks: Number of peaks to fit simultaneously (deconvolution). Default 1
               (single-peak, unchanged). peaks>1 fits N overlapping peaks of the
               same function at once (XRD/XPS/Raman); ONLY valid for gauss,
               lorentz and voigt. Composes with x_min/x_max (restrict the range,
               then deconvolve within it).
        peak_centers: For peaks>1, optional "x1,x2,..." initial centre guesses,
               one per peak (count must equal peaks). When empty, Origin's
               built-in initialisation is used (works well on well-separated
               peaks; supply centres for tightly overlapping or noisy spectra).

    Returns:
        JSON with fitted parameters and statistics. Parameter keys are
        stable across calls for a given function (e.g. "line" always
        reports "intercept"/"slope", never "A"/"B"). Each parameter has
        a "value"; "std_error" is also present except for a "line" fit
        with plot_on_graph="" (the fast fitlr path does not expose fit
        std errors, only NLFit does). Statistics include r_squared always,
        plus sum_sq_residuals/reduced_chi_sq/dof for all functions except
        a "line" fit with plot_on_graph="". Special case: for
        function="power", the result has NO "parameters" key — Origin
        2020 cannot read the exponent back over COM, so only a
        "parameters_note" string is returned instead.
        For peaks>1 the shape differs: "peaks" (the count), a shared
        "baseline" {"y0": {...}}, and "parameters" keyed "peak_1".."peak_N",
        each a block of that function's per-peak parameters (gauss/lorentz:
        xc/w/A; voigt: xc/A/wG/wL) with value/std_error where readable. When
        plot_on_graph is given, the cumulative fit envelope is drawn as the fit
        line and each per-peak component as a thin line.
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_function = labtalk_choice(function, FITTING_FUNCTIONS, "function")

    # Multi-peak (deconvolution) validation — fail honestly BEFORE touching
    # Origin. peaks>1 uses the NLFit replica engine, which only makes sense for
    # peak functions.
    _MULTIPEAK_OK = ("gauss", "lorentz", "voigt")
    if peaks < 1:
        msg = f"curve_fit peaks ({peaks}) must be >= 1."
        raise ValueError(msg)
    is_multi = peaks > 1
    if is_multi and safe_function not in _MULTIPEAK_OK:
        msg = (
            f"curve_fit peaks>1 (multi-peak) is only supported for "
            f"{', '.join(_MULTIPEAK_OK)}, not '{safe_function}'."
        )
        raise ValueError(msg)
    center_values = []
    if peak_centers.strip():
        try:
            center_values = [float(c) for c in peak_centers.split(",") if c.strip()]
        except ValueError:
            msg = f"curve_fit peak_centers ({peak_centers!r}) must be comma-separated numbers."
            raise ValueError(msg) from None
        if not is_multi:
            msg = "curve_fit peak_centers only applies when peaks>1."
            raise ValueError(msg)
        if len(center_values) != peaks:
            msg = (
                f"curve_fit peak_centers has {len(center_values)} value(s) but "
                f"peaks={peaks}; supply exactly one centre per peak."
            )
            raise ValueError(msg)

    sheet_ref = require_worksheet(safe_book, safe_sheet)
    safe_target_graph = ""
    if plot_on_graph:
        safe_target_graph = labtalk_name(plot_on_graph, "plot_on_graph")
        require_graph(safe_target_graph)

    if x_min is not None and x_max is not None and x_min >= x_max:
        msg = f"curve_fit x_min ({x_min}) must be < x_max ({x_max})."
        raise ValueError(msg)
    # Resolve an optional X-range restriction to a 1-based row subrange. The
    # bracket ``!(x,y)[i1:i2]`` is probe-verified to restrict the NLFit input.
    row_subrange = ""
    if x_min is not None or x_max is not None:
        i1, i2 = _rows_in_x_range(safe_book, safe_sheet, safe_x_col, x_min, x_max)
        row_subrange = f"[{i1}:{i2}]"

    # Set column designations so Origin knows which is X and which is Y.
    # Must use the active-sheet `wks` form — sheet-qualified assignments
    # like `[Book]Sheet!col(n).type = ...` are silently ignored.
    activate_window(safe_book, "data_book")
    execute_labtalk(f'page.active$ = "{safe_sheet}";')
    execute_labtalk(f"wks.col{safe_x_col}.type = 4;")  # 4 = X
    execute_labtalk(f"wks.col{safe_y_col}.type = 1;")  # 1 = Y

    result = {"function": safe_function, "statistics": {}}

    if safe_function == "line" and not safe_target_graph and not row_subrange:
        # Use fitlr for linear regression (fast, no GUI)
        if not execute_labtalk(f"fitlr {sheet_ref}!col({safe_y_col});"):
            msg = (
                f"Linear fit failed on {sheet_ref}!col({safe_y_col}). "
                "Check that the X and Y columns contain numeric data."
            )
            raise ValueError(msg)
        r = get_lt_var("fitlr.r")
        result["statistics"]["r_squared"] = r * r
        result["parameters"] = {
            "intercept": {"value": get_lt_var("fitlr.a")},
            "slope": {"value": get_lt_var("fitlr.b")},
        }
        return json.dumps(result, indent=2)

    if safe_function == "line":
        # Drawing the fit curve needs the NLFit engine
        fdf_name, param_names = _LINE_NLFIT
    else:
        fdf_name, param_names = FITTING_FUNCTIONS[safe_function]
    data_ref = f"{sheet_ref}!({safe_x_col},{safe_y_col})"
    if y_error_col > 0:
        safe_error_col = positive_column(y_error_col, "y_error_col")
        data_ref = f"{sheet_ref}!({safe_x_col},{safe_y_col},{safe_error_col})"
    # Restrict to the resolved row block (empty when no x_min/x_max given).
    data_ref += row_subrange

    # peaks>1: N overlapping copies of the peak function. replica:=peaks-1
    # (probe-verified: replica is the count of ADDITIONAL copies).
    replica_opt = f" replica:={peaks - 1}" if is_multi else ""
    if not execute_labtalk(
        f"nlbegin iy:={data_ref} func:={fdf_name}{replica_opt} nltree:={_TREE};"
    ):
        execute_labtalk("nlend;")
        msg = (
            f"Could not start the '{safe_function}' fit on {data_ref}. "
            "Check that the columns contain numeric data."
        )
        raise ValueError(msg)
    # Seed the per-peak centres when supplied (peak 1 = unsuffixed node `xc`,
    # peaks 2..N = `xc__k`). Omitted => Origin's built-in replica init.
    if is_multi and center_values:
        for k, c in enumerate(center_values, start=1):
            node = "xc" if k == 1 else f"xc__{k}"
            execute_labtalk(f"{_TREE}.{node} = {c};")
    if not execute_labtalk("nlfit;"):
        execute_labtalk("nlend;")
        msg = f"The '{safe_function}' fit did not converge on {data_ref}."
        raise ValueError(msg)

    # Read everything from the result tree BEFORE nlend
    for key, node in _STAT_NODES.items():
        value = _read_tree_value(node)
        if value is not None:
            result["statistics"][key] = value

    # A non-converged replica fit leaves nlfit reporting success but freezes the
    # params at their init with cod (R^2) == 0 (probe-verified). Guard it so the
    # tool never returns a frozen fit as a real one.
    if is_multi:
        cod = result["statistics"].get("r_squared")
        if cod is None or cod <= 0:
            execute_labtalk("nlend;")
            msg = (
                f"The {peaks}-peak '{safe_function}' fit did not converge on "
                f"{data_ref} (R^2 <= 0). Supply peak_centers close to the real "
                "peak positions, or reduce peaks."
            )
            raise ValueError(msg)
        # Shared baseline + one block per peak. Per-peak params are the function's
        # params after the shared y0 (param_names[0]); reads bounded to `peaks`
        # (reading a node beyond the last peak silently returns a stale value).
        per_peak_bases = param_names[1:]
        y0_entry = _read_param_entry(param_names[0])
        if y0_entry:
            result["baseline"] = {param_names[0]: y0_entry}
        peak_blocks = {}
        for k in range(1, peaks + 1):
            block = {}
            for base in per_peak_bases:
                node = base if k == 1 else f"{base}__{k}"
                entry = _read_param_entry(node)
                if entry:
                    block[base] = entry
            peak_blocks[f"peak_{k}"] = block
        result["peaks"] = peaks
        result["parameters"] = peak_blocks
        if safe_target_graph:
            sheets_before = _book_sheet_names(safe_book)
            execute_labtalk("nlend output:=1;")
            new_sheets = _book_sheet_names(safe_book) - sheets_before
            curve_sheets = [s for s in new_sheets if "Curve" in s]
            if not curve_sheets:
                result["fit_curve_note"] = (
                    "Fit succeeded but Origin produced no fit-curve sheet, "
                    "so nothing was drawn on the graph."
                )
            else:
                result.update(
                    _plot_multipeak_fit(
                        safe_book, curve_sheets[0], safe_target_graph,
                        safe_function, peaks,
                    )
                )
        else:
            execute_labtalk("nlend;")
        return json.dumps(result, indent=2)

    if param_names is None:
        result["parameters_note"] = (
            f"Parameter values for '{safe_function}' cannot be read back "
            "over COM on Origin 2020; see the fit curve and report in Origin."
        )
    else:
        params = {}
        for p in param_names:
            entry = {}
            value = _read_tree_value(p)
            if value is not None:
                entry["value"] = value
            std_error = _read_tree_value(f"e_{p}")
            if std_error is not None:
                entry["std_error"] = std_error
            key = _LINE_PARAM_MAP[p] if safe_function == "line" else p
            params[key] = entry
        result["parameters"] = params

    if safe_target_graph:
        sheets_before = _book_sheet_names(safe_book)
        # output:=1 writes the fit report + fitted curve sheets
        execute_labtalk("nlend output:=1;")
        new_sheets = _book_sheet_names(safe_book) - sheets_before
        curve_sheets = [s for s in new_sheets if "Curve" in s]
        if not curve_sheets:
            result["fit_curve_note"] = (
                "Fit succeeded but Origin produced no fit-curve sheet, "
                "so nothing was drawn on the graph."
            )
        else:
            curve_sheet = curve_sheets[0]
            # Short legend text instead of Origin's verbose default
            # ("Gauss Fit of B"Intensity"), which overflows the legend box
            curve_ws = get_origin().FindWorksheet(f"[{safe_book}]{curve_sheet}")
            if curve_ws is not None and curve_ws.Columns.Count > 1:
                curve_ws.Columns.Item(1).LongName = f"\\b({safe_function.capitalize()} fit)"
            plotted = execute_labtalk(
                f"plotxy iy:=[{safe_book}]{curve_sheet}!(1,2) plot:=200 "
                f"ogl:=[{safe_target_graph}]Layer1;"
            )
            if plotted:
                import time
                from .style_helpers import get_plot_names, place_legend_avoiding_data

                # `set` resolves plot names against the ACTIVE window —
                # plotxy with ogl:= does not activate the graph itself
                execute_labtalk(f"win -a {safe_target_graph};")
                curve_plots = get_plot_names(safe_target_graph)
                if curve_plots:
                    fit_plot = curve_plots[-1]
                    # muted brick red, distinct from the pastel palette. P8
                    # HARD RULE (probe-verified, Origin 2020): every flag is its
                    # OWN `set <ds> -flag val;` call — combining `-c` and `-w`
                    # (or any two flags) in one command silently corrupts the
                    # plot (wipes the color to black). One settle after the last.
                    execute_labtalk(f"set {fit_plot} -c color(170,68,80);")
                    execute_labtalk(f"set {fit_plot} -w 400;")  # 2 pt
                    time.sleep(0.2)
                # Adding a plot rebuilds the legend with a new entry, which
                # can push the box outside the frame — re-anchor it
                execute_labtalk(f"win -a {safe_target_graph}; legend -r;")
                time.sleep(0.3)
                place_legend_avoiding_data(safe_target_graph)
                result["fit_curve"] = {
                    "sheet": f"[{safe_book}]{curve_sheet}",
                    "plotted_on": safe_target_graph,
                }
            else:
                result["fit_curve_note"] = (
                    f"Could not draw [{safe_book}]{curve_sheet} on "
                    f"{safe_target_graph}."
                )
    else:
        execute_labtalk("nlend;")
    return json.dumps(result, indent=2)

@mcp.tool()
def list_fitting_functions() -> str:
    """List available built-in fitting functions and their parameters.

    Returns:
        JSON of function names grouped by category, with parameter names
    """
    functions = {
        "linear": {"line": ["intercept", "slope"]},
        "polynomial": {
            name: list(FITTING_FUNCTIONS[name][1])
            for name in ("poly2", "poly3", "poly4", "poly5")
        },
        "exponential": {
            name: list(FITTING_FUNCTIONS[name][1])
            for name in ("exp1", "exp2", "expgrow1", "expdecay1")
        },
        "peak": {
            name: list(FITTING_FUNCTIONS[name][1])
            for name in ("gauss", "lorentz", "voigt")
        },
        "growth_sigmoidal": {
            name: list(FITTING_FUNCTIONS[name][1])
            for name in ("boltzmann", "hill", "logistic", "lognormal")
        },
        "other": {
            "power": "fits, but parameters are not readable over COM",
            "sine": list(FITTING_FUNCTIONS["sine"][1]),
        },
    }
    return json.dumps(functions, indent=2)
