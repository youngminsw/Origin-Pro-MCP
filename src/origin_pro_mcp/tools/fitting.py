import json
from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_lt_var,
    get_origin,
    require_graph,
    require_worksheet,
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

# NLFit version of the linear fit, used when the fit curve must be drawn
_LINE_NLFIT = ("Line", ("A", "B"))


def _read_tree_value(node: str):
    """Read one node of the NLFit result tree, or None if unreadable."""
    if not execute_labtalk(f"__mcpread = {_TREE}.{node};"):
        return None
    return get_lt_var("__mcpread")


def _book_sheet_names(book: str) -> set:
    pages = get_origin().WorksheetPages
    for i in range(pages.Count):
        page = pages.Item(i)
        if page.Name == book:
            layers = page.Layers
            return {layers.Item(j).Name for j in range(layers.Count)}
    return set()


@mcp.tool()
def curve_fit(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    function: str = "line",
    y_error_col: int = 0,
    plot_on_graph: str = ""
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

    Returns:
        JSON with fitted parameters (value + std_error) and statistics
        (r_squared, sum_sq_residuals, reduced_chi_sq, dof)
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_function = labtalk_choice(function, FITTING_FUNCTIONS, "function")
    sheet_ref = require_worksheet(safe_book, safe_sheet)
    safe_target_graph = ""
    if plot_on_graph:
        safe_target_graph = labtalk_name(plot_on_graph, "plot_on_graph")
        require_graph(safe_target_graph)

    # Set column designations so Origin knows which is X and which is Y.
    # Must use the active-sheet `wks` form — sheet-qualified assignments
    # like `[Book]Sheet!col(n).type = ...` are silently ignored.
    activate_window(safe_book, "data_book")
    execute_labtalk(f'page.active$ = "{safe_sheet}";')
    execute_labtalk(f"wks.col{safe_x_col}.type = 4;")  # 4 = X
    execute_labtalk(f"wks.col{safe_y_col}.type = 1;")  # 1 = Y

    result = {"function": safe_function, "statistics": {}}

    if safe_function == "line" and not safe_target_graph:
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

    if not execute_labtalk(f"nlbegin iy:={data_ref} func:={fdf_name} nltree:={_TREE};"):
        execute_labtalk("nlend;")
        msg = (
            f"Could not start the '{safe_function}' fit on {data_ref}. "
            "Check that the columns contain numeric data."
        )
        raise ValueError(msg)
    if not execute_labtalk("nlfit;"):
        execute_labtalk("nlend;")
        msg = f"The '{safe_function}' fit did not converge on {data_ref}."
        raise ValueError(msg)

    # Read everything from the result tree BEFORE nlend
    for key, node in _STAT_NODES.items():
        value = _read_tree_value(node)
        if value is not None:
            result["statistics"][key] = value

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
            params[p] = entry
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
                from .style_helpers import get_plot_names, reposition_legend_nearest_corner

                # `set` resolves plot names against the ACTIVE window —
                # plotxy with ogl:= does not activate the graph itself
                execute_labtalk(f"win -a {safe_target_graph};")
                curve_plots = get_plot_names(safe_target_graph)
                if curve_plots:
                    fit_plot = curve_plots[-1]
                    # muted brick red, distinct from the pastel palette
                    execute_labtalk(f"set {fit_plot} -c color(170,68,80);")
                    time.sleep(0.2)
                    execute_labtalk(f"set {fit_plot} -w 400;")  # 2 pt
                    time.sleep(0.2)
                # Adding a plot rebuilds the legend with a new entry, which
                # can push the box outside the frame — re-anchor it
                execute_labtalk(f"win -a {safe_target_graph}; legend -r;")
                time.sleep(0.3)
                reposition_legend_nearest_corner(safe_target_graph)
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
