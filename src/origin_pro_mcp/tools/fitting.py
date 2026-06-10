import json
from ..app import mcp
from ..origin_connection import execute_labtalk, get_lt_var
from ..labtalk_safe import labtalk_choice, labtalk_name, positive_column

FITTING_FUNCTIONS = {
    "line",
    "poly2",
    "poly3",
    "poly4",
    "poly5",
    "exp1",
    "exp2",
    "expgrow1",
    "expdecay1",
    "gauss",
    "lorentz",
    "voigt",
    "power",
    "lognormal",
    "logistic",
    "boltzmann",
    "hill",
    "sine",
}

@mcp.tool()
def curve_fit(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    function: str = "line",
    y_error_col: int = 0
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

    Returns:
        JSON with fit parameters and statistics (R², SSR, etc.)
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_function = labtalk_choice(function, FITTING_FUNCTIONS, "function")
    sheet_ref = f"[{safe_book}]{safe_sheet}"

    # Set column designations so Origin knows which is X and which is Y
    execute_labtalk(f"win -a {safe_book};")
    execute_labtalk(f"{sheet_ref}!col({safe_x_col}).type = 4;")  # 4 = X
    execute_labtalk(f"{sheet_ref}!col({safe_y_col}).type = 1;")  # 1 = Y

    result = {"function": safe_function, "statistics": {}}

    if safe_function == "line":
        # Use fitlr for linear regression (fast, no GUI)
        execute_labtalk(f"fitlr {sheet_ref}!col({safe_y_col});")
        r = get_lt_var("fitlr.r")
        result["statistics"]["r_squared"] = r * r
        result["parameters"] = {
            "intercept": get_lt_var("fitlr.a"),
            "slope": get_lt_var("fitlr.b"),
        }
    else:
        data_ref = f"{sheet_ref}!({safe_x_col},{safe_y_col})"
        if y_error_col > 0:
            safe_error_col = positive_column(y_error_col, "y_error_col")
            data_ref = f"{sheet_ref}!({safe_x_col},{safe_y_col},{safe_error_col})"

        execute_labtalk(f"nlbegin iy:={data_ref} func:={safe_function};")
        execute_labtalk("nlfit;")

        # Read statistics BEFORE nlend
        try:
            result["statistics"]["r_squared"] = get_lt_var("nlr.r2")
        except:
            pass
        try:
            result["statistics"]["sum_sq_residuals"] = get_lt_var("nlr.ssr")
        except:
            pass
        try:
            result["statistics"]["dof"] = get_lt_var("nlr.dof")
        except:
            pass

        execute_labtalk("nlend;")

    return json.dumps(result, indent=2)

@mcp.tool()
def list_fitting_functions() -> str:
    """List available built-in fitting functions in Origin.

    Returns:
        List of function names grouped by category
    """
    functions = {
        "linear": ["line"],
        "polynomial": ["poly2", "poly3", "poly4", "poly5"],
        "exponential": ["exp1", "exp2", "expgrow1", "expdecay1"],
        "peak": ["gauss", "lorentz", "voigt"],
        "growth_sigmoidal": ["boltzmann", "hill", "logistic", "lognormal"],
        "other": ["power", "sine"],
    }
    return json.dumps(functions, indent=2)
