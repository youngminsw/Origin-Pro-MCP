import json

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_lt_str,
    get_lt_var,
    get_origin,
    require_worksheet,
)
from ..labtalk_safe import labtalk_choice, labtalk_name, positive_column, positive_int

_MISSING_MAGNITUDE = 1e100
_SMOOTH_METHODS = {"sg": 1, "savitzky_golay": 1, "adjacent": 2, "average": 2, "binomial": 3}


def _activate_sheet(book: str, sheet: str) -> None:
    require_worksheet(book, sheet)
    activate_window(book, "data_book")
    execute_labtalk(f'page.active$ = "{sheet}";')


def _ncols() -> int:
    return int(get_lt_var("wks.ncols"))


def _read_column(book: str, sheet: str, col: int) -> list:
    """Read one column's real (non-missing) numeric values."""
    data = get_origin().GetWorksheet(f"[{book}]{sheet}")
    if not isinstance(data, (list, tuple)):
        return []
    vals = []
    for row in data:
        if col - 1 < len(row):
            v = row[col - 1]
            if isinstance(v, (int, float)) and abs(v) < _MISSING_MAGNITUDE:
                vals.append(float(v))
    return vals


@mcp.tool()
def integrate(data_book: str, data_sheet: str, x_col: int, y_col: int) -> str:
    """Integrate Y over X (area under the curve).

    Returns:
        JSON with the integrated area
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    _activate_sheet(safe_book, safe_sheet)
    execute_labtalk(f"wks.col{sx}.type = 4; wks.col{sy}.type = 1;")
    if not execute_labtalk(f"integ1 iy:=({sx},{sy}) type:=math;"):
        msg = f"Integration failed on [{safe_book}]{safe_sheet} ({sx},{sy})."
        raise ValueError(msg)
    return json.dumps({"area": get_lt_var("integ1.area")})


@mcp.tool()
def differentiate(data_book: str, data_sheet: str, x_col: int, y_col: int) -> str:
    """Compute the derivative dY/dX into a new column.

    Returns:
        Success message naming the new derivative column
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    _activate_sheet(safe_book, safe_sheet)
    execute_labtalk(f"wks.col{sx}.type = 4; wks.col{sy}.type = 1;")
    out = _ncols() + 1
    execute_labtalk("wks.addCol();")
    if not execute_labtalk(f"differentiate iy:=({sx},{sy}) oy:=({sx},{out});"):
        msg = f"Differentiation failed on [{safe_book}]{safe_sheet} ({sx},{sy})."
        raise ValueError(msg)
    execute_labtalk(f'wks.col{out}.lname$ = "dY/dX";')
    return f"Derivative written to column {out} of [{safe_book}]{safe_sheet}"


@mcp.tool()
def smooth(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    method: str = "savitzky_golay",
    window: int = 5
) -> str:
    """Smooth a curve into a new column.

    Args:
        method: savitzky_golay (sg), adjacent (moving average), or binomial
        window: number of points in the smoothing window (odd, default 5)

    Returns:
        Success message naming the new smoothed column
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    safe_method = labtalk_choice(method.lower(), _SMOOTH_METHODS, "method")
    safe_window = positive_int(window, "window")
    _activate_sheet(safe_book, safe_sheet)
    execute_labtalk(f"wks.col{sx}.type = 4; wks.col{sy}.type = 1;")
    out = _ncols() + 1
    execute_labtalk("wks.addCol();")
    mid = _SMOOTH_METHODS[safe_method]
    if not execute_labtalk(
        f"smooth iy:=({sx},{sy}) method:={mid} npts:={safe_window} oy:=({sx},{out});"
    ):
        msg = f"Smoothing failed on [{safe_book}]{safe_sheet} ({sx},{sy})."
        raise ValueError(msg)
    execute_labtalk(f'wks.col{out}.lname$ = "Smoothed";')
    return f"Smoothed curve ({safe_method}, {safe_window} pts) written to column {out}"


@mcp.tool()
def interpolate(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    num_points: int = 100,
    method: str = "linear"
) -> str:
    """Resample an XY curve onto evenly spaced X values.

    Args:
        num_points: number of output points across the X range
        method: linear, spline, bspline, or akima

    Returns:
        JSON: output sheet plus the resampled X and Y
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    safe_n = positive_int(num_points, "num_points")
    safe_method = labtalk_choice(
        method.lower(), {"linear", "spline", "bspline", "akima"}, "method"
    )
    require_worksheet(safe_book, safe_sheet)
    xvals = _read_column(safe_book, safe_sheet, sx)
    if len(xvals) < 2:
        msg = f"Column {sx} of [{safe_book}]{safe_sheet} has too few X values."
        raise ValueError(msg)
    x_min, x_max = min(xvals), max(xvals)
    step = (x_max - x_min) / (safe_n - 1) if safe_n > 1 else 0
    o = get_origin()
    out_name = o.CreatePage(2, "Interp", "origin")
    execute_labtalk(f'win -a {out_name}; page.active$ = "Sheet1"; col(1) = data({x_min}, {x_max}, {step});')
    cmd = (
        f"interp1 ix:=[{out_name}]Sheet1!col(1) "
        f"iy:=[{safe_book}]{safe_sheet}!({sx},{sy}) "
        f"ox:=[{out_name}]Sheet1!col(2) method:={safe_method};"
    )
    if not execute_labtalk(cmd):
        execute_labtalk(f"win -cd {out_name};")
        msg = f"Interpolation failed on [{safe_book}]{safe_sheet} ({sx},{sy})."
        raise ValueError(msg)
    new_x = _read_column(out_name, "Sheet1", 1)
    new_y = _read_column(out_name, "Sheet1", 2)
    return json.dumps({"sheet": f"[{out_name}]Sheet1", "x": new_x, "y": new_y})


@mcp.tool()
def fft(data_book: str, data_sheet: str, x_col: int, y_col: int) -> str:
    """Forward FFT of a signal; outputs a spectrum sheet.

    Uses the X column spacing as the sampling interval, so the Frequency
    column is in real units (1 / X-unit).

    Returns:
        JSON: spectrum sheet and the dominant (peak-amplitude) frequency
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    _activate_sheet(safe_book, safe_sheet)
    execute_labtalk(f"wks.col{sx}.type = 4; wks.col{sy}.type = 1;")
    out_sheet = "FFTResult"
    execute_labtalk(f"win -a {safe_book};")
    if not execute_labtalk(f"fft1 ix:=col({sy}) rd:=[{safe_book}]{out_sheet}!;"):
        msg = f"FFT failed on [{safe_book}]{safe_sheet} col {sy}."
        raise ValueError(msg)
    # Identify Frequency and Amplitude columns by long name.
    freq = _find_named_column(safe_book, out_sheet, "freq")
    amp = _find_named_column(safe_book, out_sheet, "amplitude") or _find_named_column(
        safe_book, out_sheet, "magnitude"
    )
    dominant = None
    if freq and amp:
        fvals = _read_column(safe_book, out_sheet, freq)
        avals = _read_column(safe_book, out_sheet, amp)
        if fvals and avals:
            n = min(len(fvals), len(avals))
            # skip DC (index 0)
            peak_i = max(range(1, n), key=lambda i: avals[i]) if n > 1 else 0
            dominant = fvals[peak_i]
    return json.dumps(
        {"spectrum_sheet": f"[{safe_book}]{out_sheet}", "dominant_frequency": dominant}
    )


def _find_named_column(book: str, sheet: str, keyword: str):
    """1-based index of the first column whose long name contains keyword."""
    execute_labtalk(f'win -a {book}; page.active$ = "{sheet}";')
    n = _ncols()
    for i in range(1, n + 1):
        lname = get_lt_str(f"wks.col{i}.lname$").lower()
        if keyword in lname:
            return i
    return None


@mcp.tool()
def find_peaks(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    direction: str = "positive",
    local_points: int = 10
) -> str:
    """Find peaks in an XY curve (local-maximum method).

    Args:
        direction: positive, negative, or both
        local_points: neighborhood size for the local-maximum search

    Returns:
        JSON list of {x, y} peak positions
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    dir_map = {"positive": "p", "negative": "n", "both": "b"}
    safe_dir = labtalk_choice(direction.lower(), dir_map, "direction")
    safe_npts = positive_int(local_points, "local_points")
    _activate_sheet(safe_book, safe_sheet)
    execute_labtalk(f"wks.col{sx}.type = 4; wks.col{sy}.type = 1;")
    cx = _ncols() + 1
    cy = cx + 1
    execute_labtalk("wks.addCol(); wks.addCol();")
    cmd = (
        f"pkfind iy:=({sx},{sy}) dir:={dir_map[safe_dir]} method:=max npts:={safe_npts} "
        f"ocenter_x:=col({cx}) ocenter_y:=col({cy}) oleft:=<none> oright:=<none>;"
    )
    if not execute_labtalk(cmd):
        msg = f"Peak finding failed on [{safe_book}]{safe_sheet} ({sx},{sy})."
        raise ValueError(msg)
    xs = _read_column(safe_book, safe_sheet, cx)
    ys = _read_column(safe_book, safe_sheet, cy)
    peaks = [{"x": x, "y": y} for x, y in zip(xs, ys)]
    return json.dumps({"peaks": peaks, "count": len(peaks)})
