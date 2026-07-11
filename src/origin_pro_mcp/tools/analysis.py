import json
import math

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


def _integrate_impl(data_book: str, data_sheet: str, x_col: int, y_col: int) -> str:
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


def _differentiate_impl(data_book: str, data_sheet: str, x_col: int, y_col: int) -> str:
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


def _smooth_impl(
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


def _interpolate_impl(
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


def _fft_impl(data_book: str, data_sheet: str, x_col: int, y_col: int) -> str:
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


def _find_peaks_impl(
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
    # Clamp the local-maximum window to the data length. pkfind's local-max test
    # compares each point to `npts` neighbors, so on a short spectrum a default
    # npts=10 leaves no point that can qualify — the search returns an
    # unactionable failure. Live-probed on an 11-point Gaussian: npts up to
    # (n-2)//2 (=4 here) find the peak; (n-1)//2 (=5) already fails — so clamp to
    # (n-2)//2.
    n_points = len(_read_column(safe_book, safe_sheet, sy))
    if n_points < 3:
        msg = (
            f"find_peaks needs at least 3 numeric points in column {sy} of "
            f"[{safe_book}]{safe_sheet}; found {n_points}."
        )
        raise ValueError(msg)
    max_npts = max(1, (n_points - 2) // 2)
    npts_used = min(safe_npts, max_npts)
    execute_labtalk(f"wks.col{sx}.type = 4; wks.col{sy}.type = 1;")
    cx = _ncols() + 1
    cy = cx + 1
    execute_labtalk("wks.addCol(); wks.addCol();")
    cmd = (
        f"pkfind iy:=({sx},{sy}) dir:={dir_map[safe_dir]} method:=max npts:={npts_used} "
        f"ocenter_x:=col({cx}) ocenter_y:=col({cy}) oleft:=<none> oright:=<none>;"
    )
    if not execute_labtalk(cmd):
        msg = f"Peak finding failed on [{safe_book}]{safe_sheet} ({sx},{sy})."
        raise ValueError(msg)
    xs = _read_column(safe_book, safe_sheet, cx)
    ys = _read_column(safe_book, safe_sheet, cy)
    peaks = [{"x": x, "y": y} for x, y in zip(xs, ys)]
    return json.dumps(
        {"peaks": peaks, "count": len(peaks), "local_points_used": npts_used}
    )


def _column_statistics_impl(data_book: str, data_sheet: str, col: int) -> str:
    """Descriptive statistics for one worksheet column.

    Returns:
        JSON: mean, sd, se, variance, median, min, max, sum, n
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_col = positive_column(col, "col")
    _activate_sheet(safe_book, safe_sheet)
    if not execute_labtalk(f"stats col({safe_col});"):
        msg = f"Statistics failed on [{safe_book}]{safe_sheet} col {safe_col}."
        raise ValueError(msg)
    mean = get_lt_var("stats.mean")
    sd = get_lt_var("stats.sd")
    n = get_lt_var("stats.n")
    execute_labtalk(f"__mcp_med = median(col({safe_col}));")
    median = get_lt_var("__mcp_med")
    se = sd / math.sqrt(n) if n > 0 else 0.0
    return json.dumps({
        "mean": mean,
        "sd": sd,
        "se": se,
        "variance": sd * sd,
        "median": median,
        "min": get_lt_var("stats.min"),
        "max": get_lt_var("stats.max"),
        "sum": get_lt_var("stats.sum"),
        "n": n,
    })


def _compare_means_impl(
    data_book: str,
    data_sheet: str,
    col1: int,
    col2: int,
    equal_variance: bool = False
) -> str:
    """Two-sample t-test between two columns.

    Returns:
        JSON: t, df, p_value, mean1, mean2, equal_variance
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    c1 = positive_column(col1, "col1")
    c2 = positive_column(col2, "col2")
    _activate_sheet(safe_book, safe_sheet)
    equal = 1 if equal_variance else 0
    if not execute_labtalk(
        f"ttest2 irng:=(col({c1}),col({c2})) tail:=two equal:={equal};"
    ):
        msg = f"t-test failed on [{safe_book}]{safe_sheet} cols {c1},{c2}."
        raise ValueError(msg)
    execute_labtalk(f"__m1 = mean(col({c1})); __m2 = mean(col({c2}));")
    return json.dumps({
        "t": get_lt_var("ttest2.stat"),
        "df": get_lt_var("ttest2.df"),
        "p_value": get_lt_var("ttest2.prob"),
        "mean1": get_lt_var("__m1"),
        "mean2": get_lt_var("__m2"),
        "equal_variance": bool(equal_variance),
    })


def _frequency_count_impl(
    data_book: str,
    data_sheet: str,
    col: int,
    bin_min: float,
    bin_max: float,
    bin_size: float
) -> str:
    """Histogram-style frequency counts for one column.

    Args:
        bin_min: lowest bin start
        bin_max: highest bin end
        bin_size: bin width (increment)

    Returns:
        JSON list of {center, end, count, cumulative}
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_col = positive_column(col, "col")
    if bin_size <= 0:
        msg = "bin_size must be positive."
        raise ValueError(msg)
    if bin_max <= bin_min:
        msg = "bin_max must be greater than bin_min."
        raise ValueError(msg)
    _activate_sheet(safe_book, safe_sheet)
    out = "FreqCount"
    if get_origin().FindWorksheet(f"[{safe_book}]{out}") is not None:
        execute_labtalk(f"win -a {safe_book}; layer -d {out};")
    cmd = (
        f"freqcounts irng:=col({safe_col}) min:={bin_min} max:={bin_max} "
        f"stepby:=0 inc:={bin_size} outleft:=1 outright:=1 rd:=[{safe_book}]{out}!;"
    )
    if not execute_labtalk(cmd):
        msg = f"Frequency count failed on [{safe_book}]{safe_sheet} col {safe_col}."
        raise ValueError(msg)
    centers = _read_column(safe_book, out, 1)
    ends = _read_column(safe_book, out, 2)
    counts = _read_column(safe_book, out, 3)
    cumulative = _read_column(safe_book, out, 4)
    bins = [
        {"center": c, "end": e, "count": n, "cumulative": cu}
        for c, e, n, cu in zip(centers, ends, counts, cumulative)
    ]
    return json.dumps({"sheet": f"[{safe_book}]{out}", "bins": bins})


# --- Consolidated dispatchers (Phase 2) ---------------------------------------

_TRANSFORM_METHODS = {
    "integrate", "differentiate", "smooth", "interpolate", "fft", "find_peaks"
}


@mcp.tool()
def transform(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    method: str,
    window_size: int | None = None,
    smooth_method: str | None = None,
    num_points: int | None = None,
    interp_method: str | None = None,
    direction: str | None = None,
    local_points: int | None = None,
) -> str:
    """Apply a numerical transform to an XY curve.

    Args:
        data_book, data_sheet: Source workbook and sheet.
        x_col, y_col: X and Y column numbers (1-based).
        method: Which transform to run:
            - "integrate": area under the curve. Returns JSON {area}.
            - "differentiate": dY/dX into a new column.
            - "smooth": smoothing into a new column; uses window_size
              (points, odd; default 5) and smooth_method (savitzky_golay
              (default), adjacent, or binomial).
            - "interpolate": resample onto evenly spaced X; uses num_points
              (default 100) and interp_method (linear/spline/bspline/akima,
              default linear). Returns JSON {sheet, x, y}.
            - "fft": forward FFT spectrum. Returns JSON {spectrum_sheet,
              dominant_frequency} (dominant_frequency is null when the
              spectrum columns cannot be read back).
            - "find_peaks": local-maximum peak search; uses direction
              (positive/negative/both, default positive) and local_points
              (neighborhood size, default 10). Returns JSON {peaks, count}.
        window_size: Smoothing window in points (method="smooth").
        smooth_method: savitzky_golay/adjacent/binomial (method="smooth";
            default savitzky_golay).
        num_points: Output point count (method="interpolate").
        interp_method: linear/spline/bspline/akima (method="interpolate").
        direction: positive/negative/both (method="find_peaks").
        local_points: Local-maximum neighborhood (method="find_peaks").

    Returns:
        The result of the selected transform (see method above).
    """
    safe_method = labtalk_choice(method.lower(), _TRANSFORM_METHODS, "method")
    if safe_method == "integrate":
        return _integrate_impl(data_book, data_sheet, x_col, y_col)
    if safe_method == "differentiate":
        return _differentiate_impl(data_book, data_sheet, x_col, y_col)
    if safe_method == "smooth":
        return _smooth_impl(
            data_book, data_sheet, x_col, y_col,
            method=smooth_method if smooth_method is not None else "savitzky_golay",
            window=window_size if window_size is not None else 5,
        )
    if safe_method == "interpolate":
        return _interpolate_impl(
            data_book, data_sheet, x_col, y_col,
            num_points=num_points if num_points is not None else 100,
            method=interp_method if interp_method is not None else "linear",
        )
    if safe_method == "fft":
        return _fft_impl(data_book, data_sheet, x_col, y_col)
    return _find_peaks_impl(
        data_book, data_sheet, x_col, y_col,
        direction=direction if direction is not None else "positive",
        local_points=local_points if local_points is not None else 10,
    )


_STATS_OPS = {"column", "compare_means", "frequency"}


@mcp.tool()
def stats(
    data_book: str,
    data_sheet: str,
    op: str,
    col: int,
    col2: int | None = None,
    bin_min: float | None = None,
    bin_max: float | None = None,
    bin_size: float | None = None,
    equal_variance: bool | None = None,
) -> str:
    """Compute statistics on worksheet columns.

    Args:
        data_book, data_sheet: Source workbook and sheet.
        op: Which statistic to run:
            - "column": descriptive stats for one column (uses col). Returns
              JSON {mean, sd, se, variance, median, min, max, sum, n}.
            - "compare_means": two-sample t-test between col and col2 (requires
              col2; equal_variance default False). Returns JSON {t, df,
              p_value, mean1, mean2, equal_variance}.
            - "frequency": histogram-style counts for col (requires bin_min,
              bin_max, bin_size). Returns JSON {sheet, bins}.
        col: Primary column (1-based).
        col2: Second column (op="compare_means").
        bin_min, bin_max, bin_size: Bin range/width (op="frequency").
        equal_variance: Assume equal variance (op="compare_means").

    Returns:
        The result of the selected statistic (see op above).
    """
    safe_op = labtalk_choice(op.lower(), _STATS_OPS, "op")
    if safe_op == "column":
        return _column_statistics_impl(data_book, data_sheet, col)
    if safe_op == "compare_means":
        if col2 is None:
            msg = "stats op 'compare_means' requires col2."
            raise ValueError(msg)
        return _compare_means_impl(
            data_book, data_sheet, col, col2,
            equal_variance if equal_variance is not None else False,
        )
    if bin_min is None or bin_max is None or bin_size is None:
        msg = "stats op 'frequency' requires bin_min, bin_max, and bin_size."
        raise ValueError(msg)
    return _frequency_count_impl(data_book, data_sheet, col, bin_min, bin_max, bin_size)
