import json
import math

from ..app import mcp
from ..origin_connection import (
    execute_labtalk,
    get_lt_str,
    get_origin,
    graph_names,
    matrix_names,
    require_matrix,
    require_worksheet,
)
from ..labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_text,
    positive_column,
    positive_int,
)

# matrix plot type -> Origin plot ID (verified live on Origin 2020)
_MATRIX_PLOT_TYPES = {
    "surface": 103,   # OpenGL 3D colormap surface
    "contour": 226,   # filled contour
    "heatmap": 105,
    "image": 220,
}
# Surface is OpenGL and must own its graph window.
_MATRIX_OWN_GRAPH = {"surface"}
# The colormapped types (surface/contour/heatmap) come from an Origin system
# template that carries a data-linked color scale (colorbar) by default. The
# "image" plot (type 220) is the exception: it renders a CONTINUOUS per-cell map
# (layer.cmap.numColors == 256) but has NO color scale, and one cannot be
# scripted onto it (every layer.colorscale.* / Colorscale.visible path no-ops —
# live-probed 2026-07-11). For a continuous map that STILL keeps a colorbar, use
# a heatmap + colormap(levels=64) — NOT an image plot. Verified live (Origin 2020).
_MATRIX_TEMPLATES = {
    "surface": "glcmap",   # OpenGL Z-colored surface + color scale
    "contour": "CONTOUR",  # filled contour + color scale
    "heatmap": "HeatMap",  # cell heatmap + color scale
    "image": "image",      # image plot — continuous, NO color scale
}

# Origin stores a missing/empty cell as either a LARGE sentinel (~1.2e308) or
# its small "unset" sentinel -1.23456789e-300 (live-confirmed: a fresh matrix
# AND a cell written NaN via PutMatrix both read back as -1.23456789e-300). The
# old `abs(v) >= 1e100` test only caught the large one, so empty/NaN cells
# silently read back as a tiny number instead of null.
_MISSING_MAGNITUDE = 1e100
_ORIGIN_MISSING_SENTINEL = -1.23456789e-300


def _is_missing_cell(v) -> bool:
    """True when a GetMatrix value represents an Origin missing/empty cell:
    NaN, the large (~1.2e308) sentinel, or the small -1.23456789e-300 sentinel.
    Real 0.0 is NOT flagged (isclose against the tiny sentinel fails for 0)."""
    if not isinstance(v, (int, float)):
        return False
    if isinstance(v, float) and math.isnan(v):
        return True
    if abs(v) >= _MISSING_MAGNITUDE:
        return True
    return math.isclose(v, _ORIGIN_MISSING_SENTINEL, rel_tol=1e-6)


def _create_matrix_book(name: str) -> str:
    """Create a matrix book and return its actual (uniquified) name."""
    # newbook auto-uniquifies the name on collision; read it back.
    execute_labtalk(f'newbook mat:=1 name:="{name}" option:=1;')
    return get_lt_str("page.name$")


@mcp.tool()
def create_matrix(book_name: str, rows: int = 10, cols: int = 10) -> str:
    """Create a new matrix book in Origin.

    Matrices back 3D surface, contour, heatmap, and image plots.

    Args:
        book_name: Name for the new matrix book
        rows: Initial number of rows (default 10)
        cols: Initial number of columns (default 10)

    Returns:
        JSON object: {"name": <actual matrix book name>, "requested_name":
        <book_name>, "renamed": <bool, true if Origin uniquified the name>,
        "rows": <rows>, "cols": <cols>}
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_rows = positive_int(rows, "rows")
    safe_cols = positive_int(cols, "cols")
    name = _create_matrix_book(safe_book)
    if not execute_labtalk(f"win -a {name}; mdim cols:={safe_cols} rows:={safe_rows};"):
        msg = f"Origin could not size matrix '{name}' to {safe_cols}x{safe_rows}."
        raise ValueError(msg)
    # Verify mdim took by reading the dims back (wks.nrows/ncols on the active
    # matrix — live-confirmed the correct read). Best-effort: skip when the
    # read is unavailable (0, e.g. test doubles) so this never false-alarms.
    o = get_origin()
    o.Execute(f"win -a {name};")
    got_rows = o.LTVar("wks.nrows")
    got_cols = o.LTVar("wks.ncols")
    if (got_rows and int(round(got_rows)) != safe_rows) or (
        got_cols and int(round(got_cols)) != safe_cols
    ):
        msg = (
            f"Matrix '{name}' size did not take: requested {safe_cols}x{safe_rows}, "
            f"read back {int(round(got_cols))}x{int(round(got_rows))}."
        )
        raise ValueError(msg)
    return json.dumps({
        "name": name,
        "requested_name": safe_book,
        "renamed": name != safe_book,
        "rows": safe_rows,
        "cols": safe_cols,
    })


@mcp.tool()
def set_matrix_data(book_name: str, data: str) -> str:
    """Write a 2D numeric grid into a matrix.

    Args:
        book_name: Matrix book name
        data: JSON array of equal-length rows, e.g. [[1,2,3],[4,5,6]]. A cell
              may be JSON null (or NaN) to write an Origin missing value —
              rendered as a hole in the colormap (heatmaps with missing cells
              are a real materials-science case). Round-trips back to null via
              get_matrix_data.

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    target = require_matrix(safe_book)
    try:
        grid = json.loads(data)
    except json.JSONDecodeError as exc:
        msg = (
            "data must be a JSON array of arrays, e.g. [[1,2,3],[4,5,6]]. "
            f"Parse error: {exc}"
        )
        raise ValueError(msg) from exc
    if (
        not isinstance(grid, list)
        or not grid
        or not all(isinstance(r, list) and r for r in grid)
    ):
        msg = "data must be a non-empty JSON array of non-empty rows."
        raise ValueError(msg)
    width = len(grid[0])
    if any(len(r) != width for r in grid):
        msg = "all rows must have the same length (a rectangular grid)."
        raise ValueError(msg)
    # null / NaN -> float('nan'): live-confirmed that PutMatrix accepts NaN and
    # stores Origin's missing sentinel (reads back as -1.23456789e-300, which
    # get_matrix_data maps to null). No per-cell LabTalk fallback is needed.
    float_grid = []
    for i, row in enumerate(grid):
        frow = []
        for v in row:
            if v is None or (isinstance(v, float) and math.isnan(v)):
                frow.append(float("nan"))
                continue
            try:
                frow.append(float(v))
            except (TypeError, ValueError) as exc:
                msg = f"Row {i + 1} contains non-numeric values; only numbers are supported."
                raise ValueError(msg) from exc
        float_grid.append(frow)
    if not get_origin().PutMatrix(target, float_grid):
        msg = f"Origin rejected the matrix data for {target}."
        raise ValueError(msg)
    return f"Set {len(float_grid)}x{width} matrix in {target}"


@mcp.tool()
def get_matrix_data(book_name: str) -> str:
    """Read a matrix back as JSON.

    Args:
        book_name: Matrix book name

    Returns:
        JSON object {"rows": [[...], ...]}; empty cells are null

    Raises:
        ValueError: if the matrix is not found.
    """
    safe_book = labtalk_name(book_name, "book_name")
    target = require_matrix(safe_book)
    data = get_origin().GetMatrix(target)
    if not isinstance(data, (list, tuple)):
        mats = ", ".join(matrix_names()) or "(none)"
        msg = f"Matrix {target} not found. Open matrices: {mats}."
        raise ValueError(msg)
    rows = [
        [None if _is_missing_cell(v) else v for v in row]
        for row in data
    ]
    return json.dumps({"rows": rows})


@mcp.tool()
def worksheet_to_matrix(
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    z_col: int,
    rows: int = 20,
    cols: int = 20,
    matrix_book: str = ""
) -> str:
    """Convert XYZ worksheet columns into a matrix by gridding (xyz2mat).

    Enables 3D surface / contour / heatmap from scattered XYZ data.

    Args:
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column (1-based)
        y_col: Y column (1-based)
        z_col: Z column (1-based)
        rows: Output matrix rows (default 20)
        cols: Output matrix columns (default 20)
        matrix_book: Optional name for the output matrix book

    Returns:
        JSON object: {"name": <actual matrix book name>, "requested_name":
        <matrix_book if given, else null>, "renamed": <bool, true if Origin
        uniquified the name>, "source": <"[data_book]data_sheet">, "rows":
        <rows>, "cols": <cols>}
    """
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x = positive_column(x_col, "x_col")
    safe_y = positive_column(y_col, "y_col")
    safe_z = positive_column(z_col, "z_col")
    safe_rows = positive_int(rows, "rows")
    safe_cols = positive_int(cols, "cols")
    require_worksheet(safe_book, safe_sheet)

    # xyz2mat needs the Z column designated as Z (type 6) on the active sheet.
    execute_labtalk(f'win -a {safe_book}; page.active$ = "{safe_sheet}"; wks.col{safe_z}.type = 6;')

    requested_name = labtalk_name(matrix_book, "matrix_book") if matrix_book else None
    name = _create_matrix_book(requested_name or "Matrix")
    cmd = (
        f"xyz2mat iz:=[{safe_book}]{safe_sheet}!({safe_x},{safe_y},{safe_z}) "
        f"settings.ConvertToMatrix.columns:={safe_cols} "
        f"settings.ConvertToMatrix.rows:={safe_rows} "
        f"om:=[{name}]MSheet1!;"
    )
    if not execute_labtalk(cmd):
        execute_labtalk(f"win -cd {name};")
        msg = (
            f"XYZ gridding failed for [{safe_book}]{safe_sheet}!"
            f"({safe_x},{safe_y},{safe_z}). Check that the columns contain data."
        )
        raise ValueError(msg)
    return json.dumps({
        "name": name,
        "requested_name": requested_name,
        "renamed": bool(requested_name) and name != requested_name,
        "source": f"[{safe_book}]{safe_sheet}",
        "rows": safe_rows,
        "cols": safe_cols,
    })


@mcp.tool()
def create_matrix_plot(
    matrix_book: str,
    plot_type: str = "heatmap",
    graph_name: str = "",
    z_label: str = ""
) -> str:
    """Plot a matrix as a surface, contour, heatmap, or image.

    Args:
        matrix_book: Matrix book name (see create_matrix / worksheet_to_matrix)
        plot_type: surface (3D), contour, heatmap, or image. surface/contour/
                 heatmap carry a color scale (colorbar); "image" is a continuous
                 per-cell map with NO colorbar (Origin cannot script one on). For
                 a continuous map that KEEPS a colorbar, use heatmap + colormap(
                 levels=64), not image (a default heatmap shows ~8 bands; levels
                 smooths it while retaining the color scale).
        graph_name: Optional name for the new graph
        z_label: Optional Z label with units (e.g. "Intensity (a.u.)"); sets
                 the matrix long name, which drives both the Z-axis title
                 (3D) and the color-scale title.

    Returns:
        JSON object: {"name": <actual graph name>, "requested_name":
        <graph_name if given, else null>, "renamed": <bool, true if the
        actual name differs from the requested name>, "plot_type": <plot_type>}
    """
    safe_book = labtalk_name(matrix_book, "matrix_book")
    safe_type = labtalk_choice(plot_type, _MATRIX_PLOT_TYPES, "plot_type")
    pid = _MATRIX_PLOT_TYPES[safe_type]
    target = require_matrix(safe_book)
    requested = labtalk_name(graph_name, "graph_name") if graph_name else None
    # The matrix long name flows into %(?Z) → Z-axis title + color scale.
    if z_label:
        execute_labtalk(
            f'win -a {safe_book}; wks.col1.lname$ = {labtalk_text(z_label, "z_label")};'
        )

    tmpl = _MATRIX_TEMPLATES[safe_type]
    # Surface is an OpenGL Z-colored mesh (no 2D plot id); its colormap comes
    # from the glcmap template. The 2D types take their plot id plus the matching
    # template; contour/heatmap supply a data-linked color scale, but "image"
    # does not (continuous map, no colorbar — see _MATRIX_TEMPLATES). Activate
    # the matrix first so plotm spawns a fresh graph window.
    plot_clause = "" if safe_type in _MATRIX_OWN_GRAPH else f"plot:={pid} "
    execute_labtalk(f"win -a {safe_book};")
    before = set(graph_names())
    if not execute_labtalk(
        f"plotm im:={target}! {plot_clause}ogl:=<new template:={tmpl}>;"
    ):
        msg = f"Could not plot matrix {target} as {safe_type}."
        raise ValueError(msg)
    new = set(graph_names()) - before
    name = new.pop() if new else get_lt_str("page.name$")
    if requested and name != requested and execute_labtalk(
        f"win -r {name} {requested};"
    ):
        name = requested
    return json.dumps({
        "name": name,
        "requested_name": requested,
        "renamed": bool(requested) and name != requested,
        "plot_type": safe_type,
    })
