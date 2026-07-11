import json
import os
import shutil
import tempfile
import uuid

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_lt_str,
    get_origin,
    graph_names,
    require_graph,
    require_worksheet,
)
from ..labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_text,
    validate_text_escapes,
    positive_column,
    positive_int,
    windows_path,
)
from .style_helpers import (
    acquire_graph_layer,
    get_plot_info,
    graph_layer_execute,
    read_layer_value,
    settle_new_plots,
    verify_layer_value,
)

# Plot type IDs verified against OriginLab's "Plot Type IDs" reference and
# tested live on Origin Pro 2020. The earlier values for area/bar/box/
# histogram/pie/contour/3d were wrong (off-by-template), producing empty
# or mis-typed graphs.
PLOT_TYPES = {
    "scatter": 201,
    "line": 200,
    "line+symbol": 202,
    "column": 203,
    "bar": 215,          # horizontal bar (204 is Area)
    "area": 204,
    "pie": 225,
    "histogram": 219,    # Y range
    "box": 206,          # Y range
    "contour": 243,      # XYZ range
    "3d_scatter": 240,   # XYZ range (OpenGL, owns its graph)
}

# Take a single Y column instead of an X,Y pair.
_Y_ONLY_TYPES = {"histogram", "box"}
# Need X, Y, Z columns and are drawn with plotxyz.
_XYZ_TYPES = {"contour", "3d_scatter"}
# OpenGL 3D types that must own their graph window — plotting them into a
# pre-made 2D page collapses them to an empty/flat projection.
_OWN_GRAPH_TYPES = {"3d_scatter"}

# Formats expGraph can write on Origin 2020 (probe-verified 2026-07-11).
# Raster formats take a pixel size; vector formats are resolution-independent
# and IGNORE the pixel width/height. SVG proved out as UNSUPPORTED on this
# build (expGraph type:=svg fails and writes no file), so it is not offered.
EXPORT_RASTER_FORMATS = {"png", "jpg", "tif", "bmp"}
EXPORT_VECTOR_FORMATS = {"pdf", "eps", "emf"}
EXPORT_FORMATS = EXPORT_RASTER_FORMATS | EXPORT_VECTOR_FORMATS
# Back-compat alias (was raster-only); kept so external importers don't break.
EXPORT_IMAGE_FORMATS = EXPORT_RASTER_FORMATS


def _designate_error_column(book: str, sheet: str, col: int) -> None:
    """Mark a column as Y Error so styling tools can recognize error-bar
    plots later. Must use the active-sheet `wks` form — the sheet-qualified
    `[Book]Sheet!col(n).type = ...` is silently ignored on Origin 2020."""
    execute_labtalk(f"win -a {book};")
    execute_labtalk(f'page.active$ = "{sheet}";')
    execute_labtalk(f"wks.col{col}.type = 3;")


def export_graph_to_file(graph_name: str, file_path: str, format: str = "png") -> str:
    """Export one graph directly to a file. Shared by export tools.

    Writes the file via Origin's expGraph X-Function with NO clipboard, so
    the user's clipboard contents are preserved. Exports at ~1200px wide with
    the aspect ratio kept. Returns the output path. Raises ValueError with a
    friendly message on failure.
    """
    return _export_via_expgraph(
        graph_name, file_path, width=1200, format=format
    )


@mcp.tool()
def create_graph(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    plot_type: str = "scatter",
    y_error_col: int = 0,
    z_col: int = 0,
    title: str = "",
    template: str = ""
) -> str:
    """Create a graph from worksheet data.

    Args:
        graph_name: Name for the graph window
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column number (1-based). Ignored for box/histogram.
        y_col: Y column number (1-based)
        plot_type: scatter, line, line+symbol, column, bar, area, pie,
                   box, histogram, contour, 3d_scatter
        y_error_col: Optional Y error column (1-based, 0=none). XY plots only.
        z_col: Z column (1-based). REQUIRED for contour and 3d_scatter.
        title: Optional graph title
        template: Optional path to a saved Origin graph template (.otpu/.otp,
                  e.g. from save_graph_template). When given, the new graph is
                  built FROM that template so it reproduces the saved styling
                  (frame, fonts, line/symbol style, colors) with this data.
                  2D XY plot types only (not contour/3d_scatter).

    Returns:
        JSON string: {"name": <actual graph name>, "requested_name": <name
        passed in>, "renamed": <bool, True if Origin uniquified the name>,
        "plot_type": <plot type used>}. Use "name" for subsequent calls —
        Origin renames on collision, so it may differ from requested_name.
    """
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_plot_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    ptype = PLOT_TYPES[safe_plot_type]
    require_worksheet(safe_book, safe_sheet)

    template_path = ""
    if template:
        if safe_plot_type in _XYZ_TYPES:
            msg = (
                f"create_graph(template=...) supports 2D XY plot types only, "
                f"not '{safe_plot_type}'."
            )
            raise ValueError(msg)
        template_path = windows_path(template, "template")
        if not os.path.isfile(template_path):
            msg = f"Template not found: {template_path}"
            raise ValueError(msg)

    def _set_title():
        if title:
            execute_labtalk(
                f"label -n title -s {labtalk_text(title, 'title')}; "
                "title.x = 50; title.y = 95;"
            )

    # --- XYZ plots: contour (2D) and 3d_scatter (OpenGL), via plotxyz ---
    if safe_plot_type in _XYZ_TYPES:
        if z_col < 1:
            msg = f"plot_type '{safe_plot_type}' requires z_col (1-based Z column)."
            raise ValueError(msg)
        safe_z_col = positive_column(z_col, "z_col")
        xyz = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_z_col})"
        if safe_plot_type in _OWN_GRAPH_TYPES:
            # Activate the source sheet (a non-graph window) so plotxyz
            # creates a fresh OpenGL 3D graph instead of merging into an
            # active 2D graph and collapsing to a flat projection.
            execute_labtalk(f"win -a {safe_book};")
            before = set(graph_names())
            if not execute_labtalk(f"plotxyz iz:={xyz} plot:={ptype};"):
                msg = (
                    f"Could not create a 3D plot from {xyz}. Check that the "
                    "columns contain numeric data."
                )
                raise ValueError(msg)
            new = set(graph_names()) - before
            name = new.pop() if new else o.LTStr("page.name$")
            if name != safe_graph_name and execute_labtalk(
                f"win -r {name} {safe_graph_name};"
            ):
                name = safe_graph_name
            settle_new_plots(name, expected_min_plots=1)
        else:
            name = o.CreatePage(3, safe_graph_name, "origin")
            if not execute_labtalk(
                f"plotxyz iz:={xyz} plot:={ptype} ogl:=[{name}]Layer1;"
            ):
                execute_labtalk(f"win -cd {name};")
                msg = (
                    f"Could not plot {xyz}. Check that columns "
                    f"{safe_x_col}, {safe_y_col}, {safe_z_col} contain data."
                )
                raise ValueError(msg)
            settle_new_plots(name, expected_min_plots=1)
        _set_title()
        return json.dumps({
            "name": name,
            "requested_name": safe_graph_name,
            "renamed": name != safe_graph_name,
            "plot_type": safe_plot_type,
        })

    # --- 2D plots (plotxy) ---
    if safe_plot_type in _Y_ONLY_TYPES:
        # Box/histogram need the source column designated as Y and the
        # sheet active, or Origin draws an empty layer.
        activate_window(safe_book, "data_book")
        execute_labtalk(f'page.active$ = "{safe_sheet}"; wks.col{safe_y_col}.type = 1;')
        data_ref = f"[{safe_book}]{safe_sheet}!col({safe_y_col})"
    else:
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
        if y_error_col > 0:
            safe_error_col = positive_column(y_error_col, "y_error_col")
            _designate_error_column(safe_book, safe_sheet, safe_error_col)
            data_ref = (
                f"[{safe_book}]{safe_sheet}!"
                f"({safe_x_col},{safe_y_col},{safe_error_col})"
            )

    if template_path:
        # Build the graph FROM the saved template so it reproduces the styling
        # (probe-verified route: plotxy ... ogl:=<new template:="<path>">).
        # plotm/plotxy into <new template:=...> spawns its own window, so grab
        # the new page name by diffing the graph list, then rename to requested.
        before = set(graph_names())
        if not execute_labtalk(
            f'plotxy iy:={data_ref} plot:={ptype} '
            f'ogl:=<new template:="{template_path}">;'
        ):
            msg = (
                f"Could not plot {data_ref} onto the template {template_path}. "
                f"Check the columns and that plot_type '{safe_plot_type}' suits "
                "this data."
            )
            raise ValueError(msg)
        new = set(graph_names()) - before
        name = new.pop() if new else o.LTStr("page.name$")
        if name != safe_graph_name and execute_labtalk(
            f"win -r {name} {safe_graph_name};"
        ):
            name = safe_graph_name
    else:
        name = o.CreatePage(3, safe_graph_name, "origin")
        if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{name}]Layer1;"):
            execute_labtalk(f"win -cd {name};")
            msg = (
                f"Could not plot {data_ref}. Check that columns exist and contain "
                f"data, and that plot_type '{safe_plot_type}' suits this data."
            )
            raise ValueError(msg)
    settle_new_plots(name, expected_min_plots=1)

    _set_title()
    return json.dumps({
        "name": name,
        "requested_name": safe_graph_name,
        "renamed": name != safe_graph_name,
        "plot_type": safe_plot_type,
        "template": template_path or None,
    })

@mcp.tool()
def add_plot_to_graph(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    plot_type: str = "scatter",
    y_error_col: int = 0
) -> str:
    """Add another data series to an existing graph.

    Args:
        graph_name: Existing graph name
        data_book: Source workbook name
        data_sheet: Source sheet name
        x_col: X column number (1-based)
        y_col: Y column number (1-based)
        plot_type: Plot type (scatter, line, line+symbol, etc.)
        y_error_col: Optional Y error column (1-based, 0=none)

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_x_col = positive_column(x_col, "x_col")
    safe_y_col = positive_column(y_col, "y_col")
    safe_plot_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    if safe_plot_type in _Y_ONLY_TYPES or safe_plot_type in _XYZ_TYPES:
        msg = (
            f"add_plot_to_graph supports only X,Y plot types; "
            f"'{safe_plot_type}' must be created with create_graph."
        )
        raise ValueError(msg)
    ptype = PLOT_TYPES[safe_plot_type]
    require_graph(safe_graph_name)
    require_worksheet(safe_book, safe_sheet)
    from .style_helpers import get_plot_info
    prior_count = len(get_plot_info(safe_graph_name))
    data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col})"
    if y_error_col > 0:
        safe_error_col = positive_column(y_error_col, "y_error_col")
        _designate_error_column(safe_book, safe_sheet, safe_error_col)
        data_ref = f"[{safe_book}]{safe_sheet}!({safe_x_col},{safe_y_col},{safe_error_col})"

    if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{safe_graph_name}]Layer1;"):
        msg = (
            f"Could not add plot {data_ref} to {safe_graph_name}. Check that "
            f"columns {safe_x_col} and {safe_y_col} exist and contain data."
        )
        raise ValueError(msg)
    settle_new_plots(safe_graph_name, expected_min_plots=prior_count + 1)
    return f"Added {safe_plot_type} plot to {safe_graph_name}"

@mcp.tool()
def delete_graph(graph_name: str) -> str:
    """Delete a graph window to keep the Origin project lightweight — use this
    to remove a badly-drawn or rejected figure.

    Args:
        graph_name: Graph to delete

    Returns:
        Confirmation message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    execute_labtalk(f"win -cd {safe_graph_name};")
    return f"Deleted graph '{safe_graph_name}'."


@mcp.tool()
def remove_plot(graph_name: str, plot_index: int = 1) -> str:
    """Remove a single data plot from a graph's first layer.

    Use this to delete a stray/duplicate/dead-guide curve without touching the
    others. Removes ONLY the plot at `plot_index` via the COM DataPlot's
    `Destroy()` — so when the SAME dataset is plotted more than once, only the
    indexed copy goes (the LabTalk `layer -e <dataset>` route removed EVERY plot
    of that dataset, since it addresses by dataset NAME, not by plot; verified).
    `Destroy()` is the only working per-index route on this build — the DataPlot
    COM object has no `Remove`/`Delete`, and `range r; delete r` reports success
    but leaves the plot on the layer.

    Args:
        graph_name: Graph name
        plot_index: 1-based index over the graph's DATA plots (same numbering
                    as set_plot_style; error-bar plots are not counted)

    Returns:
        Confirmation message naming the removed plot and the survivors.
    """
    from .style_helpers import acquire_graph_layer, get_plot_info

    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    infos = get_plot_info(safe_graph)
    # DataPlots.Item indexes over ALL plots (error bars included); map the
    # 1-based DATA-plot index to its position in the full collection.
    com_positions = [i for i, p in enumerate(infos) if not p["is_error"]]
    idx = plot_index - 1
    if idx < 0 or idx >= len(com_positions):
        n = len(com_positions)
        valid = f"1-{n}" if n else "(none)"
        data_names = [infos[i]["name"] for i in com_positions]
        msg = (
            f"Plot index {plot_index} not found on {safe_graph}. "
            f"Valid range: {valid}. Data plots: {data_names}"
        )
        raise ValueError(msg)
    com_index = com_positions[idx]
    pname = infos[com_index]["name"]
    # Fresh, activated layer handle so this works on graphs loaded from a .opju,
    # not only in-session ones. Destroy() removes just this one DataPlot.
    gl = acquire_graph_layer(safe_graph)
    try:
        gl.DataPlots.Item(com_index).Destroy()
    except Exception as exc:  # noqa: BLE001 - surface any COM failure truthfully
        msg = f"Origin could not remove plot {plot_index} ({pname}) from {safe_graph}: {exc}"
        raise ValueError(msg) from exc
    survivors = [p["name"] for p in get_plot_info(safe_graph)]
    return (
        f"Removed data plot {plot_index} ({pname}) from {safe_graph}. "
        f"Remaining plots: {survivors}"
    )


# `set <err> -o <y>` error-bar flag + the LabTalk Yerr/Xerr column designation
# to stamp on the error column afterwards (so it reads as an error plot, not a
# second data curve/legend entry).
_ERROR_BAR_FLAGS = {"y": "-o", "x": "-ox"}
_ERROR_BAR_DESIG = {"y": 3, "x": 7}  # wks.col.type: 3=Y Error, 7=X Error


@mcp.tool()
def set_error_bars(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    y_col: int,
    err_col: int,
    direction: str = "y",
) -> str:
    """Attach error bars to an EXISTING plot from an error column, in place.

    Uses Origin's `set <err> -o <y>` idiom: the error column is plotted into the
    layer, SETTLED (the missing settle after this plotxy is what silently broke
    the attach — the reassignment and the column-type designation raced the new
    plot and no-op'd), reassigned as error bars of the target Y plot, then
    designated as an error column and the legend rebuilt — so no stray curve or
    extra legend entry is left behind. The attach is VERIFIED by reading the
    error column's designation back (and, for Y, confirming no extra data curve
    survived); on failure the stray curve is removed and the tool raises instead
    of reporting a false success. The error and Y columns must live in the same
    sheet, and the target Y column must already be plotted on the graph.

    Args:
        graph_name: Graph that already shows the Y plot
        data_book: Workbook holding both the Y and error columns
        data_sheet: Sheet holding both columns
        y_col: 1-based column of the plotted Y data to decorate
        err_col: 1-based column holding the error values (SD/SE)
        direction: "y" for Y error bars (default) or "x" for X error bars

    Returns:
        Success message. Raises ValueError (after removing the stray curve) if
        the error column did not convert to error bars.
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    safe_y = positive_column(y_col, "y_col")
    safe_err = positive_column(err_col, "err_col")
    safe_dir = labtalk_choice(direction.lower(), _ERROR_BAR_FLAGS, "direction")
    if safe_y == safe_err:
        msg = "y_col and err_col must be different columns."
        raise ValueError(msg)
    require_graph(safe_graph)
    require_worksheet(safe_book, safe_sheet)
    flag = _ERROR_BAR_FLAGS[safe_dir]
    desig = _ERROR_BAR_DESIG[safe_dir]
    ref = f"[{safe_book}]{safe_sheet}"

    infos_before = get_plot_info(safe_graph)
    data_before = sum(1 for p in infos_before if not p["is_error"])

    # 1. Plot the error column into the layer (so it is a "plotted dataset").
    if not execute_labtalk(
        f"win -a {safe_graph}; "
        f"plotxy iy:={ref}!col({safe_err}) plot:=200 ogl:=[{safe_graph}]Layer1;"
    ):
        msg = (
            f"Origin could not plot error column {safe_err} of {ref} onto "
            f"{safe_graph}. Check that the column contains numeric data."
        )
        raise ValueError(msg)
    # 2. SETTLE — the fix. Without this the following designation/reassign race
    #    the just-added plot and silently no-op (col.type stays 0, no whiskers).
    settle_new_plots(safe_graph, expected_min_plots=len(infos_before) + 1)
    # 3. Designate the error column as Y/X Error, 4. reassign it as error bars of
    #    the Y plot, 5. rebuild the legend WITHOUT the error-column entry.
    execute_labtalk(
        f'win -a {safe_book}; page.active$ = "{safe_sheet}"; '
        f"wks.col{safe_err}.type = {desig};"
    )
    execute_labtalk(
        f"win -a {safe_graph}; "
        f"range __mcp_yr = {ref}!col({safe_y}); "
        f"range __mcp_er = {ref}!col({safe_err}); "
        f"set __mcp_er {flag} __mcp_yr;"
    )
    execute_labtalk(f"win -a {safe_graph}; legend -r;")

    # 6. VERIFY the attach took. The single reliable signal (live-probed) is the
    #    error column's designation sticking; for Y the error plot also drops out
    #    of the data-plot count (get_plot_info recognizes Y-error columns), so a
    #    surviving extra DATA curve means the reassignment failed.
    o = get_origin()
    o.Execute(
        f'win -a {safe_book}; page.active$ = "{safe_sheet}"; '
        f"__mcp_ct = wks.col{safe_err}.type;"
    )
    col_type = int(round(o.LTVar("__mcp_ct")))
    infos_after = get_plot_info(safe_graph)
    data_after = sum(1 for p in infos_after if not p["is_error"])
    attached = col_type == desig
    if safe_dir == "y":
        attached = (
            attached
            and data_after == data_before
            and any(p["is_error"] for p in infos_after)
        )
    if not attached:
        # Remove the stray error-column curve we added (the last plot), then
        # fail honestly instead of leaving damage behind a success message.
        try:
            gl = acquire_graph_layer(safe_graph)
            if infos_after:
                gl.DataPlots.Item(len(infos_after) - 1).Destroy()
        except Exception:  # noqa: BLE001 - best-effort cleanup; the raise below is the point
            pass
        msg = (
            f"Could not attach {safe_dir}-error bars from column {safe_err} to "
            f"the column {safe_y} plot on {safe_graph}: the error column did not "
            f"convert to error bars (read back col.type={col_type}, expected "
            f"{desig}). Removed the stray curve. Rebuild the graph with "
            f"y_error_col via create_graph / add_plot_to_graph instead."
        )
        raise ValueError(msg)
    return (
        f"Attached {safe_dir}-error bars (column {safe_err}) to the column "
        f"{safe_y} plot on {safe_graph}."
    )


@mcp.tool()
def set_layer_geometry(
    graph_name: str,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> str:
    """Set a graph layer's panel geometry (position and size).

    Values are in the layer's current units — Origin's default is percent of
    the page, so left=15, top=12, width=75, height=75 places a single panel
    with even margins. Use this to stop axis titles from being clipped outside
    the frame or to line up multi-panel figures. Only the provided fields are
    changed. Operates on the graph's first layer.

    Args:
        graph_name: Graph name
        left: Panel left edge (None = leave)
        top: Panel top edge (None = leave)
        width: Panel width (None = leave)
        height: Panel height (None = leave)

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    fields = {"left": left, "top": top, "width": width, "height": height}
    provided = {k: v for k, v in fields.items() if v is not None}
    if not provided:
        msg = "Provide at least one of left, top, width, or height."
        raise ValueError(msg)
    # Route each field through the fresh, activated layer handle
    # (graph_layer_execute) — like the axis-range sibling — so it also lands on
    # graphs loaded from a .opju, then READ IT BACK. Geometry rounds to page
    # units (~0.5%), so a loose tolerance still catches a total no-op (the
    # loaded-graph freeze) while tolerating that rounding.
    for k, v in provided.items():
        if not graph_layer_execute(safe_graph, f"layer.{k} = {float(v)};"):
            msg = f"Origin could not set layer {k} of {safe_graph}."
            raise ValueError(msg)
        got = read_layer_value(safe_graph, f"layer.{k}")
        if got is not None:
            tol = max(1.0, abs(float(v)) * 0.05)
            if abs(got - float(v)) > tol:
                msg = (
                    f"layer {k} did not take on {safe_graph} (set {v}, read back "
                    f"{got}). The layer ignored the change — the graph was most "
                    f"likely loaded from a .opju Origin freezes over COM; recreate "
                    f"it in-session with create_graph / plotxy."
                )
                raise ValueError(msg)
    return f"Set layer geometry of {safe_graph}: {', '.join(provided)}."


def _set_axis_labels_impl(
    graph_name: str,
    x_label: str = "",
    y_label: str = "",
    title: str = ""
) -> str:
    """Set axis labels and title for a graph.

    Args:
        graph_name: Graph name
        x_label: X axis label
        y_label: Y axis label
        title: Graph title

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    activate_window(safe_graph_name, "graph_name")
    if x_label:
        if not execute_labtalk(f"xb.text$ = {labtalk_text(x_label, 'x_label')};"):
            msg = f"Could not set the x-axis label on {safe_graph_name}."
            raise ValueError(msg)
    if y_label:
        if not execute_labtalk(f"yl.text$ = {labtalk_text(y_label, 'y_label')};"):
            msg = f"Could not set the y-axis label on {safe_graph_name}."
            raise ValueError(msg)
    if title:
        if not execute_labtalk(f"label -n title -s {labtalk_text(title, 'title')}; title.x = 50; title.y = 95;"):
            msg = f"Could not set the title on {safe_graph_name}."
            raise ValueError(msg)
    return f"Updated labels for {safe_graph_name}"


def _set_border_axis_title_impl(graph_name: str, side: str, label: str) -> str:
    """Set the RIGHT (dual-Y layer-2 ``yr``) or TOP (layer-1 ``xt``) axis
    title, then CONFIRM it landed by reading ``<obj>.text$`` back.

    Fail-honest: axis="right" requires a real second Y layer (from
    add_second_y_axis). Writing ``yr.text$`` on a single-layer graph
    "succeeds" at the LabTalk level (Execute returns true and the property
    reads back) but titles a right axis that isn't drawn — the false success
    the usability agent hit. So the second layer is required up front; the
    read-back then guards against a genuine no-op."""
    import time

    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    nlayers = int(get_origin().LTVar("page.nlayers"))
    if side == "right":
        if nlayers < 2:
            msg = (
                f"axis(op='labels', axis='right') needs a right-Y layer on "
                f"{safe_graph}, but it has {nlayers} layer(s). Add one with "
                "add_second_y_axis first, then set the right-axis title."
            )
            raise ValueError(msg)
        obj, layer = "yr", 2
    else:  # top
        obj, layer = "xt", 1
    safe = labtalk_text(label, "label")
    got = ""
    for attempt in range(3):
        execute_labtalk(f"win -a {safe_graph}; page.active={layer};")
        execute_labtalk(f"{obj}.text$ = {safe};")
        got = get_lt_str(f"{obj}.text$") or ""
        if label in got:
            return f"Set {side}-axis title on {safe_graph} (layer {layer})"
        if attempt < 2:
            time.sleep(0.3)
    msg = (
        f"Could not confirm the {side}-axis title on {safe_graph} "
        f"(read back: {got!r})."
    )
    raise ValueError(msg)

def _set_axis_range_impl(
    graph_name: str,
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None
) -> str:
    """Set axis range for a graph.

    Args:
        graph_name: Graph name
        x_min: X axis minimum (None=auto)
        x_max: X axis maximum (None=auto)
        y_min: Y axis minimum (None=auto)
        y_max: Y axis maximum (None=auto)

    Returns:
        Success message
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph_name)
    # Route each bound through the fresh, activated layer handle
    # (graph_layer_execute), then READ IT BACK: a graph loaded from a .opju used
    # to accept `layer.y.from = ...` and silently ignore it while returning
    # success. verify_layer_value turns that no-op into a loud, actionable error.
    for bound, value, label in (
        ("x.from", x_min, "x-axis minimum"),
        ("x.to", x_max, "x-axis maximum"),
        ("y.from", y_min, "y-axis minimum"),
        ("y.to", y_max, "y-axis maximum"),
    ):
        if value is None:
            continue
        prop = f"layer.{bound}"
        if not graph_layer_execute(safe_graph_name, f"{prop} = {value};"):
            msg = f"Could not set {label} on {safe_graph_name}."
            raise ValueError(msg)
        verify_layer_value(safe_graph_name, prop, float(value), label)
    return f"Set axis range for {safe_graph_name}"

def _export_graph_impl(
    graph_name: str,
    file_path: str,
    format: str = "png",
) -> str:
    """Export a graph at the default ~1200px-wide size (aspect ratio kept).

    Writes the file directly via Origin's expGraph X-Function (no clipboard),
    so the user's clipboard is preserved. This is the no-explicit-size path;
    export_graph routes an explicit width/height to the sized export instead.

    Args:
        graph_name: Graph name to export
        file_path: Output path (Windows or WSL style, e.g.
                   C:\\Users\\me\\fig1.png or /mnt/c/Users/me/fig1.png).
                   Missing directories are created.
        format: Output format. Raster: png, jpg, tif, bmp. Vector: pdf, eps,
                emf. Used as the file extension when file_path has none.

    Returns:
        Path to exported file
    """
    safe_format = labtalk_choice(format.lower(), EXPORT_FORMATS, "format")
    path = windows_path(file_path, "file_path")
    if not os.path.splitext(path)[1]:
        path = f"{path}.{safe_format}"

    out = export_graph_to_file(graph_name, path, format=safe_format)
    size = os.path.getsize(out)
    return f"Exported to: {out} ({size} bytes)"


# Origin axis scale type codes (layer.x/y.type).
_AXIS_SCALES = {"linear": 1, "log10": 2, "ln": 8, "log2": 9}
_LOG_SCALES = {"log10", "ln", "log2"}


def _rescale_axis_to_data(graph_name: str, axis: str, is_log: bool) -> bool:
    """Reset an axis's from/to to the plotted data extent (ACTUAL values, not
    exponents). For a log axis, non-positive points are dropped so the range
    can't collapse to a garbage decade like 1E-9. Best-effort: returns False
    (leaving Origin's auto range) when the data can't be read."""
    from .style_helpers import _collect_xy

    try:
        xs, ys = _collect_xy(graph_name)
    except Exception:
        return False
    vals = [v for v in (xs if axis == "x" else ys) if v is not None]
    if is_log:
        vals = [v for v in vals if v > 0]
    if len(vals) < 2:
        return False
    lo, hi = min(vals), max(vals)
    if lo == hi:
        return False
    return bool(graph_layer_execute(
        graph_name, f"layer.{axis}.from = {lo}; layer.{axis}.to = {hi};"
    ))


def _set_axis_scale_impl(
    graph_name: str, axis: str = "y", scale: str = "log10", rescale: bool = True
) -> str:
    """Set an axis to linear or a logarithmic scale.

    Args:
        graph_name: Graph name
        axis: "x" or "y"
        scale: linear, log10, ln, or log2
        rescale: When True (default), reset the axis range to the data extent
                 after the scale change so a log switch doesn't leave garbage
                 ticks (e.g. 1E-9 … 100). Range bounds are ACTUAL data values,
                 not exponents.

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_axis = labtalk_choice(axis.lower(), {"x", "y"}, "axis")
    safe_scale = labtalk_choice(scale.lower(), _AXIS_SCALES, "scale")
    require_graph(safe_graph)
    # Fresh, activated layer handle so the scale change also lands on graphs
    # loaded from a .opju (where a global `layer.*` after `win -a` froze).
    if not graph_layer_execute(safe_graph, f"layer.{safe_axis}.type = {_AXIS_SCALES[safe_scale]};"):
        msg = f"Could not set {safe_axis} axis of {safe_graph} to {safe_scale}."
        raise ValueError(msg)
    rescaled = ""
    if rescale and _rescale_axis_to_data(safe_graph, safe_axis, safe_scale in _LOG_SCALES):
        rescaled = " (auto-rescaled to data)"
    return f"Set {safe_axis} axis of {safe_graph} to {safe_scale} scale{rescaled}"


_FRAME_MODES = {"closed", "open"}


def _set_axis_frame_impl(
    graph_name: str, frame: str = "closed", frame_width: float | None = None
) -> str:
    """Close or open a graph's frame — the top (opposite X) and right
    (opposite Y) border axes — and/or set the frame LINE width.

    Args:
        graph_name: Graph name
        frame: "closed" (draw top+right border axes) or "open" (hide them)
        frame_width: Frame line width in points, applied to all 4 sides
                     (x/y/x2/y2.thickness — probe-confirmed on Origin 2020:
                     0.5 is hairline, 2 is a typical publication frame, 8 is
                     very heavy). None = leave the current width unchanged.

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_frame = labtalk_choice(frame.lower(), _FRAME_MODES, "frame")
    require_graph(safe_graph)
    opposite = 1 if safe_frame == "closed" else 0
    # Fresh, activated layer handle so the frame toggle also lands on loaded graphs.
    if not graph_layer_execute(safe_graph, f"layer.x.opposite = {opposite}; layer.y.opposite = {opposite};"):
        msg = f"Could not set the frame of {safe_graph} to {safe_frame}."
        raise ValueError(msg)
    width_note = ""
    if frame_width is not None:
        w = float(frame_width)
        for axis_prop in ("layer.x.thickness", "layer.y.thickness",
                           "layer.x2.thickness", "layer.y2.thickness"):
            if not graph_layer_execute(safe_graph, f"{axis_prop} = {w};"):
                msg = f"Could not set {axis_prop} on {safe_graph}."
                raise ValueError(msg)
        verify_layer_value(safe_graph, "layer.x.thickness", w, "frame_width")
        width_note = f", frame_width {w}pt"
    return f"Set frame of {safe_graph} to {safe_frame}{width_note}"

@mcp.tool()
def add_second_y_axis(
    graph_name: str,
    data_book: str,
    data_sheet: str,
    x_col: int,
    y_col: int,
    plot_type: str = "line+symbol",
    color: str = "",
) -> str:
    """Add a right-side Y axis layer and plot a second dataset on it.

    Args:
        graph_name: Existing graph
        data_book, data_sheet: Source data
        x_col, y_col: Columns for the right-axis series (1-based)
        plot_type: scatter, line, line+symbol, column, bar, area
        color: Optional "r,g,b" (each 0-255) — always this comma-separated
            numeric form, unlike `set_plot_style`'s `color` (named, e.g.
            "blue") vs. its separate `rgb` param. When given, the RIGHT axis
            LINE (with its ticks) and the RIGHT tick-number LABELS are
            colored to match — the dual-Y "right axis colored like its data"
            convention. Leave empty to keep them black.

    Returns:
        Success message

    The visible right axis is layer 2's ``y2`` (positional right) — NOT its
    ``y``, which is a hidden LEFT axis and a silent-no-op trap. `color` above
    sets the axis line + tick labels only; the right-axis TITLE is a separate
    ``YR`` text object this tool does not touch. Full recipe (title text/
    color/size, why ``layer.y.*``/``yl`` are traps here):
    `get_skill('labtalk-gotchas')`.
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_book = labtalk_name(data_book, "data_book")
    safe_sheet = labtalk_name(data_sheet, "data_sheet")
    sx = positive_column(x_col, "x_col")
    sy = positive_column(y_col, "y_col")
    safe_type = labtalk_choice(plot_type, PLOT_TYPES, "plot_type")
    if safe_type in _Y_ONLY_TYPES or safe_type in _XYZ_TYPES:
        msg = f"add_second_y_axis supports only X,Y plot types, not '{safe_type}'."
        raise ValueError(msg)
    ptype = PLOT_TYPES[safe_type]
    require_graph(safe_graph)
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_graph, "graph_name")
    if not execute_labtalk("layadd type:=righty;"):
        msg = f"Could not add a right-Y layer to {safe_graph}."
        raise ValueError(msg)
    layer = int(get_origin().LTVar("page.nlayers"))
    data_ref = f"[{safe_book}]{safe_sheet}!({sx},{sy})"
    if not execute_labtalk(f"plotxy iy:={data_ref} plot:={ptype} ogl:=[{safe_graph}]{layer}!;"):
        msg = f"Added the right-Y layer but could not plot {data_ref} on it."
        raise ValueError(msg)
    msg = f"Added right-Y axis (layer {layer}) to {safe_graph} with {safe_type} plot"
    if color:
        from .style import _parse_rgb
        r, g, b = _parse_rgb(color)
        col = f"color({r},{g},{b})"
        # The right axis is layer 2's y2 (positional right). Select that layer,
        # then color the LINE (y2.color) and the tick-NUMBER labels
        # (y2.label.color). Issued as separate statements: Origin can fail a
        # 3+ layer.* batch as a whole and apply nothing.
        execute_labtalk(f"win -a {safe_graph};")
        execute_labtalk(f"page.active={layer};")
        ok_line = execute_labtalk(f"layer.y2.color={col};")
        ok_label = execute_labtalk(f"layer.y2.label.color={col};")
        execute_labtalk("doc -uw;")
        if ok_line and ok_label:
            msg += f"; colored right axis line + tick labels {col}"
        else:
            msg += f"; WARNING right-axis color {col} may not have applied"
    # Adding the second-layer plot leaves the legend with only layer 1's
    # entries, a default border, and often overlapping data (the wfB_dualy
    # regression, item 30). Rebuild it to cover BOTH layers, drop the border,
    # and re-place it away from the data via the shared placer.
    import time

    from .style_helpers import place_legend_avoiding_data

    execute_labtalk(f"win -a {safe_graph}; legend -r;")
    execute_labtalk("legend.background = 0;")
    time.sleep(0.3)
    placement = place_legend_avoiding_data(safe_graph)
    msg += f"; legend rebuilt borderless, placed {placement}"
    return msg


@mcp.tool()
def add_layer(graph_name: str, layer_type: str = "right-y") -> str:
    """Add a new layer (panel/axis) to a graph.

    Args:
        graph_name: Graph name
        layer_type: right-y, top-x, inset, or independent

    Returns:
        Success message naming the new layer index
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    type_map = {
        "right-y": "righty",
        "top-x": "topx",
        "inset": "inset",
        "independent": "independent",
    }
    safe_layer_type = labtalk_choice(layer_type.lower(), type_map, "layer_type")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    if not execute_labtalk(f"layadd type:={type_map[safe_layer_type]};"):
        msg = f"Could not add a {safe_layer_type} layer to {safe_graph}."
        raise ValueError(msg)
    layer = int(get_origin().LTVar("page.nlayers"))
    return f"Added {safe_layer_type} layer (layer {layer}) to {safe_graph}"


def _add_reference_line_impl(
    graph_name: str,
    orientation: str,
    value: float
) -> str:
    """Draw a horizontal or vertical reference line at a data value.

    Args:
        graph_name: Graph name
        orientation: "horizontal" (constant Y) or "vertical" (constant X)
        value: Axis value where the line is drawn

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_orientation = labtalk_choice(
        orientation.lower(), {"horizontal", "vertical"}, "orientation"
    )
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    flag = "-h" if safe_orientation == "horizontal" else "-v"
    if not execute_labtalk(f"draw -l {flag} {float(value)};"):
        msg = f"Could not draw a {safe_orientation} reference line on {safe_graph}."
        raise ValueError(msg)
    return f"Drew {safe_orientation} reference line at {value} on {safe_graph}"


def _add_text_annotation_impl(
    graph_name: str,
    text: str,
    x: float,
    y: float,
    name: str = "anno"
) -> str:
    """Add a text label to a graph at data coordinates.

    Args:
        graph_name: Graph name
        text: Annotation text (no quotes, line breaks, or ';')
        x: X position in data coordinates
        y: Y position in data coordinates
        name: Internal object name (letters/numbers/underscore)

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_name = labtalk_name(name, "name")
    if any(ch in text for ch in ('"', "\r", "\n", ";")) or not text.strip():
        msg = "text cannot be empty or contain quotes, line breaks, or ';'."
        raise ValueError(msg)
    validate_text_escapes(text, "text")  # reject \q() etc. (LaTeX modal wedge)
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    if not execute_labtalk(f"label -p {float(x)} {float(y)} -n {safe_name} {text};"):
        msg = f"Could not add the annotation to {safe_graph}."
        raise ValueError(msg)
    return f"Added annotation '{text}' at ({x}, {y}) on {safe_graph}"


def _export_via_expgraph(
    graph_name: str,
    file_path: str,
    *,
    width: int,
    format: str = "png",
) -> str:
    """Export one graph directly to an image file via Origin's expGraph.

    Writes the file with NO clipboard (the user's clipboard is preserved).
    Verified on Origin 2020: `tr1.unit:=2` selects pixels and `tr1.width` sets
    the pixel width; the HEIGHT is always derived from the graph's aspect ratio
    (probe-verified: `tr1.height` is silently ignored — 1600x1000 requested
    yields 1600x1224), so only width is controllable. Returns the output path.
    Raises ValueError with a friendly message on failure.
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_format = labtalk_choice(format.lower(), EXPORT_FORMATS, "format")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    path = windows_path(file_path, "file_path")
    if not os.path.splitext(path)[1]:
        path = f"{path}.{safe_format}"
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    fname = os.path.splitext(os.path.basename(path))[0]
    # Vector formats (pdf/eps/emf) are resolution-independent — passing a
    # pixel size option makes expGraph fail on Origin 2020 (probe-verified),
    # so size options are emitted only for raster formats.
    size_opts = ""
    if safe_format in EXPORT_RASTER_FORMATS:
        safe_width = positive_int(width, "width")
        size_opts = f" tr1.unit:=2 tr1.width:={safe_width}"
    cmd = (
        f'expGraph type:={safe_format} path:="{out_dir}" filename:="{fname}" '
        f"overwrite:=replace{size_opts};"
    )
    if not execute_labtalk(cmd):
        msg = f"Origin could not export {safe_graph} to {path}."
        raise ValueError(msg)
    if not os.path.exists(path):
        hint = ""
        if path.startswith("/"):
            hint = (
                " This looks like a WSL/Linux path Origin (Windows) cannot write "
                "to — use a Windows path (C:\\...) or /mnt/<drive>/... instead."
            )
        msg = f"Export failed: {path} was not created.{hint}"
        raise ValueError(msg)
    return path


def _export_graph_sized_impl(
    graph_name: str,
    file_path: str,
    width: int = 1200,
    format: str = "png"
) -> str:
    """Export a graph to an image at a chosen pixel WIDTH (expGraph).

    Like export_graph (also clipboard-free via expGraph), but controls the
    output pixel width directly. The height is always derived from the graph's
    aspect ratio (Origin's expGraph ignores an explicit height on this build).

    Args:
        graph_name: Graph to export
        file_path: Output path (Windows or WSL style)
        width: Image width in pixels (default 1200; raster only)
        format: Raster (png, jpg, tif, bmp) or vector (pdf, eps, emf). Vector
                formats are resolution-independent and ignore width.

    Returns:
        Path and pixel/byte size of the exported file
    """
    safe_width = positive_int(width, "width")
    path = _export_via_expgraph(
        graph_name, file_path, width=width, format=format
    )
    return (
        f"Exported {labtalk_name(graph_name, 'graph_name')} to {path} "
        f"({safe_width}px wide, {os.path.getsize(path)} bytes)"
    )


# Perceptually-uniform, colorblind-safe colormaps that Origin 2020 does NOT
# ship (viridis/cividis/etc. were added to Origin only in later versions). We
# bundle them as RIFF .pal files and load them by full path. Names are matched
# case-insensitively against this directory.
_BUNDLED_PAL_DIR = os.path.normpath(
    os.path.join(os.path.dirname(__file__), os.pardir, "palettes")
)


def _bundled_palette_path(name: str) -> str:
    """Full path to a bundled .pal whose stem matches `name` (case-insensitive),
    or "" if none. Bundled maps win over same-named built-ins."""
    if not name or not os.path.isdir(_BUNDLED_PAL_DIR):
        return ""
    want = f"{name.strip().lower()}.pal"
    for fn in os.listdir(_BUNDLED_PAL_DIR):
        if fn.lower() == want:
            return os.path.join(_BUNDLED_PAL_DIR, fn)
    return ""


def _origin_palette_dir() -> str:
    """Origin's built-in ``.pal`` folder (``<install>\\Palettes``), or "" when
    it can't be resolved (fakes / CI / odd install). Read from
    ``system.path.program$`` at runtime so it isn't hard-coded to one install."""
    try:
        prog = get_lt_str("system.path.program$")
    except Exception:
        return ""
    if not prog:
        return ""
    pal_dir = os.path.join(prog, "Palettes")
    return pal_dir if os.path.isdir(pal_dir) else ""


def _palette_is_valid(name: str):
    """True if ``name`` is a bundled or Origin built-in ``.pal``; False if we can
    prove it is neither; None if Origin's palette folder can't be read (so the
    caller must NOT reject upfront and should fall back to ``load()``).

    Needed because ``layer.cmap.load()`` returns True even for a NONEXISTENT
    palette (live-confirmed) and no readable cmap state changes between
    palettes, so a mistyped name would otherwise silently do nothing behind an
    "Applied palette" success message."""
    if _bundled_palette_path(name):
        return True
    pal_dir = _origin_palette_dir()
    if not pal_dir:
        return None
    want = f"{str(name).strip().lower()}.pal"
    try:
        return want in {fn.lower() for fn in os.listdir(pal_dir)}
    except OSError:
        return None


def _apply_color_map_impl(graph_name: str, palette: str = "Viridis") -> str:
    """Apply a color palette to a contour/heatmap/surface graph.

    Args:
        graph_name: Graph name (must hold a colormapped plot)
        palette: Palette name. Bundled perceptually-uniform, colorblind-safe
            maps (recommended for quantitative data): Viridis, Cividis, Plasma,
            Inferno, Magma; muted/pastel variants matching a soft figure
            aesthetic: PastelViridis, PastelCividis. Also accepts built-in
            Origin .pal names, e.g. Heatmap4ColorBlind, GrayScale, RedWhiteBlue,
            Fire, Temperature.

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    # Reject a palette name Origin doesn't ship BEFORE touching the graph:
    # layer.cmap.load() returns True even for a nonexistent palette and no
    # readable cmap state distinguishes a real recolor from a no-op, so an
    # unknown name would otherwise report success while doing nothing.
    if _palette_is_valid(palette) is False:
        msg = (
            f"Unknown palette '{palette}'. Origin's colormap load() reports "
            f"success even for a nonexistent palette, so applying this would "
            f"silently do nothing. Use a bundled map (Viridis, Cividis, Plasma, "
            f"Inferno, Magma, PastelViridis, PastelCividis) or an Origin "
            f"built-in .pal name (e.g. Fire, GrayScale, RedWhiteBlue, "
            f"Temperature, Heatmap4ColorBlind)."
        )
        raise ValueError(msg)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")

    bundled = _bundled_palette_path(palette)
    if bundled:
        # Copy to a clean, space-free temp path so LabTalk's load never trips on
        # spaces in the install path (e.g. "My Drive"). The file is guaranteed
        # to exist, so load cannot raise Origin's modal "file not found" dialog.
        disp = os.path.splitext(os.path.basename(bundled))[0]
        tmp = os.path.join(tempfile.gettempdir(), f"opm_{disp}.pal")
        try:
            shutil.copyfile(bundled, tmp)
        except OSError as exc:
            msg = f"Could not stage bundled palette '{disp}': {exc}"
            raise ValueError(msg) from exc
        load_arg = f'"{tmp}"'
    else:
        disp = labtalk_name(palette, "palette")
        load_arg = f"{disp}.pal"

    if not execute_labtalk(
        f"layer.cmap.load({load_arg}); layer.cmap.updateScale();"
    ):
        msg = f"Could not apply palette '{disp}' to {safe_graph}."
        raise ValueError(msg)
    return f"Applied '{disp}' palette to {safe_graph}"


def _set_colormap_levels_impl(graph_name: str, z_min: float, z_max: float) -> str:
    """Set the Z range (color scale levels) of a colormapped graph.

    Args:
        graph_name: Graph name (contour/heatmap/surface)
        z_min: Minimum Z for the color scale
        z_max: Maximum Z for the color scale

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    if z_max <= z_min:
        msg = "z_max must be greater than z_min."
        raise ValueError(msg)
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    if not execute_labtalk(
        f"layer.cmap.zmin = {float(z_min)}; layer.cmap.zmax = {float(z_max)}; "
        "layer.cmap.SetLevels(); layer.cmap.updateScale();"
    ):
        msg = f"Could not set colormap levels on {safe_graph}."
        raise ValueError(msg)
    # Read zmin/zmax back and confirm they took (mirrors the levels branch's
    # numColors read-back). On a non-colormap graph — or a build where the set
    # no-ops — the read-back won't match and this raises instead of reporting a
    # false success (live-confirmed readable on a real colormap plot).
    o = get_origin()
    got_min = o.LTVar("layer.cmap.zmin")
    got_max = o.LTVar("layer.cmap.zmax")
    tol_min = max(1e-6, abs(float(z_min)) * 1e-6)
    tol_max = max(1e-6, abs(float(z_max)) * 1e-6)
    if abs(got_min - float(z_min)) > tol_min or abs(got_max - float(z_max)) > tol_max:
        msg = (
            f"Colormap Z range did not take on {safe_graph}: requested "
            f"[{z_min}, {z_max}], read back [{got_min}, {got_max}]. "
            f"(Is this graph a colormapped plot?)"
        )
        raise ValueError(msg)
    return f"Set colormap Z range to [{z_min}, {z_max}] on {safe_graph}"


def _set_colormap_level_count_impl(graph_name: str, levels: int) -> str:
    """Raise the colormap's level count so a banded map reads as continuous.

    Origin's colormap is inherently banded; a Viridis-looking "continuous" heat
    map is just many levels. The verified recipe (Origin 2020) is
    ``numMajorLevels`` + ``setLevels(1)`` + ``updateScale()``. NOTE:
    ``numColors=<n>; setLevels();`` (no-arg) does NOT rebuild the map — it is a
    silent no-op — which is why raising the level count looked impossible.
    Verified by reading ``layer.cmap.numColors`` back and comparing.

    Args:
        graph_name: Graph name (contour/heatmap/surface)
        levels: Number of color levels (>= 2). ~32-64 reads as continuous.

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    if int(levels) < 2:
        msg = "levels must be >= 2."
        raise ValueError(msg)
    n = int(levels)
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    # Separate statements: Origin can fail a multi-statement layer.cmap batch as
    # a whole and apply nothing. numMinorLevels=0 keeps numColors == numMajorLevels
    # so the read-back check is exact.
    execute_labtalk("layer.cmap.numMinorLevels = 0;")
    execute_labtalk(f"layer.cmap.numMajorLevels = {n};")
    execute_labtalk("layer.cmap.setLevels(1);")
    if not execute_labtalk("layer.cmap.updateScale();"):
        msg = f"Could not update the colormap scale on {safe_graph}."
        raise ValueError(msg)
    got = get_origin().LTVar("layer.cmap.numColors")
    if int(round(got)) != n:
        msg = (
            f"Colormap level count did not take on {safe_graph}: requested {n}, "
            f"read back {got}. (Is this graph a colormapped plot?)"
        )
        raise ValueError(msg)
    return f"Set colormap to {n} continuous levels on {safe_graph}"


def _add_line_impl(
    graph_name: str,
    x1: float,
    y1: float,
    x2: float,
    y2: float
) -> str:
    """Draw a straight line between two data points on a graph.

    Useful for guides, connectors, and trend indicators. For a line with
    an arrowhead, use add_arrow instead.

    Args:
        graph_name: Graph name
        x1, y1: Start point in data coordinates
        x2, y2: End point in data coordinates

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    coords = f"{{{float(x1)},{float(y1)},{float(x2)},{float(y2)}}}"
    if not execute_labtalk(f"draw -l {coords};"):
        msg = f"Could not draw a line on {safe_graph}."
        raise ValueError(msg)
    return f"Drew line ({x1},{y1})->({x2},{y2}) on {safe_graph}"


def _add_arrow_impl(
    graph_name: str,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    double_headed: bool = False,
    head_size: int = 10
) -> str:
    """Draw an arrow from (x1,y1) to (x2,y2) at data coordinates.

    The arrowhead sits at the (x2,y2) end (and at the start too when
    double_headed). Set double_headed=False for a single-ended arrow.

    Args:
        graph_name: Graph name
        x1, y1: Tail (start) point in data coordinates
        x2, y2: Head (end) point in data coordinates
        double_headed: Put an arrowhead on both ends
        head_size: Arrowhead size in points (default 10)

    Returns:
        Success message
    """
    safe_graph = labtalk_name(graph_name, "graph_name")
    safe_size = positive_int(head_size, "head_size")
    require_graph(safe_graph)
    activate_window(safe_graph, "graph_name")
    execute_labtalk("layer1;")
    name = "arr" + uuid.uuid4().hex[:8]
    coords = f"{{{float(x1)},{float(y1)},{float(x2)},{float(y2)}}}"
    if not execute_labtalk(f"draw -n {name} -l {coords};"):
        msg = f"Could not draw an arrow on {safe_graph}."
        raise ValueError(msg)
    width = max(safe_size * 2 // 3, 3)
    script = (
        f"{name}.arrowEndShape = 1; "
        f"{name}.arrowEndLength = {safe_size}; {name}.arrowEndWidth = {width};"
    )
    if double_headed:
        script += (
            f" {name}.arrowBeginShape = 1; "
            f"{name}.arrowBeginLength = {safe_size}; {name}.arrowBeginWidth = {width};"
        )
    execute_labtalk(script)
    ends = "double-headed" if double_headed else "single-headed"
    return f"Drew {ends} arrow ({x1},{y1})->({x2},{y2}) on {safe_graph}"


# --- Consolidated dispatchers (Phase 2) ---------------------------------------

_AXIS_OPS = {"labels", "range", "scale", "tick", "frame"}


@mcp.tool()
def axis(
    graph_name: str,
    op: str,
    axis: str = "both",
    label: str | None = None,
    range_min: float | None = None,
    range_max: float | None = None,
    scale: str | None = None,
    tick_direction: str | None = None,
    major_length: int | None = None,
    minor_count: int | None = None,
    show_minor: bool | None = None,
    rescale: bool = True,
    frame: str | None = None,
    frame_width: float | None = None,
) -> str:
    """Configure a graph's axes.

    Args:
        graph_name: Graph name.
        op: Which axis aspect to set:
            - "labels": set an axis label. Uses `axis` ("x", "y", "both",
              "right", or "top") and `label` (the text). "right" sets the
              dual-Y right-axis title (requires a second Y layer from
              add_second_y_axis — raises otherwise, never a false success);
              "top" sets the top-axis title. Note: does not set the graph
              title.
            - "range": set an axis range. Uses `axis` plus `range_min`/
              `range_max` (None = auto). With axis="both" the same range is
              applied to X and Y.
            - "scale": set an axis scale. Uses `axis` ("x" or "y") and `scale`
              (linear, log10, ln, log2). By default the axis range is reset to
              the data extent after the change (pass rescale=False to keep the
              current range); range bounds are ACTUAL values, not exponents.
            - "tick": set tick-mark style. Uses `axis` ("x", "y", "both",
              "top", or "right" — "top"/"right" target ONLY the opposite-side
              border axis (x2/y2), e.g. to strip its tick marks without
              touching the bottom/left axis), `tick_direction` (in/out/both/
              none, default in — "none" removes that side's tick MARKS when
              present while always leaving its number labels alone (a default
              closed frame may show few/no opposite-side marks, so the visible
              change can be subtle); NEVER emits `majorTicks`, which
              probe-confirmed wipes ALL axes' number labels), plus
              `major_length` (default 8), `minor_count` (default 4),
              `show_minor` (default True).
            - "frame": close/open the frame (top+right border axes) and/or
              set the frame LINE width. Uses `frame` ("closed" or "open",
              default "closed") and `frame_width` (points, applied to all 4
              sides; None = leave unchanged).
        axis: Target axis ("x", "y", "both", "top", or "right"; default "both").
            For op="labels", "right"/"top" set the right/top-axis TITLE.
        label: Axis label text (op="labels").
        range_min, range_max: Axis range bounds (op="range").
        scale: Axis scale type (op="scale").
        tick_direction, major_length, minor_count, show_minor: Tick style
            (op="tick").
        rescale: For op="scale", auto-rescale to the data extent (default True).
        frame: For op="frame", "closed" or "open" (default "closed").
        frame_width: For op="frame", frame line width in points (None = leave
            unchanged).

    Returns:
        Success message for the selected operation.
    """
    safe_op = labtalk_choice(op.lower(), _AXIS_OPS, "op")
    if safe_op == "labels":
        if label is None:
            msg = "axis op 'labels' requires label."
            raise ValueError(msg)
        safe_axis = labtalk_choice(
            axis.lower(), {"x", "y", "both", "right", "top"}, "axis"
        )
        if safe_axis in ("right", "top"):
            return _set_border_axis_title_impl(graph_name, safe_axis, label)
        x_label = label if safe_axis in ("x", "both") else ""
        y_label = label if safe_axis in ("y", "both") else ""
        return _set_axis_labels_impl(graph_name, x_label=x_label, y_label=y_label)
    if safe_op == "range":
        safe_axis = labtalk_choice(axis.lower(), {"x", "y", "both"}, "axis")
        kwargs: dict = {}
        if safe_axis in ("x", "both"):
            kwargs["x_min"] = range_min
            kwargs["x_max"] = range_max
        if safe_axis in ("y", "both"):
            kwargs["y_min"] = range_min
            kwargs["y_max"] = range_max
        return _set_axis_range_impl(graph_name, **kwargs)
    if safe_op == "scale":
        if scale is None:
            msg = "axis op 'scale' requires scale."
            raise ValueError(msg)
        # _set_axis_scale_impl only takes a single axis ("x" or "y"); apply
        # to both when the caller asked for "both" (or left it unset) instead
        # of silently narrowing to Y and leaving X untouched.
        safe_scale_axis = labtalk_choice(
            (axis or "both").lower(), {"x", "y", "both"}, "axis"
        )
        if safe_scale_axis == "both":
            x_msg = _set_axis_scale_impl(
                graph_name, axis="x", scale=scale, rescale=rescale
            )
            y_msg = _set_axis_scale_impl(
                graph_name, axis="y", scale=scale, rescale=rescale
            )
            return f"{x_msg}. {y_msg}"
        return _set_axis_scale_impl(
            graph_name, axis=safe_scale_axis, scale=scale, rescale=rescale
        )
    if safe_op == "frame":
        return _set_axis_frame_impl(
            graph_name,
            frame=frame if frame is not None else "closed",
            frame_width=frame_width,
        )
    # tick
    from .style import _set_tick_style_impl
    safe_tick_axis = labtalk_choice(
        axis.lower(), {"x", "y", "both", "top", "right"}, "axis"
    )
    return _set_tick_style_impl(
        graph_name,
        axis=safe_tick_axis,
        tick_direction=tick_direction if tick_direction is not None else "in",
        major_length=major_length if major_length is not None else 8,
        minor_count=minor_count if minor_count is not None else 4,
        show_minor=show_minor if show_minor is not None else True,
    )


_ANNOTATE_KINDS = {"reference_line", "text", "line", "arrow"}


@mcp.tool()
def annotate(
    graph_name: str,
    kind: str,
    x1: float | None = None,
    y1: float | None = None,
    x2: float | None = None,
    y2: float | None = None,
    value: float | None = None,
    text: str | None = None,
    orientation: str | None = None,
    double_headed: bool = False,
    head_size: int = 10,
    name: str = "anno",
) -> str:
    """Add an annotation object to a graph.

    Args:
        graph_name: Graph name.
        kind: Which annotation to add:
            - "reference_line": horizontal/vertical line. Uses `orientation`
              ("horizontal"/"vertical") and `value` (axis value).
            - "text": text label. Uses `text` and position `x1`/`y1` (data
              coordinates) and `name` (object name).
            - "line": straight line from (x1,y1) to (x2,y2).
            - "arrow": arrow from (x1,y1) to (x2,y2). Uses `double_headed`
              and `head_size`.
        x1, y1: Start/position in data coordinates.
        x2, y2: End point in data coordinates (line/arrow).
        value: Axis value (kind="reference_line").
        text: Annotation text (kind="text").
        orientation: "horizontal" or "vertical" (kind="reference_line").
        double_headed: Arrowhead on both ends (kind="arrow").
        head_size: Arrowhead size in points (kind="arrow").
        name: Object name (kind="text").

    Returns:
        Success message for the selected annotation.
    """
    safe_kind = labtalk_choice(kind.lower(), _ANNOTATE_KINDS, "kind")
    if safe_kind == "reference_line":
        if orientation is None or value is None:
            msg = "annotate kind 'reference_line' requires orientation and value."
            raise ValueError(msg)
        return _add_reference_line_impl(graph_name, orientation, value)
    if safe_kind == "text":
        if text is None or x1 is None or y1 is None:
            msg = "annotate kind 'text' requires text, x1, and y1."
            raise ValueError(msg)
        return _add_text_annotation_impl(graph_name, text, x1, y1, name=name)
    if safe_kind == "line":
        if x1 is None or y1 is None or x2 is None or y2 is None:
            msg = "annotate kind 'line' requires x1, y1, x2, and y2."
            raise ValueError(msg)
        return _add_line_impl(graph_name, x1, y1, x2, y2)
    # arrow
    if x1 is None or y1 is None or x2 is None or y2 is None:
        msg = "annotate kind 'arrow' requires x1, y1, x2, and y2."
        raise ValueError(msg)
    return _add_arrow_impl(
        graph_name, x1, y1, x2, y2,
        double_headed=double_headed, head_size=head_size,
    )


@mcp.tool()
def colormap(
    graph_name: str,
    palette: str | None = None,
    z_min: float | None = None,
    z_max: float | None = None,
    levels: int | None = None,
) -> str:
    """Configure the color map of a contour/heatmap/surface graph.

    Applies a palette when `palette` is given, sets the Z color-scale range when
    both `z_min` and `z_max` are given, and/or raises the number of color levels
    when `levels` is given. At least one of those must be supplied.

    Args:
        graph_name: Graph name (must hold a colormapped plot).
        palette: Palette name (e.g. Viridis, Cividis, Plasma, Fire). Bundled
            perceptually-uniform, colorblind-safe maps are preferred for
            quantitative data; Origin built-in .pal names are also accepted.
            (This DOES recolor the map — do not be misled if an exported PNG's
            byte size is unchanged; the pixels change.) An unknown palette name
            is REJECTED up front: Origin's load() reports success for a
            nonexistent palette, so a typo would otherwise silently do nothing.
        z_min, z_max: Z range for the color scale (both required together). The
            applied range is read back (layer.cmap.zmin/zmax) and raises if it
            did not take (e.g. the graph is not a colormapped plot).
        levels: Number of color levels (>= 2). Origin's colormap is inherently
            banded; a default heatmap shows ~8 bands. Set ~32-64 to make it read
            as a continuous (smooth) map. Verified read-back
            (`layer.cmap.numColors`); raises if it did not take. Implemented via
            `numMajorLevels` + `setLevels(1)` + `updateScale()` — the no-arg
            `setLevels()` after `numColors=` is a silent no-op on this build.

    Returns:
        Success message describing what was changed.
    """
    if palette is None and z_min is None and z_max is None and levels is None:
        msg = "colormap requires palette, z_min+z_max, and/or levels."
        raise ValueError(msg)
    if (z_min is None) != (z_max is None):
        msg = "colormap requires both z_min and z_max together."
        raise ValueError(msg)
    messages = []
    if palette is not None:
        messages.append(_apply_color_map_impl(graph_name, palette))
    if z_min is not None and z_max is not None:
        messages.append(_set_colormap_levels_impl(graph_name, z_min, z_max))
    if levels is not None:
        messages.append(_set_colormap_level_count_impl(graph_name, levels))
    return " ".join(messages)


@mcp.tool()
def export_graph(
    graph_name: str,
    file_path: str,
    format: str = "png",
    width: int = 0,
    sized: bool = False,
) -> str:
    """Export a graph to an image file.

    Export never touches the Windows clipboard: the file is written directly
    via Origin's expGraph X-Function, so the user's clipboard is preserved.

    Args:
        graph_name: Graph to export.
        file_path: Output path (Windows or WSL style). Missing directories
            are created.
        format: Output format. Raster: png, jpg, tif, bmp. Vector (journal
            submission): pdf, eps, emf. Vector formats are resolution-
            independent and ignore width.
        width: Output raster pixel WIDTH. 0 (default) = auto ~1200px wide. Any
            value > 0 is applied — no sized=True needed. The HEIGHT always
            follows the graph's aspect ratio (Origin's expGraph cannot set an
            independent pixel height on this build), so there is no height
            parameter.
        sized: Kept for compatibility. When True, force the sized export even
            with the default width (~1200px wide). Redundant now that an
            explicit width is always honored.

    Returns:
        Path to the exported file.
    """
    if sized or width > 0:
        return _export_graph_sized_impl(
            graph_name, file_path, width=width or 1200, format=format
        )
    return _export_graph_impl(graph_name, file_path, format=format)
