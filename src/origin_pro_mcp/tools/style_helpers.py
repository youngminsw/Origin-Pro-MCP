from ..labtalk_safe import labtalk_choice, labtalk_name
from ..origin_connection import execute_labtalk, get_lt_var, get_origin

# COM Column.Type designation codes (Origin 2020 type library)
_COM_COLTYPE_Y_ERROR = 2


def get_plot_names(graph_name: str) -> list:
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    gl = o.FindGraphLayer(f"[{safe_graph_name}]Layer1")
    if not gl:
        return []
    dp = gl.DataPlots
    return [dp.Item(i).Name for i in range(dp.Count)]


def find_plot_column(plot_name: str):
    """COM Column object backing a plot named '<book>_<col short name>'.

    Returns None when the plot name doesn't follow that pattern or no
    matching column exists.
    """
    parts = plot_name.rsplit("_", 1)
    if len(parts) != 2:
        return None
    book, short_name = parts
    o = get_origin()
    pages = o.WorksheetPages
    for i in range(pages.Count):
        page = pages.Item(i)
        if page.Name != book:
            continue
        layers = page.Layers
        for j in range(layers.Count):
            # Layers items are generic Layer objects without Columns —
            # resolve each sheet through FindWorksheet instead
            sheet = o.FindWorksheet(f"[{book}]{layers.Item(j).Name}")
            if sheet is None:
                continue
            cols = sheet.Columns
            for k in range(cols.Count):
                col = cols.Item(k)
                if col.Name == short_name:
                    return col
    return None


def get_plot_info(graph_name: str) -> list:
    """[{"name", "is_error"}] for each plot in Layer1, in plot order.

    Error-bar plots appear in DataPlots like normal plots; they are
    recognized by their source column's Y-Error designation (set by
    create_graph/add_plot_to_graph when y_error_col is used).
    """
    infos = []
    for pname in get_plot_names(graph_name):
        col = find_plot_column(pname)
        is_error = col is not None and col.Type == _COM_COLTYPE_Y_ERROR
        infos.append({"name": pname, "is_error": is_error})
    return infos


def graph_layer_execute(graph_name: str, script: str) -> bool:
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    gl = o.FindGraphLayer(f"[{safe_graph_name}]Layer1")
    if not gl:
        return False
    return gl.Execute(script)


def set_legend_entries(graph_name: str, entries: list) -> None:
    """Set legend texts by writing column Long Names of the data plots.

    Error-bar plots are skipped so entries line up with data series.
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    data_plots = [p for p in get_plot_info(safe_graph_name) if not p["is_error"]]
    for entry, plot in zip(entries, data_plots):
        col = find_plot_column(plot["name"])
        if col is not None:
            col.LongName = entry
    execute_labtalk(f"win -a {safe_graph_name}; legend -r;")


def _read_legend_geometry(graph_name: str):
    """Axes ranges + legend box size/center, or None when unavailable."""
    execute_labtalk(f"win -a {graph_name};")
    values = {}
    reads = {
        "x_from": "layer.x.from",
        "x_to": "layer.x.to",
        "y_from": "layer.y.from",
        "y_to": "layer.y.to",
        "dx": "legend.dx",
        "dy": "legend.dy",
        "cx": "legend.x",
        "cy": "legend.y",
    }
    for key, expr in reads.items():
        if not execute_labtalk(f"__mcp_{key} = {expr};"):
            return None  # no legend or no axes — nothing to position
        values[key] = get_lt_var(f"__mcp_{key}")
    return values


def _place_legend(graph_name: str, position: str) -> None:
    """Place the legend fully inside the frame at the given corner.

    legend.x/y are the box CENTER in data coordinates, so the box size
    (legend.dx/dy) is read back to keep the whole box inside the axes
    with a small padding — otherwise wide legends cover the tick labels.
    """
    pad = 0.03
    values = _read_legend_geometry(graph_name)
    if values is None:
        return
    x_range = values["x_to"] - values["x_from"]
    y_range = values["y_to"] - values["y_from"]
    if "left" in position:
        cx = values["x_from"] + pad * x_range + values["dx"] / 2
    else:
        cx = values["x_to"] - pad * x_range - values["dx"] / 2
    if "top" in position:
        cy = values["y_to"] - pad * y_range - values["dy"] / 2
    else:
        cy = values["y_from"] + pad * y_range + values["dy"] / 2
    execute_labtalk(f"legend.x = {cx}; legend.y = {cy};")


def position_legend(graph_name: str, position: str) -> None:
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_position = labtalk_choice(
        position, {"top-left", "top-right", "bottom-left", "bottom-right"}, "position"
    )
    _place_legend(safe_graph_name, safe_position)


def reposition_legend_nearest_corner(graph_name: str) -> None:
    """Re-anchor the legend at whichever corner it currently sits in.

    Used after the legend is rebuilt (its box size changes), so a box
    that grew past the frame is pulled back inside.
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    values = _read_legend_geometry(safe_graph_name)
    if values is None:
        return
    x_mid = (values["x_from"] + values["x_to"]) / 2
    y_mid = (values["y_from"] + values["y_to"]) / 2
    horizontal = "left" if values["cx"] <= x_mid else "right"
    vertical = "top" if values["cy"] >= y_mid else "bottom"
    _place_legend(safe_graph_name, f"{vertical}-{horizontal}")
