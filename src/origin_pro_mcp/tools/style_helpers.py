from ..labtalk_safe import labtalk_choice, labtalk_name
from ..origin_connection import execute_labtalk, get_lt_var, get_origin

# COM Column.Type designation codes (Origin 2020 type library)
_COM_COLTYPE_Y_ERROR = 2
_COM_COLTYPE_X = 3


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



def _collect_xy(graph_name: str):
    """(x, y) points of every non-error data plot, for legend placement.

    Returns two parallel lists. Points whose X is unknown are dropped.
    Best-effort: silently skips plots whose source columns can't be read.
    """
    o = get_origin()
    xs, ys = [], []
    for info in get_plot_info(graph_name):
        if info["is_error"]:
            continue
        parts = info["name"].rsplit("_", 1)
        if len(parts) != 2:
            continue
        book, short = parts
        pages = o.WorksheetPages
        for i in range(pages.Count):
            page = pages.Item(i)
            if page.Name != book:
                continue
            for j in range(page.Layers.Count):
                sheet = o.FindWorksheet(f"[{book}]{page.Layers.Item(j).Name}")
                if sheet is None:
                    continue
                cols = sheet.Columns
                y_idx = x_idx = None
                for k in range(cols.Count):
                    col = cols.Item(k)
                    if col.Name == short:
                        y_idx = k
                    if col.Type == _COM_COLTYPE_X:
                        x_idx = k
                if y_idx is None:
                    continue
                data = o.GetWorksheet(f"[{book}]{sheet.Name}")
                if not isinstance(data, (list, tuple)):
                    continue
                for row in data:
                    try:
                        y = float(row[y_idx])
                    except (TypeError, ValueError, IndexError):
                        continue
                    try:
                        x = float(row[x_idx]) if x_idx is not None else None
                    except (TypeError, ValueError, IndexError):
                        x = None
                    xs.append(x)
                    ys.append(y)
    return xs, ys


def _corner_rect(values: dict, corner: str, pad: float = 0.03):
    """Data-coordinate rectangle the legend box occupies at a corner."""
    x_range = values["x_to"] - values["x_from"]
    y_range = values["y_to"] - values["y_from"]
    box_w = pad * x_range + values["dx"]
    box_h = pad * y_range + values["dy"]
    if "left" in corner:
        rx0, rx1 = values["x_from"], values["x_from"] + box_w
    else:
        rx0, rx1 = values["x_to"] - box_w, values["x_to"]
    if "top" in corner:
        ry0, ry1 = values["y_to"] - box_h, values["y_to"]
    else:
        ry0, ry1 = values["y_from"], values["y_from"] + box_h
    return rx0, rx1, ry0, ry1


def choose_legend_corner(graph_name: str, preferred: str | None = None) -> str:
    """Pick the corner whose legend box overlaps the fewest data points.

    Guarantees the legend never sits on the data when any corner is
    clear. Ties prefer `preferred`, then top-right, top-left, bottom-right,
    bottom-left. Falls back to `preferred` (or top-right) if geometry or
    data can't be read.
    """
    fallback = preferred or "top-right"
    values = _read_legend_geometry(graph_name)
    if values is None:
        return fallback
    xs, ys = _collect_xy(graph_name)
    points = [(x, y) for x, y in zip(xs, ys) if x is not None]
    if not points:
        return fallback
    corners = ["top-right", "top-left", "bottom-right", "bottom-left"]
    counts = {}
    for corner in corners:
        rx0, rx1, ry0, ry1 = _corner_rect(values, corner)
        counts[corner] = sum(
            1 for x, y in points if rx0 <= x <= rx1 and ry0 <= y <= ry1
        )
    fewest = min(counts.values())
    order = ([preferred] if preferred else []) + corners
    for corner in order:
        if counts.get(corner) == fewest:
            return corner
    return fallback


def place_legend_avoiding_data(graph_name: str, preferred: str | None = None) -> str:
    """Place the legend at the emptiest corner so it never covers data.

    Returns the corner used.
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    corner = choose_legend_corner(safe_graph_name, preferred)
    _place_legend(safe_graph_name, corner)
    return corner
