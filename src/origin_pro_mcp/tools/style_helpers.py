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


def choose_legend_corner_overlap(
    graph_name: str, preferred: str | None = None
) -> tuple[str, int]:
    """(emptiest corner, how many data points its legend box still covers).

    An overlap count of 0 means that corner is clean. When geometry or data
    can't be read, returns (fallback, 0) — i.e. assume clear, preserving the
    old corner-placement behavior.
    """
    fallback = preferred or "top-right"
    values = _read_legend_geometry(graph_name)
    if values is None:
        return fallback, 0
    xs, ys = _collect_xy(graph_name)
    points = [(x, y) for x, y in zip(xs, ys) if x is not None]
    if not points:
        return fallback, 0
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
            return corner, fewest
    return fallback, fewest


def choose_legend_corner(graph_name: str, preferred: str | None = None) -> str:
    """Pick the corner whose legend box overlaps the fewest data points."""
    return choose_legend_corner_overlap(graph_name, preferred)[0]


def _read_page_layout(graph_name: str):
    """Page size + plot-box geometry + legend size, or None when unavailable.

    `page.width/height` and `legend.left/top/width/height` are in page units;
    `layer.left/top/width/height` are in % of the page (layer.unit == 1).
    """
    execute_labtalk(f"win -a {graph_name};")
    reads = {
        "page_w": "page.width",
        "page_h": "page.height",
        "layer_left": "layer.left",
        "layer_top": "layer.top",
        "layer_w": "layer.width",
        "layer_h": "layer.height",
        "legend_w": "legend.width",
        "legend_h": "legend.height",
    }
    vals = {}
    for key, expr in reads.items():
        if not execute_labtalk(f"__mcp_{key} = {expr};"):
            return None
        vals[key] = get_lt_var(f"__mcp_{key}")
    if vals["page_w"] <= 0 or vals["layer_w"] <= 0:
        return None
    return vals


_OUTSIDE_RIGHT_EDGE_PCT = 66.0  # plot-box right edge after shrinking, % of page


def _place_legend_outside(graph_name: str) -> bool:
    """Shrink the plot and park the legend in the freed strip to the RIGHT.

    Used when no inside corner is clear. The plot box is shrunk to a fixed
    right edge (so repeated calls converge to the same layout — idempotent),
    then the legend is moved just past that edge and vertically centered.
    `legend.attach = 1` plus setting `legend.left` TWICE is required: Origin
    clamps the first assignment back inside the frame, and only the second
    escapes it (confirmed against Origin 2020). Returns True on success, False
    if geometry is unavailable or there isn't room (caller falls back to a
    corner).
    """
    vals = _read_page_layout(graph_name)
    if vals is None:
        return False
    page_w, page_h = vals["page_w"], vals["page_h"]
    legend_w, legend_h = vals["legend_w"], vals["legend_h"]
    new_w = _OUTSIDE_RIGHT_EDGE_PCT - vals["layer_left"]
    if new_w < 35.0:  # too cramped to free a usable strip
        return False
    execute_labtalk(f"win -a {graph_name}; layer.width = {new_w};")
    plot_right = _OUTSIDE_RIGHT_EDGE_PCT / 100.0 * page_w
    left = plot_right + 0.02 * page_w
    left = min(left, page_w - legend_w - 0.01 * page_w)  # keep on the page
    top = (vals["layer_top"] + vals["layer_h"] / 2.0) / 100.0 * page_h - legend_h / 2.0
    # First left-set clamps inside the frame; the second escapes it.
    execute_labtalk(f"legend.attach = 1; legend.left = {left};")
    execute_labtalk(f"legend.left = {left}; legend.top = {top};")
    return True


def place_legend_avoiding_data(graph_name: str, preferred: str | None = None) -> str:
    """Place the legend so it never covers data.

    Prefers the emptiest inside corner; when every corner still overlaps data,
    falls back to placing the legend OUTSIDE the frame on the right (shrinking
    the plot to make room). Returns the placement used ("top-right", …, or
    "outside-right").
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    corner, overlap = choose_legend_corner_overlap(safe_graph_name, preferred)
    if overlap > 0 and _place_legend_outside(safe_graph_name):
        return "outside-right"
    _place_legend(safe_graph_name, corner)
    return corner
