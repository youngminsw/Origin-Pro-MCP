from ..labtalk_safe import labtalk_choice, labtalk_name
from ..origin_connection import (
    activate_window, execute_labtalk, get_lt_var, get_origin, graph_names,
    sheet_names,
)

# COM Column.Type designation codes (Origin 2020 type library)
_COM_COLTYPE_Y_ERROR = 2
_COM_COLTYPE_X = 3


def _find_layer_com(o, safe_graph_name: str):
    """A FRESH Layer1 COM handle for a graph, via a fallback chain.

    1. ``FindGraphLayer("[name]Layer1")`` — the direct route.
    2. Walk the ``GraphPages`` collection for the page by name and take its
       first layer.

    Never caches — a stale handle is exactly what freezes graphs loaded from a
    project file. Returns the layer or None. The second path is LIVE-UNVERIFIED
    (defensive backstop for loaded pages that don't resolve through
    ``FindGraphLayer``)."""
    gl = o.FindGraphLayer(f"[{safe_graph_name}]Layer1")
    if gl is not None:
        return gl
    try:
        pages = o.GraphPages
        count = pages.Count
    except Exception:
        return None
    for i in range(count):
        try:
            page = pages.Item(i)
            if page.Name != safe_graph_name:
                continue
            layers = page.Layers
            if layers.Count:
                return layers.Item(0)
        except Exception:
            continue
    return None


def acquire_graph_layer(graph_name: str, *, activate: bool = True):
    """THE single fresh-handle acquisition path for every graph-targeting op.

    Given a graph name it (1) activates the page (``win -a``), then (2) takes a
    FRESH Layer1 COM handle (never cached) through ``_find_layer_com``. The
    activation is what un-freezes graphs loaded from a ``.opju``: their layer
    reports zero DataPlots and silently no-ops ``layer.*`` / per-plot commands
    until the page is the active window, and only a handle taken AFTER
    activation sees the real plots. Raises ValueError (with the open-graph
    list) if the graph can't be resolved — never returns a dead handle a caller
    would silently no-op against.
    """
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    o = get_origin()
    if activate:
        # Raises with the open-window list on a real Origin if the name is
        # unknown; a fresh handle is then taken against the now-active page.
        activate_window(safe_graph_name, "graph_name")
    gl = _find_layer_com(o, safe_graph_name)
    if gl is None:
        graphs = ", ".join(graph_names()) or "(none)"
        msg = (
            f"Could not acquire a graph layer for '{safe_graph_name}'. "
            f"Open graphs: {graphs}."
        )
        raise ValueError(msg)
    return gl


def _plot_names_from_layer(gl) -> list:
    out: list = []
    try:
        dp = gl.DataPlots
        count = dp.Count
    except Exception:
        return out
    for i in range(count):
        try:
            out.append(dp.Item(i).Name)
        except Exception:
            continue  # isolate a plot whose name can't be read
    return out


def get_plot_names(graph_name: str) -> list:
    """Names of the data plots on a graph's first layer, in plot order.

    Acquires the layer through ``acquire_graph_layer`` — the page is activated
    and a fresh handle taken FIRST. On a graph loaded from a ``.opju`` the
    DataPlots collection reports zero until its page is active, so enumerating
    without activating is exactly what returned an empty (frozen) plot list and
    made per-curve edits silently no-op.
    """
    gl = acquire_graph_layer(graph_name)
    return _plot_names_from_layer(gl)


def _com_book_sheet_names(o, book: str) -> list:
    """Isolated COM fallback for sheet names of one book — used only when the
    LabTalk enumeration yields nothing (test fakes / odd COM builds). Bounded to
    the matching book and isolates each item, so it cannot abort the caller."""
    names: list = []
    try:
        pages = o.WorksheetPages
        count = pages.Count
    except Exception:
        return names
    for i in range(count):
        try:
            page = pages.Item(i)
            if page.Name != book:
                continue
            layers = page.Layers
            for j in range(layers.Count):
                try:
                    names.append(layers.Item(j).Name)
                except Exception:
                    continue
        except Exception:
            continue
    return names


def _find_source_column(o, book: str, short: str):
    """Locate a plot's source column ``<book>_<short>`` WITHOUT the deep
    all-pages + page.Layers COM traversal that can crash Origin on heavy
    projects. Uses the crash-safe shared ``sheet_names`` (LabTalk) + direct
    ``FindWorksheet``, isolating each column read.

    Returns (sheet_name, y_index, x_index, col_object) or None.
    """
    for sheet in (sheet_names(book) or _com_book_sheet_names(o, book)):
        ws = o.FindWorksheet(f"[{book}]{sheet}")
        if ws is None:
            continue
        try:
            cols = ws.Columns
            ncols = cols.Count
        except Exception:
            continue
        y_idx = x_idx = None
        col_obj = None
        for k in range(ncols):
            try:
                col = cols.Item(k)
                name = col.Name
                ctype = col.Type
            except Exception:
                continue  # isolate a corrupt/unreadable column
            if name == short:
                y_idx, col_obj = k, col
            if ctype == _COM_COLTYPE_X:
                x_idx = k
        if col_obj is not None:
            return sheet, y_idx, x_idx, col_obj
    return None


def find_plot_column(plot_name: str):
    """COM Column object backing a plot named '<book>_<col short name>'.

    Returns None when the plot name doesn't follow that pattern or no
    matching column exists.
    """
    parts = plot_name.rsplit("_", 1)
    if len(parts) != 2:
        return None
    book, short_name = parts
    found = _find_source_column(get_origin(), book, short_name)
    return found[3] if found is not None else None


def settle_new_plots(graph_name: str, expected_min_plots: int, timeout_s: float = 4.0) -> None:
    """Settle barrier for a freshly-plotted graph page.

    A graph page immediately after ``CreatePage``/``plotxy``/
    ``add_plot_to_graph`` can silently ignore or partially apply the FIRST
    styling/read/export command issued against it — no exception, just a
    no-op (probe-confirmed; distinct from the loaded-.opju freeze that
    ``require_data_plots`` guards). Polls ``get_plot_info`` until at least
    ``expected_min_plots`` plots enumerate (cheap/instant against fakes, so
    this never slows the WSL test suite), then adds a short fixed settle
    tail. The tail is SKIPPED when the plots were already there on the very
    first poll (the page was already settled) to keep the common case cheap.
    Call this once, right after building/rebuilding a page's plots, before
    the first styling command touches it.
    """
    import time

    start = time.monotonic()
    first = True
    while True:
        try:
            infos = get_plot_info(graph_name)
        except Exception:
            # Graph not resolvable at all — not a timing issue, so polling
            # longer won't help (this is what a freshly CreatePage'd page
            # unknown to a test double hits; a real graph always resolves).
            return
        if len(infos) >= expected_min_plots:
            if not first:
                time.sleep(0.3)
            return
        first = False
        if time.monotonic() - start >= timeout_s:
            return
        time.sleep(0.1)


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


def require_data_plots(graph_name: str) -> tuple[list, list]:
    """(all plot infos, data-plot names) for a graph, RAISING if it has no
    editable data plots once its page has been activated.

    The anti-silent-no-op guard for per-curve edits: a multi-series graph
    loaded from a ``.opju`` can report zero DataPlots (and quietly ignore every
    style command) until its page is activated and a fresh handle taken — which
    ``get_plot_info`` now does. If plots are STILL zero we refuse instead of
    returning a fake success.
    """
    infos = get_plot_info(graph_name)
    data_names = [p["name"] for p in infos if not p["is_error"]]
    if not data_names:
        msg = (
            f"No editable data plots found on '{graph_name}' — its layer "
            f"reports zero plots even after activating the window. This graph "
            f"was most likely loaded from a project file (.opju) in a state "
            f"Origin will not expose over COM. Recreate it in-session "
            f"(create_graph / plotxy) or reopen the project fresh; refusing to "
            f"report success on an edit that would silently do nothing."
        )
        raise ValueError(msg)
    return infos, data_names


def graph_layer_execute(graph_name: str, script: str) -> bool:
    """Run a LabTalk script against a graph's Layer1 COM object.

    Routes through ``acquire_graph_layer`` so the page is activated and a
    FRESH layer handle is taken before every call — this is what unfreezes
    ``layer.*`` and per-plot commands on graphs loaded from a project file.
    """
    gl = acquire_graph_layer(graph_name)
    return gl.Execute(script)


def read_layer_value(graph_name: str, expr: str):
    """Read a numeric ``layer.*`` property back through a fresh layer handle.

    Returns a float, or None when the read-back is unavailable (so callers do
    not raise a false alarm). Used to confirm a mutation actually took effect.
    """
    var = "__mcp_rb"
    if not graph_layer_execute(graph_name, f"{var} = {expr};"):
        return None
    try:
        return float(get_lt_var(var))
    except Exception:
        return None


def verify_layer_value(graph_name: str, prop: str, expected: float, label: str) -> None:
    """Read ``prop`` back after setting it and raise if it did not change.

    Turns the loaded-graph silent no-op into a loud, actionable failure. A tiny
    relative+absolute tolerance absorbs float round-trip noise. If the property
    can't be read back at all, this is a no-op (preconditions still guard the
    freeze case)."""
    got = read_layer_value(graph_name, prop)
    if got is None:
        return
    tol = max(1e-6, abs(expected) * 1e-6)
    if abs(got - expected) > tol:
        msg = (
            f"{label} did not take effect on '{graph_name}' (set {expected}, "
            f"read back {got}). The layer ignored the change — the graph was "
            f"most likely loaded from a .opju in a state Origin freezes over "
            f"COM; recreate it in-session with create_graph / plotxy."
        )
        raise ValueError(msg)


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
        found = _find_source_column(o, book, short)
        if found is None:
            continue
        sheet_name, y_idx, x_idx, _col = found
        if y_idx is None:
            continue
        data = o.GetWorksheet(f"[{book}]{sheet_name}")
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
