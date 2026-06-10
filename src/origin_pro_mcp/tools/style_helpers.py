from ..labtalk_safe import labtalk_choice, labtalk_name, labtalk_string
from ..origin_connection import execute_labtalk, get_origin


def get_plot_names(graph_name: str) -> list:
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    gl = o.FindGraphLayer(f"[{safe_graph_name}]Layer1")
    if not gl:
        return []
    dp = gl.DataPlots
    return [dp.Item(i).Name for i in range(dp.Count)]


def graph_layer_execute(graph_name: str, script: str) -> bool:
    o = get_origin()
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    gl = o.FindGraphLayer(f"[{safe_graph_name}]Layer1")
    if not gl:
        return False
    return gl.Execute(script)


def set_legend_entries(graph_name: str, entries: list) -> None:
    plot_names = get_plot_names(graph_name)
    for i, entry in enumerate(entries):
        if i >= len(plot_names):
            break
        pname = plot_names[i]
        parts = pname.rsplit("_", 1)
        if len(parts) == 2:
            book, col = parts
            safe_book = labtalk_name(book, "plot_book")
            safe_col = labtalk_name(col, "plot_column")
            execute_labtalk(f"[{safe_book}]Sheet1!col({safe_col})[L]$ = {labtalk_string(entry, 'legend_entries')};")
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    execute_labtalk(f"win -a {safe_graph_name}; legend -r;")


def position_legend(graph_name: str, position: str) -> None:
    pos_fractions = {
        "top-left": (0.05, 0.85),
        "top-right": (0.65, 0.85),
        "bottom-left": (0.05, 0.15),
        "bottom-right": (0.65, 0.15),
    }
    safe_graph_name = labtalk_name(graph_name, "graph_name")
    safe_position = labtalk_choice(position, pos_fractions, "position")
    fx, fy = pos_fractions[safe_position]
    execute_labtalk(
        f"win -a {safe_graph_name}; "
        f"legend.x = layer.x.from + {fx} * (layer.x.to - layer.x.from); "
        f"legend.y = layer.y.from + {fy} * (layer.y.to - layer.y.from);"
    )
