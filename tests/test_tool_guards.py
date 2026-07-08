"""Unit tests for tool-level guards using a fake Origin COM object.

These run anywhere (no Windows/pywin32/Origin needed) and lock in the
fail-loudly behavior: tools must raise friendly errors instead of
returning success messages when Origin reports failure.
"""
import json

import pytest

from origin_pro_mcp import origin_connection

from conftest import FakeColumn, FakeSheet, FakeBook, FakeGraph


def test_activate_window_raises_with_open_windows(fake_origin):
    fake_origin.execute_results["win -a Nope"] = False
    with pytest.raises(ValueError, match="Open workbooks: Book1"):
        origin_connection.activate_window("Nope")


def test_load_project_missing_file_never_reaches_origin(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import load_project

    with pytest.raises(ValueError, match="not found"):
        load_project(str(tmp_path / "missing.opju"))
    # Load must not be called: a failed Load can clear the open project
    assert fake_origin.load_result is True


def test_save_project_reports_origin_failure(fake_origin):
    from origin_pro_mcp.tools.project import save_project

    fake_origin.save_result = False
    with pytest.raises(ValueError, match="no file location"):
        save_project()


def test_save_project_rejects_wrong_extension(fake_origin):
    from origin_pro_mcp.tools.project import save_project

    with pytest.raises(ValueError, match=".opj"):
        save_project("C:\\data\\project.png")


def test_windows_path_converts_wsl_style():
    from origin_pro_mcp.labtalk_safe import windows_path

    assert windows_path("/mnt/c/Users/me/fig.png", "p") == "C:\\Users\\me\\fig.png"
    assert windows_path("C:\\Users\\me\\fig.png", "p") == "C:\\Users\\me\\fig.png"


def test_get_worksheet_data_handles_hresult_return(fake_origin):
    from origin_pro_mcp.tools.worksheet import get_worksheet_data

    with pytest.raises(ValueError, match="not found") as exc_info:
        get_worksheet_data("NoBook", "NoSheet")
    assert "Book1" in str(exc_info.value)


def test_set_worksheet_data_rejects_bad_json(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    with pytest.raises(ValueError, match="JSON array of arrays"):
        set_worksheet_data("Book1", "Sheet1", "not json")


def test_set_worksheet_data_rejects_non_numeric(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    with pytest.raises(ValueError, match="non-numeric"):
        set_worksheet_data("Book1", "Sheet1", '[["a","b"]]')


def test_set_worksheet_data_accepts_flat_array(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    msg = set_worksheet_data("Book1", "Sheet1", "[1,2,3]")
    assert "1 columns x 3 rows" in msg


def test_set_worksheet_data_unknown_book_lists_open_ones(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    with pytest.raises(ValueError, match="Open workbooks: Book1"):
        set_worksheet_data("Ghost", "Sheet1", "[[1,2]]")


def test_list_worksheets_returns_books_sheets_graphs(fake_origin):
    from origin_pro_mcp.tools.worksheet import list_worksheets

    result = json.loads(list_worksheets())
    assert result["workbooks"] == [{"name": "Book1", "sheets": ["Sheet1"]}]
    assert result["graphs"] == ["Graph1"]
    assert result["matrices"] == []


def test_create_graph_unknown_worksheet(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    with pytest.raises(ValueError, match="Worksheet \\[Ghost\\]Sheet1 not found"):
        create_graph("Fig1", "Ghost", "Sheet1", 1, 2)


def test_create_graph_cleans_up_when_plot_fails(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    fake_origin.execute_results["plotxy"] = False
    with pytest.raises(ValueError, match="Could not plot"):
        create_graph("Fig1", "Book1", "Sheet1", 1, 2)
    assert any(s.startswith("win -cd Fig1") for s in fake_origin.executed)


def test_add_plot_unknown_graph_lists_open_graphs(fake_origin):
    from origin_pro_mcp.tools.graph import add_plot_to_graph

    with pytest.raises(ValueError, match="Open graphs: Graph1"):
        add_plot_to_graph("Ghost", "Book1", "Sheet1", 1, 2)


def test_import_csv_missing_file(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    with pytest.raises(ValueError, match="File not found"):
        import_data(str(tmp_path / "missing.csv"))


def test_export_graph_unknown_graph(fake_origin, tmp_path):
    from origin_pro_mcp.tools.graph import export_graph

    with pytest.raises(ValueError, match="Open graphs: Graph1"):
        export_graph("Ghost", str(tmp_path / "fig.png"))


def test_export_all_graphs_empty_project(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import export_all_graphs

    fake_origin.graphs = []
    assert "No graphs" in export_all_graphs(str(tmp_path))


def test_curve_fit_unknown_function_lists_options(fake_origin):
    from origin_pro_mcp.tools.fitting import curve_fit

    with pytest.raises(ValueError, match="function must be one of"):
        curve_fit("Book1", "Sheet1", 1, 2, function="quadratic")


def test_curve_fit_nlbegin_failure_is_reported(fake_origin):
    from origin_pro_mcp.tools.fitting import curve_fit

    fake_origin.execute_results["nlbegin"] = False
    with pytest.raises(ValueError, match="Could not start the 'gauss' fit"):
        curve_fit("Book1", "Sheet1", 1, 2, function="gauss")
    assert any(s.startswith("nlend") for s in fake_origin.executed)


def test_curve_fit_plot_on_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.fitting import curve_fit

    with pytest.raises(ValueError, match="Open graphs: Graph1"):
        curve_fit("Book1", "Sheet1", 1, 2, function="gauss", plot_on_graph="Ghost")


def _book_with_error_column():
    sheet = FakeSheet(
        "Sheet1",
        columns=[
            FakeColumn("A", col_type=3),  # X
            FakeColumn("B", col_type=0),  # Y data
            FakeColumn("C", col_type=2),  # Y error
        ],
    )
    return FakeBook("Book1", sheets=[sheet])


def test_get_plot_info_classifies_error_plots(fake_origin):
    from origin_pro_mcp.tools.style_helpers import get_plot_info

    fake_origin.books = [_book_with_error_column()]
    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B", "Book1_C"])]

    infos = get_plot_info("Graph1")
    assert infos == [
        {"name": "Book1_B", "is_error": False},
        {"name": "Book1_C", "is_error": True},
    ]


def test_set_legend_entries_skips_error_plots(fake_origin):
    from origin_pro_mcp.tools.style_helpers import set_legend_entries

    book = _book_with_error_column()
    fake_origin.books = [book]
    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B", "Book1_C"])]

    set_legend_entries("Graph1", ["Pristine"])
    columns = book.sheets[0].columns
    assert columns[1].LongName == "Pristine"  # data column renamed
    assert columns[2].LongName == ""  # error column untouched


def test_position_legend_keeps_box_inside_frame(fake_origin):
    from origin_pro_mcp.tools.style_helpers import position_legend

    fake_origin.lt_vars = {
        "__mcp_x_from": 0.0,
        "__mcp_x_to": 10.0,
        "__mcp_y_from": 0.0,
        "__mcp_y_to": 2.0,
        "__mcp_dx": 4.0,
        "__mcp_dy": 0.5,
    }
    position_legend("Graph1", "top-left")
    # center = from + 3% padding + half the box size, so the box edge
    # never covers the axis or tick labels
    assert "legend.x = 2.3; legend.y = 1.69;" in fake_origin.executed


def test_apply_publication_style_rejects_bad_legend_position(fake_origin):
    from origin_pro_mcp.tools.style import apply_publication_style

    # Invalid legend_position must fail fast, before any styling commands
    # mutate the graph.
    with pytest.raises(ValueError, match="legend_position"):
        apply_publication_style("Graph1", legend_position="center")
    assert not any(s.startswith("xb.text$") for s in fake_origin.executed)


def test_create_graph_contour_requires_z_col(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    with pytest.raises(ValueError, match="requires z_col"):
        create_graph("G", "Book1", "Sheet1", 1, 2, plot_type="contour")


def test_create_graph_histogram_plots_single_column(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    msg = create_graph("G", "Book1", "Sheet1", 1, 2, plot_type="histogram")
    assert "histogram" in msg
    # Y-range plot: a single column, with the corrected histogram ID 219
    assert any("col(2)" in s and "plot:=219" in s for s in fake_origin.executed)


def test_create_graph_uses_corrected_area_id(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    create_graph("G", "Book1", "Sheet1", 1, 2, plot_type="area")
    assert any("plot:=204" in s for s in fake_origin.executed)  # 204 = Area


def test_add_plot_rejects_non_xy_type(fake_origin):
    from origin_pro_mcp.tools.graph import add_plot_to_graph

    with pytest.raises(ValueError, match="only X,Y"):
        add_plot_to_graph("Graph1", "Book1", "Sheet1", 1, 2, plot_type="histogram")


# --- publication-figure quality rules ----------------------------------------

def test_nice_increment_caps_at_six_ticks():
    from origin_pro_mcp.tools.style import _nice_increment

    # Every non-None increment must yield at most 5 intervals (<= 6 major ticks).
    for lo, hi in [(0, 10), (0, 5), (0, 1), (0, 100), (-5, 5), (0, 8), (0, 60), (0, 0.4)]:
        inc = _nice_increment(lo, hi)
        if inc is not None:
            intervals = abs(hi - lo) / inc
            assert intervals <= 5 + 1e-9, (lo, hi, inc, intervals)


def test_nice_increment_tighter_than_old_eight_interval_cap():
    from origin_pro_mcp.tools.style import _nice_increment

    # span 60: the old 3-8 rule picked inc=10 (6 intervals); the capped rule
    # must give <= 5 intervals.
    inc = _nice_increment(0, 60)
    assert inc is not None
    assert (60 / inc) <= 5


def test_apply_publication_style_tightens_axes_to_data(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style

    monkeypatch.setattr(style, "_collect_xy",
                        lambda g: ([0.0, 5.0, 10.0], [2.0, 8.0, 20.0]))
    scripts = []
    monkeypatch.setattr(style, "graph_layer_execute",
                        lambda g, s: scripts.append(s) or True)

    style.apply_publication_style("Graph1")
    joined = " ".join(scripts)
    # Tight to the data extent: no padding before/after the data.
    assert "layer.x.from = 0.0;" in joined
    assert "layer.x.to = 10.0;" in joined
    assert "layer.x.inc" in joined
    assert "layer.y.from = 2.0;" in joined
    assert "layer.y.to = 20.0;" in joined
    assert "layer.y.inc" in joined
    # Minor ticks reduced to 1 (not 4) so ticks are not dense.
    assert "layer.x.minor = 1;" in joined
    assert "layer.y.minor = 1;" in joined


def test_apply_publication_style_uses_explicit_axis_bounds(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style

    monkeypatch.setattr(style, "_collect_xy",
                        lambda g: ([0.0, 10.0], [0.0, 100.0]))
    scripts = []
    monkeypatch.setattr(style, "graph_layer_execute",
                        lambda g, s: scripts.append(s) or True)

    style.apply_publication_style("Graph1", x_min=1.0, x_max=9.0)
    joined = " ".join(scripts)
    # Explicit bounds win over the data extent.
    assert "layer.x.from = 1.0;" in joined
    assert "layer.x.to = 9.0;" in joined
    # Y still tightens to the data.
    assert "layer.y.from = 0.0;" in joined
    assert "layer.y.to = 100.0;" in joined


def test_apply_publication_style_skips_axis_when_data_unreadable(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style

    def _boom(_g):
        raise RuntimeError("cannot read data")

    monkeypatch.setattr(style, "_collect_xy", _boom)
    scripts = []
    monkeypatch.setattr(style, "graph_layer_execute",
                        lambda g, s: scripts.append(s) or True)

    # Best-effort: an unreadable graph must not raise, and no from/to is set.
    style.apply_publication_style("Graph1")
    joined = " ".join(scripts)
    assert "layer.x.from" not in joined
    assert "layer.y.from" not in joined


def test_delete_graph_closes_window(fake_origin):
    from origin_pro_mcp.tools.graph import delete_graph

    msg = delete_graph("Graph1")
    assert "Deleted graph 'Graph1'." == msg
    assert any(s.startswith("win -cd Graph1") for s in fake_origin.executed)


def test_delete_graph_unknown_lists_open_graphs(fake_origin):
    from origin_pro_mcp.tools.graph import delete_graph

    with pytest.raises(ValueError, match="Open graphs: Graph1"):
        delete_graph("Ghost")
