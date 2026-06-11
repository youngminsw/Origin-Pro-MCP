"""Unit tests for tool-level guards using a fake Origin COM object.

These run anywhere (no Windows/pywin32/Origin needed) and lock in the
fail-loudly behavior: tools must raise friendly errors instead of
returning success messages when Origin reports failure.
"""
import json

import pytest

from origin_pro_mcp import origin_connection


class FakePages:
    def __init__(self, pages):
        self._pages = pages

    @property
    def Count(self):
        return len(self._pages)

    def Item(self, i):
        return self._pages[i]


class FakeColumn:
    def __init__(self, name, col_type=0, long_name=""):
        self.Name = name
        self.Type = col_type  # COM designation: 0=Y, 2=Y Error, 3=X
        self.LongName = long_name


class FakeSheet:
    def __init__(self, name, columns=()):
        self.Name = name
        self.columns = list(columns)

    @property
    def Columns(self):
        return FakePages(self.columns)


class FakeBook:
    def __init__(self, name, sheets=("Sheet1",)):
        self.Name = name
        self.sheets = [
            s if isinstance(s, FakeSheet) else FakeSheet(s) for s in sheets
        ]

    @property
    def Layers(self):
        return FakePages(self.sheets)


class FakePlot:
    def __init__(self, name):
        self.Name = name


class FakeLayer:
    def __init__(self, plot_names):
        self.DataPlots = FakePages([FakePlot(n) for n in plot_names])

    def Execute(self, script):
        return True


class FakeGraph:
    def __init__(self, name, plot_names=()):
        self.Name = name
        self.plot_names = list(plot_names)


class FakeOrigin:
    """Mimics the Origin COM surface the tools rely on."""

    def __init__(self):
        self.books = [FakeBook("Book1")]
        self.graphs = [FakeGraph("Graph1")]
        self.execute_results = {}
        self.executed = []
        self.save_result = True
        self.load_result = True
        self.put_result = True
        self.worksheet_data = ((1.0, 4.0), (2.0, 5.0))
        self.lt_vars = {}

    @property
    def WorksheetPages(self):
        return FakePages(self.books)

    @property
    def GraphPages(self):
        return FakePages(self.graphs)

    def Execute(self, script):
        self.executed.append(script)
        for prefix, result in self.execute_results.items():
            if script.startswith(prefix):
                return result
        return True

    def FindWorksheet(self, target):
        for book in self.books:
            for j in range(book.Layers.Count):
                sheet = book.Layers.Item(j)
                if target == f"[{book.Name}]{sheet.Name}":
                    return sheet
        return None

    def FindGraphLayer(self, target):
        for graph in self.graphs:
            if target == f"[{graph.Name}]Layer1":
                return FakeLayer(graph.plot_names)
        return None

    def GetWorksheet(self, target):
        if self.FindWorksheet(target) is None:
            return -2147352568  # HRESULT int, as observed on Origin 2020
        return self.worksheet_data

    def PutWorksheet(self, target, data, row, col):
        return self.put_result

    def Save(self, path):
        return self.save_result

    def Load(self, path):
        return self.load_result

    def CreatePage(self, kind, name, template):
        return name

    def LTVar(self, name):
        return self.lt_vars.get(name, 0.0)

    def LTStr(self, name):
        return ""


@pytest.fixture
def fake_origin(monkeypatch):
    fake = FakeOrigin()
    monkeypatch.setattr(origin_connection, "_origin", fake)
    return fake


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

    result = json.loads(get_worksheet_data("NoBook", "NoSheet"))
    assert "not found" in result["error"]
    assert "Book1" in result["error"]


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
    from origin_pro_mcp.tools.worksheet import import_csv_to_worksheet

    with pytest.raises(ValueError, match="File not found"):
        import_csv_to_worksheet(str(tmp_path / "missing.csv"))


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
