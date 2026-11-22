"""list_worksheets(details=True): per-sheet Rows/Cols and per-graph plot
counts read from COM object properties, with every read isolated (unknown ->
None, never a guess), and details=False byte-identical to the old output."""
import json

from fakes import FakeBook, FakeGraph, FakeOrigin, FakePages, FakeSheet
from origin_pro_mcp import origin_connection


def _call(fake, **kwargs):
    origin_connection.set_session_origin(fake)
    try:
        from origin_pro_mcp.tools.worksheet import list_worksheets
        return list_worksheets(**kwargs)
    finally:
        origin_connection.clear_session_origin()


class SizedSheet(FakeSheet):
    """A worksheet COM double that exposes Rows/Cols."""

    def __init__(self, name, rows, cols):
        super().__init__(name)
        self.Rows = rows
        self.Cols = cols


class BrokenRowsSheet(FakeSheet):
    """Rows raises (unverified property on this Origin); Cols works."""

    def __init__(self, name):
        super().__init__(name)
        self.Cols = 3

    @property
    def Rows(self):
        raise AttributeError("Rows")


class MultiLayerGraph(FakeGraph):
    """A graph page with several layers, each with its own plot count."""

    def __init__(self, name, per_layer):
        super().__init__(name)
        self._per_layer = per_layer

    @property
    def Layers(self):
        return FakePages([_CountLayer(n) for n in self._per_layer])


class _CountLayer:
    def __init__(self, n):
        self._n = n

    @property
    def DataPlots(self):
        if self._n is None:
            raise RuntimeError("DataPlots unreadable")
        return FakePages([object()] * self._n)


class BrokenLayersGraph(FakeGraph):
    @property
    def Layers(self):
        raise RuntimeError("corrupt page")


class _BadNamePage:
    @property
    def Name(self):
        raise RuntimeError("corrupt COM object")


# --------------------------------------------------------------------------- #
# details=False: unchanged                                                    #
# --------------------------------------------------------------------------- #


def test_details_false_output_is_byte_identical_to_old_version(fake_origin):
    # The exact string the pre-`details` list_worksheets produced for the
    # default fake (json.dumps of the same three keys, same order).
    old = ('{"workbooks": [{"name": "Book1", "sheets": ["Sheet1"]}], '
           '"graphs": ["Graph1"], "matrices": []}')
    assert _call(fake_origin) == old
    assert _call(fake_origin, details=False) == old


def test_details_false_has_no_detail_keys():
    fake = FakeOrigin()
    fake.books = [FakeBook("B", sheets=(SizedSheet("S", 10, 2),))]
    result = json.loads(_call(fake))
    assert set(result) == {"workbooks", "graphs", "matrices"}


# --------------------------------------------------------------------------- #
# sheet_details                                                               #
# --------------------------------------------------------------------------- #


def test_sheet_details_reads_rows_and_cols_when_present():
    fake = FakeOrigin()
    fake.books = [
        FakeBook("Bk1", sheets=(SizedSheet("Data", 1200, 4),
                                SizedSheet("My Sheet", 0, 2))),
        FakeBook("Bk2", sheets=(SizedSheet("Sheet1", 7, 1),)),
    ]
    result = json.loads(_call(fake, details=True))
    assert result["sheet_details"] == [
        {"book": "Bk1", "sheet": "Data", "rows": 1200, "cols": 4},
        {"book": "Bk1", "sheet": "My Sheet", "rows": 0, "cols": 2},
        {"book": "Bk2", "sheet": "Sheet1", "rows": 7, "cols": 1},
    ]
    # The plain listing is still there, unchanged.
    assert result["workbooks"][0] == {"name": "Bk1", "sheets": ["Data", "My Sheet"]}


def test_sheet_details_none_when_properties_missing(fake_origin):
    # Default FakeSheet has no Rows/Cols at all (unverified on Origin 2018).
    result = json.loads(_call(fake_origin, details=True))
    assert result["sheet_details"] == [
        {"book": "Book1", "sheet": "Sheet1", "rows": None, "cols": None},
    ]


def test_sheet_details_isolates_each_property():
    fake = FakeOrigin()
    fake.books = [FakeBook("B", sheets=(BrokenRowsSheet("S"),))]
    result = json.loads(_call(fake, details=True))
    assert result["sheet_details"] == [{"book": "B", "sheet": "S", "rows": None, "cols": 3}]


def test_sheet_details_none_when_find_worksheet_fails():
    class Boom(FakeOrigin):
        def FindWorksheet(self, target):
            raise RuntimeError("COM error")

    fake = Boom()
    fake.books = [FakeBook("B", sheets=(SizedSheet("S", 5, 5),))]
    result = json.loads(_call(fake, details=True))
    assert result["sheet_details"] == [{"book": "B", "sheet": "S", "rows": None, "cols": None}]


def test_sheet_details_never_reads_cell_data():
    class NoGetWorksheet(FakeOrigin):
        def GetWorksheet(self, target):
            raise AssertionError("details must not read cell data")

    fake = NoGetWorksheet()
    fake.books = [FakeBook("B", sheets=(SizedSheet("S", 5, 2),))]
    json.loads(_call(fake, details=True))  # must not raise


# --------------------------------------------------------------------------- #
# graph_details                                                               #
# --------------------------------------------------------------------------- #


def test_graph_details_counts_plots_and_flags_empty():
    fake = FakeOrigin()
    fake.graphs = [
        FakeGraph("Full", plot_names=("p1", "p2")),
        FakeGraph("Blank"),
        MultiLayerGraph("Panels", [1, 0, 3]),
    ]
    result = json.loads(_call(fake, details=True))
    assert result["graph_details"] == [
        {"name": "Full", "layers": 1, "plots": 2, "empty": False},
        {"name": "Blank", "layers": 1, "plots": 0, "empty": True},
        {"name": "Panels", "layers": 3, "plots": 4, "empty": False},
    ]
    assert result["graphs"] == ["Full", "Blank", "Panels"]


def test_graph_details_unknown_when_a_layer_fails():
    # One unreadable layer: never report the partial sum as the total.
    fake = FakeOrigin()
    fake.graphs = [MultiLayerGraph("G", [2, None, 1])]
    result = json.loads(_call(fake, details=True))
    assert result["graph_details"] == [
        {"name": "G", "layers": 3, "plots": None, "empty": None},
    ]


def test_graph_details_single_bad_page_does_not_break_the_rest():
    fake = FakeOrigin()
    fake.graphs = [
        FakeGraph("Ok", plot_names=("p",)),
        BrokenLayersGraph("Broken"),
        _BadNamePage(),
        FakeGraph("Ok2"),
    ]
    result = json.loads(_call(fake, details=True))
    assert result["graph_details"] == [
        {"name": "Ok", "layers": 1, "plots": 1, "empty": False},
        {"name": "Broken", "layers": None, "plots": None, "empty": None},
        {"name": "Ok2", "layers": 1, "plots": 0, "empty": True},
    ]
    # the plain list skips the unreadable page the same way
    assert result["graphs"] == ["Ok", "Broken", "Ok2"]


def test_graph_details_empty_list_when_graph_pages_unreadable():
    class NoGraphs(FakeOrigin):
        @property
        def GraphPages(self):
            raise RuntimeError("COM error")

    fake = NoGraphs()
    origin_connection.set_session_origin(fake)
    try:
        from origin_pro_mcp.tools.worksheet import _graph_details
        assert _graph_details(fake) == []
    finally:
        origin_connection.clear_session_origin()


def test_details_does_not_switch_windows():
    # details=True must issue exactly the same LabTalk as details=False: no
    # extra `win -a`, so the user's active window is untouched.
    def scripts(details):
        fake = FakeOrigin()
        fake.books = [FakeBook("B", sheets=(SizedSheet("S", 1, 1),))]
        fake.graphs = [FakeGraph("G", plot_names=("p",))]
        _call(fake, details=details)
        return fake.executed, fake.graphs[0].activated

    plain, _ = scripts(False)
    detailed, activated = scripts(True)
    assert detailed == plain
    assert activated is False
