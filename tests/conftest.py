import importlib.util
import sys

import pytest

from origin_pro_mcp import origin_connection


WINDOWS_ORIGIN_AVAILABLE = (
    sys.platform == "win32" and importlib.util.find_spec("win32com") is not None
)


def pytest_configure(config: pytest.Config) -> None:
    config.addinivalue_line(
        "markers",
        "requires_origin: test requires Windows Python, pywin32, and Origin Pro COM",
    )


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    if WINDOWS_ORIGIN_AVAILABLE:
        return
    skip_origin = pytest.mark.skip(
        reason="requires Windows Python with pywin32 and Origin Pro COM"
    )
    for item in items:
        if "requires_origin" in item.keywords:
            item.add_marker(skip_origin)


# --- Shared fake Origin COM doubles (used by test_tool_guards, test_cli) ---


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
