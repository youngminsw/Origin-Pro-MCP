"""Enumeration hardening: the shared name helpers isolate corrupt entries, and
the plot source-column finder prefers the crash-safe LabTalk sheet enumeration
over the deep all-pages/.Layers COM traversal."""
from fakes import FakeOrigin, FakeBook, FakeSheet, FakeColumn, FakePages
from origin_pro_mcp import origin_connection


class _BoomName:
    @property
    def Name(self):
        raise RuntimeError("corrupt COM object")


def test_safe_page_names_isolates_bad_entry():
    good = FakeBook("Good")
    pages = FakePages([_BoomName(), good])
    assert origin_connection.safe_page_names(pages) == ["Good"]


def test_graph_names_isolated():
    class IsolationGraphs(FakeOrigin):
        @property
        def GraphPages(self):
            return FakePages([_BoomName(), self.graphs[0]])
    fake = IsolationGraphs()
    origin_connection.set_session_origin(fake)
    try:
        assert origin_connection.graph_names() == ["Graph1"]
    finally:
        origin_connection.clear_session_origin()


class LabTalkSheetOrigin(FakeOrigin):
    """LTStr resolves the LabTalk sheet-enum var, so sheet_names returns names
    WITHOUT any COM page/Layers traversal (the crash path)."""

    def __init__(self):
        super().__init__()
        self.books = [FakeBook("Bk", sheets=[FakeSheet(
            "S1", columns=[FakeColumn("A", col_type=3), FakeColumn("Pris", col_type=0)])])]
        self._active = ""
        self.layers_reads = 0

    def Execute(self, script):
        self.executed.append(script)
        if "win -a " in script and "_opm_sh$" in script:
            book = script.split("win -a ", 1)[1].split(";", 1)[0].strip()
            self._active = book
            m = next((b for b in self.books if b.Name == book), None)
            self._sheet_var = ("\x1e".join(s.Name for s in m.sheets) + "\x1e") if m else ""
        return True

    def LTStr(self, name):
        if name == "_opm_sh$":
            return getattr(self, "_sheet_var", "")
        return ""


def test_sheet_names_uses_labtalk():
    fake = LabTalkSheetOrigin()
    origin_connection.set_session_origin(fake)
    try:
        assert origin_connection.sheet_names("Bk") == ["S1"]
        assert any("layer$(_opmk).name$" in s for s in fake.executed)
    finally:
        origin_connection.clear_session_origin()


def test_find_source_column_prefers_labtalk_path():
    from origin_pro_mcp.tools.style_helpers import _find_source_column
    fake = LabTalkSheetOrigin()
    origin_connection.set_session_origin(fake)
    try:
        found = _find_source_column(fake, "Bk", "Pris")
        assert found is not None
        sheet_name, y_idx, x_idx, col = found
        assert sheet_name == "S1"
        assert col.Name == "Pris"
        assert x_idx == 0  # column A is the X column (col_type 3)
    finally:
        origin_connection.clear_session_origin()
