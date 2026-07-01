"""list_worksheets must enumerate via the crash-safe LabTalk layer loop and
isolate a bad window, never doing the deep .Layers COM traversal that
hard-crashes Origin on heavy (60-window) projects."""
import json

from fakes import FakeOrigin, FakeBook, FakePages
from origin_pro_mcp import origin_connection


def _run(fake):
    origin_connection.set_session_origin(fake)
    try:
        from origin_pro_mcp.tools.worksheet import list_worksheets
        return json.loads(list_worksheets())
    finally:
        origin_connection.clear_session_origin()


class LabTalkEnumOrigin(FakeOrigin):
    """Models the real Origin LabTalk path: `win -a <book>` sets the active book
    and `_opm_sh$` resolves to that book's sheet names joined by the
    record-separator delimiter — so list_worksheets reads sheets WITHOUT the
    page.Layers COM traversal (the crash path)."""

    def __init__(self):
        super().__init__()
        self.books = [
            FakeBook("Bk1", sheets=("Sheet1", "Sheet2", "Sheet3")),
            FakeBook("Bk2", sheets=("Data",)),
        ]
        self._active = ""
        self._sheet_var = ""

    def Execute(self, script):
        self.executed.append(script)
        if "win -a " in script and "_opm_sh$" in script:
            book = script.split("win -a ", 1)[1].split(";", 1)[0].strip()
            self._active = book
            match = next((b for b in self.books if b.Name == book), None)
            self._sheet_var = (
                "\x1e".join(s.Name for s in match.sheets) + "\x1e" if match else ""
            )
        elif script.startswith("win -a "):
            self._active = script.split("win -a ", 1)[1].split(";", 1)[0].strip()
        return True

    def LTStr(self, name):
        if name == "_opm_sh$":
            return self._sheet_var
        if name == "_opm_act$":
            return self._active
        return ""


def test_list_worksheets_uses_labtalk_layer_enum():
    fake = LabTalkEnumOrigin()
    result = _run(fake)
    assert result["workbooks"] == [
        {"name": "Bk1", "sheets": ["Sheet1", "Sheet2", "Sheet3"]},
        {"name": "Bk2", "sheets": ["Data"]},
    ]
    # the crash-safe LabTalk layer loop was issued for each workbook
    assert any("layer$(_opmk).name$" in s for s in fake.executed)


def test_list_worksheets_falls_back_to_com_when_labtalk_empty(fake_origin):
    # default FakeOrigin: LTStr -> "" so the LabTalk path yields nothing and the
    # isolated per-sheet COM fallback supplies the names.
    result = _run(fake_origin)
    assert result["workbooks"] == [{"name": "Book1", "sheets": ["Sheet1"]}]
    assert result["graphs"] == ["Graph1"]
    assert result["matrices"] == []


class _Boom:
    @property
    def Name(self):
        raise RuntimeError("corrupt COM object")


class IsolationOrigin(FakeOrigin):
    """First workbook raises on .Name; enumeration must skip it, not abort."""

    @property
    def WorksheetPages(self):
        return FakePages([_Boom(), FakeBook("Good", sheets=("Sheet1",))])


def test_list_worksheets_isolates_unreadable_workbook():
    result = _run(IsolationOrigin())
    assert [w["name"] for w in result["workbooks"]] == ["Good"]
