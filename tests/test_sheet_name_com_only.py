"""Sheet names with LabTalk substitution syntax (``%``, ``$(``).

LabTalk rewrites them inside quoted strings, so a tool that quotes the sheet
name into LabTalk (``page.active$ = "..."``) must reject them. A tool that only
hands ``[Book]Sheet`` to COM is not affected and keeps accepting such sheets:
get_worksheet_data, export_worksheet, and set_worksheet_data when it writes
plain numbers without column names.
"""
import json

import pytest
from fakes import FakeBook

from origin_pro_mcp.labtalk_safe import labtalk_sheet_name

NAMES = ["10% load", "cost $(x)"]


@pytest.mark.parametrize("name", NAMES)
def test_com_only_validation_accepts_substitution_syntax(name):
    assert labtalk_sheet_name(name, "sheet_name", in_labtalk=False) == name


@pytest.mark.parametrize("name", NAMES)
def test_labtalk_validation_rejects_substitution_syntax(name):
    with pytest.raises(ValueError, match="quotes the sheet name into LabTalk"):
        labtalk_sheet_name(name, "sheet_name")


@pytest.mark.parametrize("name", ["a]b", 'a"b', "a;b"])
def test_com_only_validation_still_rejects_breakout_characters(name):
    with pytest.raises(ValueError, match="cannot contain"):
        labtalk_sheet_name(name, "sheet_name", in_labtalk=False)


def _labtalk_touched_sheet(fake_origin, name):
    return [s for s in fake_origin.executed if name in s or "page.active$" in s]


def test_get_worksheet_data_reads_a_percent_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import get_worksheet_data

    fake_origin.books = [FakeBook("Book1", sheets=("10% load",))]
    out = json.loads(get_worksheet_data("Book1", "10% load"))
    assert out["columns"] == [[1.0, 2.0], [4.0, 5.0]]
    assert _labtalk_touched_sheet(fake_origin, "10% load") == []


def test_export_worksheet_reads_a_percent_sheet_over_com_only(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import export_worksheet

    fake_origin.books = [FakeBook("Book1", sheets=("10% load",))]
    out_file = tmp_path / "out.csv"
    msg = export_worksheet("Book1", "10% load", str(out_file))
    assert "2 rows x 2 cols" in msg
    lines = out_file.read_text(encoding="utf-8").splitlines()
    assert lines[0] == "Col1,Col2"           # header falls back to column index
    assert lines[1:] == ["1.0,4.0", "2.0,5.0"]
    # No window activation and no LabTalk string carrying the name.
    assert _labtalk_touched_sheet(fake_origin, "10% load") == []
    assert not any(s.startswith("win -a") for s in fake_origin.executed)


@pytest.fixture
def recorded_puts(fake_origin):
    puts = []

    def put(target, data, row, col):
        puts.append((target, col))
        return True

    fake_origin.PutWorksheet = put
    fake_origin.books = [FakeBook("Book1", sheets=("10% load",))]
    return puts


def test_set_worksheet_data_plain_numbers_accept_a_percent_sheet(fake_origin, recorded_puts):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    set_worksheet_data("Book1", "10% load", "[[1,2],[3,4]]")
    assert recorded_puts == [("[Book1]10% load", 0), ("[Book1]10% load", 1)]


@pytest.mark.parametrize("columns, names", [("[[1,null]]", ""), ("[[1,2]]", "X")])
def test_set_worksheet_data_needing_labtalk_rejects_before_any_write(
        fake_origin, recorded_puts, columns, names):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    with pytest.raises(ValueError, match="quotes the sheet name into LabTalk"):
        set_worksheet_data("Book1", "10% load", columns, column_names=names)
    assert recorded_puts == []                  # nothing half-written
