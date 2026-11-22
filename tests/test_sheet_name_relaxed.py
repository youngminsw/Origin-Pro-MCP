"""Relaxed sheet-name validation (origin2018-improvements).

``labtalk_sheet_name`` accepts EXISTING sheet names Origin allows (spaces,
CJK, dots, parentheses, ...) for call sites that hand the name to COM or quote
it with ``labtalk_string``. Range-interpolating call sites keep the strict
``labtalk_name`` check, whose error now tells the user to rename the sheet.
"""
import json

import pytest
from fakes import FakeBook, FakeSheet

from origin_pro_mcp.labtalk_safe import labtalk_name, labtalk_sheet_name


@pytest.mark.parametrize(
    "name",
    [
        "wpd_datasets _1_",   # the real Origin 2018 sheet that motivated this
        "Sheet1",
        "數據 一",             # CJK + space
        "量測資料",
        "run.2",               # dot
        "data (raw)",          # parentheses
        "a-b_c",
    ],
)
def test_labtalk_sheet_name_accepts_origin_sheet_names(name):
    assert labtalk_sheet_name(name, "sheet_name") == name


@pytest.mark.parametrize("name", ["", "   ", " lead", "trail ", "\tx"])
def test_labtalk_sheet_name_rejects_empty_or_padded(name):
    with pytest.raises(ValueError, match="sheet_name"):
        labtalk_sheet_name(name, "sheet_name")


def test_labtalk_sheet_name_rejects_non_string():
    with pytest.raises(ValueError, match="cannot be empty"):
        labtalk_sheet_name(None, "sheet_name")  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "name",
    [
        'a"b',          # would break page.active$ = "..."
        "a[b",
        "a]b",          # would break [Book]Sheet
        "a!b",          # range separator
        "a;b",          # statement separator
        "a\nb",
        "a\rb",
        "a\tb",
    ],
)
def test_labtalk_sheet_name_rejects_breakout_characters(name):
    with pytest.raises(ValueError, match="cannot contain"):
        labtalk_sheet_name(name, "sheet_name")


def test_labtalk_name_sheet_field_error_suggests_renaming():
    with pytest.raises(ValueError) as exc:
        labtalk_name("wpd_datasets _1_", "data_sheet")
    msg = str(exc.value)
    assert "must start with a letter" in msg
    assert "rename the sheet" in msg


def test_labtalk_name_non_sheet_field_has_no_rename_hint():
    with pytest.raises(ValueError) as exc:
        labtalk_name("bad name", "book_name")
    assert "rename the sheet" not in str(exc.value)


def test_labtalk_name_empty_sheet_field_has_no_rename_hint():
    # The hint only makes sense for a real (non-blank) name.
    with pytest.raises(ValueError) as exc:
        labtalk_name("", "sheet_name")
    assert "rename the sheet" not in str(exc.value)


# --- call-site behaviour --------------------------------------------------- #


def test_get_worksheet_data_accepts_sheet_with_space(fake_origin):
    from origin_pro_mcp.tools.worksheet import get_worksheet_data

    fake_origin.books = [FakeBook("Book1", sheets=("wpd_datasets _1_",))]
    out = json.loads(get_worksheet_data("Book1", "wpd_datasets _1_"))
    assert out["columns"] == [[1.0, 2.0], [4.0, 5.0]]


def test_relaxed_sort_quotes_sheet_name_in_labtalk(fake_origin):
    from origin_pro_mcp.tools.worksheet import sort_worksheet

    fake_origin.books = [FakeBook("Book1", sheets=("wpd_datasets _1_",))]
    sort_worksheet("Book1", "wpd_datasets _1_", 1)
    assert 'page.active$ = "wpd_datasets _1_";' in fake_origin.executed


def test_relaxed_analysis_tool_accepts_sheet_with_space(fake_origin):
    from origin_pro_mcp.tools.analysis import _column_statistics_impl

    fake_origin.books = [FakeBook("Book1", sheets=("my data",))]
    _column_statistics_impl("Book1", "my data", 1)
    assert 'page.active$ = "my data";' in fake_origin.executed


def test_relaxed_tool_still_rejects_bracket_in_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import get_worksheet_data

    with pytest.raises(ValueError, match="cannot contain"):
        get_worksheet_data("Book1", "Sheet1]x")


def test_range_tool_keeps_strict_check_with_rename_hint(fake_origin):
    # Interpolation (_interpolate_impl) builds [Book]Sheet!(x,y) ranges, so it
    # must stay strict even for a sheet that really exists.
    from origin_pro_mcp.tools.analysis import _interpolate_impl

    fake_origin.books = [FakeBook("Book1", sheets=(FakeSheet("my data"),))]
    with pytest.raises(ValueError, match="rename the sheet"):
        _interpolate_impl("Book1", "my data", 1, 2, 10)
