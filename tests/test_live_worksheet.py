"""Live (Windows + Origin Pro COM) tests for the g2-g8 sweep worksheet fixes:
JSON null -> Origin missing value in set_worksheet_data (#18), and
create_worksheet adding a sheet to an EXISTING workbook instead of spawning a
second, auto-renamed one (Misc).

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_worksheet.py -v

Safety: every test runs against its OWN isolated Origin instance spawned via
``DispatchEx("Origin.Application")`` (same pattern as test_live_loaded_graph.py
/ test_live_styling.py). Never touches ``Origin.ApplicationSI``.
``origin.Exit()`` always runs in the fixture teardown.
"""
import json

import pytest

pytestmark = pytest.mark.requires_origin


@pytest.fixture()
def live_origin():
    """A fresh, isolated Origin.exe bound to this thread; closed afterwards."""
    import pythoncom
    import win32com.client

    from origin_pro_mcp.origin_connection import (
        clear_session_origin,
        set_session_origin,
    )

    pythoncom.CoInitialize()

    def _factory():
        o = win32com.client.DispatchEx("Origin.Application")
        try:
            o.Visible = 1
        except Exception:
            pass
        return o

    origin = _factory()
    set_session_origin(origin, factory=_factory)
    try:
        yield origin
    finally:
        try:
            origin.Exit()
        except Exception:
            pass
        clear_session_origin()


def test_set_worksheet_data_null_becomes_missing_value(live_origin):
    from origin_pro_mcp.tools.worksheet import (
        create_worksheet,
        get_worksheet_data,
        set_worksheet_data,
    )

    made = json.loads(create_worksheet("WKST18"))
    book, sheet = made["name"], made["sheet"]

    set_worksheet_data(book, sheet, json.dumps([[1, 2, None, 4]]))
    out = json.loads(get_worksheet_data(book, sheet))

    assert out["columns"] == [[1.0, 2.0, None, 4.0]]


def test_create_worksheet_existing_book_adds_sheet_not_new_book(live_origin):
    from origin_pro_mcp.tools.worksheet import create_worksheet, list_worksheets

    made = json.loads(create_worksheet("WKSTMISC"))
    book = made["name"]

    out = json.loads(create_worksheet(book, "Sheet2"))
    assert out["name"] == book
    assert out["added_to_existing_book"] is True
    assert out["renamed"] is False

    listing = json.loads(list_worksheets())
    matches = [wb for wb in listing["workbooks"] if wb["name"] == book]
    assert len(matches) == 1, f"expected exactly one workbook named {book}, got {matches}"
    assert set(matches[0]["sheets"]) == {"Sheet1", "Sheet2"}
