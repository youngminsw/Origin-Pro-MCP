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


def test_find_peaks_on_short_gaussian_live(live_origin):
    """Item 28: find_peaks with the default local_points=10 now finds the peak
    of an 11-point Gaussian (clamped window) instead of failing."""
    import math

    from origin_pro_mcp.tools.analysis import transform
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("PKS"))
    book, sheet = made["name"], made["sheet"]
    xs = list(range(11))
    ys = [math.exp(-((x - 5) ** 2) / 2.0) for x in xs]  # peak at x=5
    set_worksheet_data(book, sheet, json.dumps([xs, ys]))

    out = json.loads(transform(book, sheet, 1, 2, method="find_peaks"))
    assert out["count"] >= 1, out
    assert out["local_points_used"] <= 5, out
    assert any(abs(p["x"] - 5) < 1.5 for p in out["peaks"]), out


def test_set_matrix_data_null_roundtrips_to_missing(live_origin):
    """Item 5: a null cell written into a matrix reads back as null (Origin's
    missing value), while surrounding real numbers are untouched."""
    from origin_pro_mcp.tools.matrix import (
        create_matrix,
        get_matrix_data,
        set_matrix_data,
    )

    made = json.loads(create_matrix("MTXNULL", rows=2, cols=3))
    book = made["name"]

    set_matrix_data(book, json.dumps([[1, None, 3], [4, 5, 6]]))
    out = json.loads(get_matrix_data(book))

    assert out["rows"][0][0] == 1.0
    assert out["rows"][0][1] is None
    assert out["rows"][0][2] == 3.0
    assert out["rows"][1] == [4.0, 5.0, 6.0]


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
