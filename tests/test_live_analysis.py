"""Live (Windows + Origin Pro COM) tests for transform orphan hygiene
(item 11): interpolate reuses a stable book; find_peaks reuses its two output
columns — neither grows the project on repeat calls.

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_analysis.py -v

Safety: an isolated ``DispatchEx("Origin.Application")`` per test; never
``Origin.ApplicationSI``; ``origin.Exit()`` in teardown.
"""
import json
import math

import pytest

pytestmark = pytest.mark.requires_origin


@pytest.fixture()
def live_origin():
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


def _spectrum(book="ORPHSPEC"):
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet(book))
    b, sh = made["name"], made["sheet"]
    xs = [i * 0.5 for i in range(0, 41)]
    ys = [10 * math.exp(-((x - 5.0) ** 2) / 2) + 8 * math.exp(-((x - 15.0) ** 2) / 2)
          for x in xs]
    set_worksheet_data(b, sh, json.dumps([xs, ys]))
    return b, sh


def _ncols(book, sheet):
    from origin_pro_mcp.origin_connection import execute_labtalk, get_lt_var

    execute_labtalk(f'win -a {book}; page.active$ = "{sheet}";')
    return int(get_lt_var("wks.ncols"))


def test_interpolate_reuses_one_stable_book(live_origin):
    from origin_pro_mcp.origin_connection import workbook_names
    from origin_pro_mcp.tools.analysis import transform

    b, sh = _spectrum()
    transform(b, sh, 1, 2, method="interpolate", num_points=50)
    transform(b, sh, 1, 2, method="interpolate", num_points=30)
    transform(b, sh, 1, 2, method="interpolate", num_points=80)

    interp_books = [n for n in workbook_names() if n == "Interp" or n.startswith("Interp")]
    assert interp_books == ["Interp"], f"orphan Interp books: {interp_books}"


def test_find_peaks_no_column_growth_on_repeat(live_origin):
    from origin_pro_mcp.tools.analysis import transform

    b, sh = _spectrum()
    before = _ncols(b, sh)
    transform(b, sh, 1, 2, method="find_peaks", local_points=3)
    after_first = _ncols(b, sh)
    transform(b, sh, 1, 2, method="find_peaks", local_points=3)
    transform(b, sh, 1, 2, method="find_peaks", local_points=3)
    after_third = _ncols(b, sh)

    # First call adds exactly the two peak-output columns…
    assert after_first == before + 2, (before, after_first)
    # …and further calls reuse them (no further growth).
    assert after_third == after_first, (after_first, after_third)
