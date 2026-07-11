"""Live (Windows + Origin Pro COM) test for batch/folder import (item 31b).

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_import.py -v

Safety: an isolated ``DispatchEx("Origin.Application")`` per test; never
``Origin.ApplicationSI``; ``origin.Exit()`` in teardown.
"""
import json
import os
import tempfile

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


def test_batch_import_directory_three_files(live_origin):
    from origin_pro_mcp.origin_connection import workbook_names
    from origin_pro_mcp.tools.worksheet import import_data

    folder = tempfile.mkdtemp(
        prefix="batch_", dir=r"C:\Users\swym4\probe_out\roundb_tests"
    )
    stems = ["alpha", "beta", "gamma"]
    for s in stems:
        with open(os.path.join(folder, f"{s}.csv"), "w") as fh:
            fh.write("X,Y\n1,10\n2,20\n3,30\n")

    out = json.loads(import_data(folder))
    assert out["batch"] is True
    assert out["matched"] == 3
    assert out["imported"] == 3
    assert all(r["ok"] for r in out["results"])

    # Each file landed in its own book, named from the stem.
    names = set(workbook_names())
    for s in stems:
        assert s in names, f"missing book {s}; have {names}"

    # And each imported book actually holds the data (2 columns x 3 rows).
    from origin_pro_mcp.origin_connection import sheet_names
    from origin_pro_mcp.tools.worksheet import get_worksheet_data

    sheet = sheet_names("alpha")[0]
    cols = json.loads(get_worksheet_data("alpha", sheet))["columns"]
    assert len(cols) == 2
    assert len(cols[0]) == 3
