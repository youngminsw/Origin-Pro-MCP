"""Live (Windows + Origin Pro COM) pixel-verified tests for the styling-report
fix round (2026-07-10 plan): the Task 0.5 settle barrier, set_plot_style
partial styling + error-bar knobs, axis frame width / per-side ticks, and
apply_publication_style integrity.

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_styling.py -v

Safety: every test runs against its OWN isolated Origin instance spawned via
``DispatchEx("Origin.Application")`` (same pattern as test_live_loaded_graph.py).
Never touches ``Origin.ApplicationSI``. ``origin.Exit()`` always runs in the
fixture teardown.
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


def _name(create_result: str) -> str:
    return json.loads(create_result)["name"]


def _red_pixel_count(path: str, threshold_r=200, threshold_gb=80) -> int:
    from PIL import Image

    im = Image.open(path).convert("RGB")
    px = im.load()
    w, h = im.size
    count = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            r, g, b = px[x, y]
            if r > threshold_r and g < threshold_gb and b < threshold_gb:
                count += 1
    return count


def _build_line_symbol_with_error(book="SMOKE", y_error=True):
    """A worksheet + one line+symbol series (X,Y,Yerr), for styling tests."""
    from origin_pro_mcp.tools.graph import create_graph
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet(book))
    b, sheet = made["name"], made["sheet"]
    set_worksheet_data(
        b, sheet,
        json.dumps([[1, 2, 3, 4, 5], [1, 4, 9, 16, 25], [0.5, 0.5, 1.0, 1.0, 1.5]]),
    )
    kwargs = {"y_error_col": 3} if y_error else {}
    g = _name(create_graph("LineG", b, sheet, 1, 2, plot_type="line+symbol", **kwargs))
    return g, b, sheet


def test_settle_barrier_immediate_color_set_takes_effect(tmp_path, live_origin):
    """Task 0.5 regression: setting a curve's color IMMEDIATELY after
    create_graph must actually render (no silent no-op from the new-page
    settle hazard)."""
    from origin_pro_mcp.tools.graph import export_graph_to_file
    from origin_pro_mcp.tools.style_helpers import get_plot_info, graph_layer_execute

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    pname = get_plot_info(g)[0]["name"]
    graph_layer_execute(g, f"set {pname} -c color(255,0,0);")
    out = str(tmp_path / "settle_regress.png")
    export_graph_to_file(g, out)
    assert _red_pixel_count(out) > 100
