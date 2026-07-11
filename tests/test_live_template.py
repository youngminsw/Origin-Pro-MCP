"""Live (Windows + Origin Pro COM) round-trip test for create_graph(template=)
(item 31a): style a graph, save it as a template, then build a NEW graph from
that template with different data and confirm the styling was reproduced.

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_template.py -v

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


def _win_png():
    d = r"C:\Users\swym4\probe_out\roundb_tests"
    os.makedirs(d, exist_ok=True)
    fd, path = tempfile.mkstemp(suffix=".png", dir=d)
    os.close(fd)
    os.remove(path)
    return path


def _red_pixels(path, tr=170, tgb=90):
    from PIL import Image

    im = Image.open(path).convert("RGB")
    px = im.load()
    w, h = im.size
    n = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            r, g, b = px[x, y]
            if r > tr and g < tgb and b < tgb:
                n += 1
    return n


def test_template_round_trip_reproduces_style(live_origin):
    from origin_pro_mcp.tools.graph import create_graph, export_graph
    from origin_pro_mcp.tools.project import save_graph_template
    from origin_pro_mcp.tools.style import set_plot_style
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("TPLB"))
    b, sh = made["name"], made["sheet"]
    set_worksheet_data(b, sh, json.dumps([[1, 2, 3, 4, 5], [1, 4, 9, 16, 25]]))

    # Style the source graph with a distinctive THICK RED line.
    g1 = json.loads(create_graph("TPL1", b, sh, 1, 2, plot_type="line"))["name"]
    set_plot_style(g1, plot_index=1, line_width=8, color="red")

    tmpl = os.path.join(r"C:\Users\swym4\probe_out\roundb_tests", "tpl_roundtrip.otpu")
    save_graph_template(g1, tmpl)

    # New data, built FROM the template — should come out thick+red.
    made2 = json.loads(create_worksheet("TPLB2"))
    b2, sh2 = made2["name"], made2["sheet"]
    set_worksheet_data(b2, sh2, json.dumps([[1, 2, 3, 4, 5], [5, 4, 3, 2, 1]]))

    res = json.loads(create_graph("TPL2", b2, sh2, 1, 2, plot_type="line", template=tmpl))
    assert res["template"], res
    g2 = res["name"]

    # Control: same data, NO template (Origin default = thin black line).
    g3 = json.loads(create_graph("TPL3", b2, sh2, 1, 2, plot_type="line"))["name"]

    p2, p3 = _win_png(), _win_png()
    export_graph(g2, p2)
    export_graph(g3, p3)
    red2 = _red_pixels(p2)
    red3 = _red_pixels(p3)

    # The template graph shows a substantial red line; the default shows ~none.
    assert red2 > 200, f"template graph not red/thick enough: {red2}"
    assert red2 > red3 * 5 + 50, f"template red {red2} not clearly > default {red3}"
