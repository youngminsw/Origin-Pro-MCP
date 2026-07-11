"""Live (Windows + Origin Pro COM) tests for the export round (items 6 & 10):
vector formats (pdf/eps/emf), honored width/height, removed dpi.

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_export.py -v

Safety: every test runs against its OWN isolated Origin instance spawned via
``DispatchEx("Origin.Application")``. Never touches ``Origin.ApplicationSI``.
``origin.Exit()`` always runs in the fixture teardown.
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


def _win_tmp(suffix: str) -> str:
    d = r"C:\Users\swym4\probe_out\roundb_tests"
    os.makedirs(d, exist_ok=True)
    fd, path = tempfile.mkstemp(suffix=suffix, dir=d)
    os.close(fd)
    os.remove(path)
    return path


def _build_graph(name="EXPG"):
    from origin_pro_mcp.tools.graph import create_graph
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("EXPB"))
    b, sh = made["name"], made["sheet"]
    xs = list(range(1, 21))
    set_worksheet_data(b, sh, json.dumps([xs, [v * v for v in xs]]))
    return json.loads(create_graph(name, b, sh, 1, 2, plot_type="line+symbol"))["name"]


@pytest.mark.parametrize("fmt,sniff", [
    ("pdf", b"%PDF"),
    ("eps", b"%!PS"),
    ("emf", None),
])
def test_vector_export_writes_valid_file(live_origin, fmt, sniff):
    """Item 6: pdf/eps/emf export to a real, non-trivial file with the right
    magic bytes (pdf/eps)."""
    from origin_pro_mcp.tools.graph import export_graph

    g = _build_graph()
    path = _win_tmp(f".{fmt}")
    out = export_graph(g, path, format=fmt)
    assert os.path.exists(path), out
    size = os.path.getsize(path)
    assert size > 500, f"{fmt} export suspiciously small: {size} bytes"
    if sniff is not None:
        with open(path, "rb") as fh:
            head = fh.read(8)
        assert head.startswith(sniff), f"{fmt} header={head!r}"


def test_svg_is_rejected(live_origin):
    """Item 6: SVG proved out unsupported — must raise, not silently no-op."""
    from origin_pro_mcp.tools.graph import export_graph

    g = _build_graph()
    with pytest.raises(ValueError):
        export_graph(g, _win_tmp(".svg"), format="svg")


def test_width_honored_without_sized_flag(live_origin):
    """Item 10: an explicit width applies to the raster output with no
    sized=True needed (no silent ignore). Height follows the aspect ratio,
    which expGraph controls, not us."""
    from PIL import Image

    from origin_pro_mcp.tools.graph import export_graph

    g = _build_graph()
    path = _win_tmp(".png")
    export_graph(g, path, format="png", width=1600)
    with Image.open(path) as im:
        w, _ = im.size
    assert w == 1600, f"width not honored: got {w}"


def test_export_tools_have_no_dead_params(live_origin):
    """Item 10: dpi (unsupported by expGraph) and height (silently aspect-
    locked) are gone from both export tools — nothing accepted-and-ignored."""
    import inspect

    from origin_pro_mcp.tools.graph import export_graph
    from origin_pro_mcp.tools.project import export_all_graphs

    for fn in (export_graph, export_all_graphs):
        params = inspect.signature(fn).parameters
        assert "dpi" not in params, f"{fn.__name__} still has dpi"
        assert "height" not in params, f"{fn.__name__} still has height"


def test_export_all_graphs_honors_width(live_origin):
    """Item 10: export_all_graphs width flows through to each exported file."""
    from PIL import Image

    from origin_pro_mcp.tools.project import export_all_graphs

    _build_graph("ALLG")
    out_dir = r"C:\Users\swym4\probe_out\roundb_tests\all"
    export_all_graphs(out_dir, format="png", width=900)
    pngs = [f for f in os.listdir(out_dir) if f.lower().endswith(".png")]
    assert pngs, "no graphs exported"
    with Image.open(os.path.join(out_dir, pngs[0])) as im:
        assert im.size[0] == 900, f"width not honored: {im.size}"
