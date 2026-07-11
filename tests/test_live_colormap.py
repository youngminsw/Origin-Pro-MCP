"""Live (Windows + Origin Pro COM) tests for the colormap read-back fixes
(review items 23/24): reject an unknown palette name up front, and read the
Z-range back so a non-colormap graph errors instead of a false success.

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_colormap.py -v

Safety: isolated ``DispatchEx("Origin.Application")``; never ApplicationSI;
``origin.Exit()`` in teardown.
"""
import json

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


def _heatmap(name="CMAP"):
    from origin_pro_mcp.tools.matrix import (
        create_matrix,
        create_matrix_plot,
        set_matrix_data,
    )

    m = json.loads(create_matrix(name, rows=6, cols=6))["name"]
    grid = [[i * 6 + j for j in range(6)] for i in range(6)]
    set_matrix_data(m, json.dumps(grid))
    return json.loads(create_matrix_plot(m, plot_type="heatmap"))["name"]


def test_colormap_rejects_bogus_palette_live(live_origin):
    from origin_pro_mcp.tools.graph import colormap

    g = _heatmap("CMAPBAD")
    with pytest.raises(ValueError, match="Unknown palette"):
        colormap(g, palette="NoSuchPaletteZZZ")


def test_colormap_valid_palette_applies_live(live_origin):
    from origin_pro_mcp.tools.graph import colormap

    g = _heatmap("CMAPOK")
    msg = colormap(g, palette="Fire")  # a real Origin built-in
    assert "Fire" in msg


def test_colormap_zrange_reads_back_live(live_origin):
    from origin_pro_mcp.tools.graph import colormap

    g = _heatmap("CMAPZ")
    msg = colormap(g, z_min=2.0, z_max=20.0)
    assert "Z range" in msg


def test_colormap_zrange_errors_on_non_colormap_graph_live(live_origin):
    """A plain line graph has no color scale; setting a Z range must not report
    a false success — the read-back gate should reject it."""
    from origin_pro_mcp.tools.graph import colormap, create_graph
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("CMLB"))
    book, sheet = made["name"], made["sheet"]
    set_worksheet_data(book, sheet, json.dumps([[1, 2, 3, 4], [2, 4, 6, 8]]))
    g = json.loads(create_graph("LineNoCmap", book, sheet, 1, 2, plot_type="line"))["name"]
    with pytest.raises(ValueError, match="did not take"):
        colormap(g, z_min=2.0, z_max=20.0)
