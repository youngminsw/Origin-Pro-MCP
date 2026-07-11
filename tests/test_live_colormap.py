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


def _distinct_cell_colors(path: str) -> int:
    """Count distinct quantized colors down vertical strips through the heatmap
    CELLS (left ~60%, avoiding the colorbar on the right). A default banded
    heatmap yields ~8; a continuous one yields many more. This is the pixel
    ground-truth that settles the #17b "level-count is a no-op" report: the tool
    changes real pixels, not just a read-back."""
    from PIL import Image

    im = Image.open(path).convert("RGB")
    w, h = im.size
    px = im.load()
    best = 0
    for xf in (0.30, 0.38, 0.46, 0.54):
        x = int(w * xf)
        seen = set()
        for y in range(int(h * 0.15), int(h * 0.85)):
            r, g, b = px[x, y]
            if (r > 245 and g > 245 and b > 245) or (r < 12 and g < 12 and b < 12):
                continue  # skip frame / background / overflow blocks
            seen.add((r // 8, g // 8, b // 8))
        best = max(best, len(seen))
    return best


def test_colormap_levels_makes_heatmap_continuous_live(live_origin, tmp_path):
    """Regression guard for #17b: colormap(levels=) turns a banded heatmap into a
    continuous map on a real graph — and the heatmap KEEPS its colorbar (unlike
    an image plot). Pixel band-count, not just the numColors read-back."""
    from origin_pro_mcp.tools.graph import colormap, export_graph_to_file
    from origin_pro_mcp.tools.matrix import (
        create_matrix,
        create_matrix_plot,
        set_matrix_data,
    )

    # A tall smooth gradient (0..15 down 40 rows) so a vertical cell strip
    # crosses many cells — a 6x6 grid is too coarse to reveal continuity.
    rows, cols = 40, 24
    m = json.loads(create_matrix("CMLVL", rows=rows, cols=cols))["name"]
    grid = [[round(15.0 * (i / (rows - 1)), 4) for _ in range(cols)] for i in range(rows)]
    set_matrix_data(m, json.dumps(grid))
    g = json.loads(create_matrix_plot(m, plot_type="heatmap"))["name"]
    colormap(g, palette="Viridis", z_min=0, z_max=15)

    banded = str(tmp_path / "cmlvl_banded.png")
    export_graph_to_file(g, banded)
    banded_colors = _distinct_cell_colors(banded)

    msg = colormap(g, levels=48)
    assert "48" in msg  # read-back-verified numColors in the tool

    smooth = str(tmp_path / "cmlvl_smooth.png")
    export_graph_to_file(g, smooth)
    smooth_colors = _distinct_cell_colors(smooth)

    # Banded map is coarse (~8 bands); continuous map is much richer. Require a
    # clear jump so a future silent no-op (the reported failure) fails this test.
    assert banded_colors <= 16, f"expected banded (~8), got {banded_colors}"
    assert smooth_colors >= 24, f"expected continuous (>=24), got {smooth_colors}"
    assert smooth_colors >= banded_colors * 2


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
