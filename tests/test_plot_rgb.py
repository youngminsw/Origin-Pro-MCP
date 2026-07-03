"""N7: typed per-plot RGB color / line width, targeting the exact plot via COM
DataPlots Activate + %C (no layer -s N scramble), plus ungroup_plots."""
import pytest

from fakes import FakeGraph
from origin_pro_mcp import origin_connection


def _graph3(fake):
    fake.graphs = [FakeGraph("G", plot_names=["G_A", "G_B", "G_C"])]


def test_set_plot_color_activates_right_plot_and_sets_rgb(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_color
    _graph3(fake_origin)
    out = set_plot_color("G", 2, 128, 0, 200)
    assert "G_B" in out and "RGB(128,0,200)" in out
    # the LabTalk color command targets the activated plot via %C
    assert any("set %C -c color(128,0,200)" in s for s in fake_origin.executed)
    # plot index 2 (0-based 1 = G_B) was the one activated
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert layer.DataPlots.Item(1).activated is True
    assert layer.DataPlots.Item(0).activated is False


@pytest.mark.parametrize("bad", [(300, 0, 0), (0, -1, 0), (0, 0, 256)])
def test_set_plot_color_rejects_out_of_range_rgb(fake_origin, bad):
    from origin_pro_mcp.tools.style import set_plot_color
    _graph3(fake_origin)
    with pytest.raises(ValueError, match="must be 0-255"):
        set_plot_color("G", 1, *bad)


def test_set_plot_color_rejects_bad_index(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_color
    _graph3(fake_origin)
    with pytest.raises(ValueError, match="out of range"):
        set_plot_color("G", 9, 10, 20, 30)


def test_set_plot_line_width_targets_plot(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_line_width
    _graph3(fake_origin)
    out = set_plot_line_width("G", 3, 2.5)
    assert "G_C" in out
    # 2.5 pt * 200 units/pt = 500
    assert any("set %C -w 500" in s for s in fake_origin.executed)
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert layer.DataPlots.Item(2).activated is True


def test_set_plot_line_width_rejects_nonpositive(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_line_width
    _graph3(fake_origin)
    with pytest.raises(ValueError, match="must be positive"):
        set_plot_line_width("G", 1, 0)


def test_ungroup_plots(fake_origin):
    from origin_pro_mcp.tools.style import ungroup_plots
    _graph3(fake_origin)
    out = ungroup_plots("G")
    assert "Ungrouped" in out
    assert any("layer -g" in s for s in fake_origin.executed)
