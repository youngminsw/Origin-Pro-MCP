"""N7: per-plot RGB color is set through set_plot_style(rgb="r,g,b"), targeting
the exact plot with LabTalk `layer -s <index>; set %C ...` in ONE script (a COM
Activate is invisible to a separate execute_labtalk's %C); plus ungroup_plots,
which runs `layer -g` on Layer1's GraphLayer object directly."""
import pytest

from fakes import FakeGraph
from origin_pro_mcp import origin_connection


def _graph3(fake):
    fake.graphs = [FakeGraph("G", plot_names=["G_A", "G_B", "G_C"])]


def test_set_plot_style_rgb_targets_plot_via_percentC(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_style
    _graph3(fake_origin)
    set_plot_style("G", plot_index=2, rgb="128,0,200")
    # styling runs on the Layer1 COM object (gl.Execute), not global execute:
    # `layer -s 2; set %C -c color(...)` in ONE script.
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert any("layer -s 2; set %C -c color(128,0,200)" in s
               for s in layer.executed)


def test_set_plot_style_rgb_overrides_named_color(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_style
    _graph3(fake_origin)
    set_plot_style("G", plot_index=1, color="red", rgb="10,20,30")
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert any("set %C -c color(10,20,30)" in s for s in layer.executed)
    assert not any("-c red" in s or "-c color(255" in s for s in layer.executed)


@pytest.mark.parametrize("bad", ["300,0,0", "0,-1,0", "1,2", "a,b,c"])
def test_set_plot_style_rejects_bad_rgb(fake_origin, bad):
    from origin_pro_mcp.tools.style import set_plot_style
    _graph3(fake_origin)
    with pytest.raises(ValueError):
        set_plot_style("G", plot_index=1, rgb=bad)


def test_ungroup_plots(fake_origin):
    from origin_pro_mcp.tools.style import ungroup_plots
    _graph3(fake_origin)
    out = ungroup_plots("G")
    assert "Ungrouped" in out
    # `layer -g` runs on Layer1's GraphLayer object, not the global context.
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert any("layer -g" in s for s in layer.executed)
