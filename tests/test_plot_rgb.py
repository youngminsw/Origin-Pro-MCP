"""N7: per-plot RGB color via set_plot_style(rgb="r,g,b"). Each plot is targeted
by its DATASET NAME (`set <name> ...`) run on the Layer1 COM object (gl.Execute),
verified on Origin 2020 to color each curve of an ungrouped multi-plot graph
independently; needs no active window (works on .opju-loaded graphs) and no
`layer -s` (which only ever selects plot 1). Plus ungroup_plots, which runs
`layer -g` on Layer1's GraphLayer object directly."""
import pytest

from fakes import FakeGraph
from origin_pro_mcp import origin_connection


def _graph3(fake):
    fake.graphs = [FakeGraph("G", plot_names=["G_A", "G_B", "G_C"])]


def test_set_plot_style_rgb_targets_plot_by_name(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_style
    _graph3(fake_origin)
    set_plot_style("G", plot_index=2, rgb="128,0,200")
    # plot 2 == dataset "G_B"; styling runs on the Layer1 COM object (gl.Execute).
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert any("set G_B -c color(128,0,200)" in s for s in layer.executed)


def test_set_plot_style_rgb_overrides_named_color(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_style
    _graph3(fake_origin)
    set_plot_style("G", plot_index=1, color="red", rgb="10,20,30")
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert any("set G_A -c color(10,20,30)" in s for s in layer.executed)
    assert not any("-c red" in s or "-c color(255" in s for s in layer.executed)


@pytest.mark.parametrize("bad", ["300,0,0", "0,-1,0", "1,2", "a,b,c"])
def test_set_plot_style_rejects_bad_rgb(fake_origin, bad):
    from origin_pro_mcp.tools.style import set_plot_style
    _graph3(fake_origin)
    with pytest.raises(ValueError):
        set_plot_style("G", plot_index=1, rgb=bad)


def test_ungroup_plots_rebuilds_independent_plots(fake_origin):
    from origin_pro_mcp.tools.style import ungroup_plots
    _graph3(fake_origin)
    out = ungroup_plots("G")
    assert "Ungrouped" in out and "rebuilt 3" in out
    # rebuild = remove every plot, then re-plot each dataset on its own (ungrouped),
    # all on the Layer1 COM object (gl.Execute).
    layer = origin_connection.get_origin()._graph_layers["[G]Layer1"]
    assert any("layer -e G_A" in s for s in layer.executed)
    assert any("plotxy iy:=G_A plot:=200 ogl:=[G]Layer1" in s for s in layer.executed)
    assert any("plotxy iy:=G_C plot:=200 ogl:=[G]Layer1" in s for s in layer.executed)


def test_ungroup_plots_raises_when_no_plot_rebuilds(fake_origin):
    """Item 25: ungroup removes every plot first; if NONE re-plot, the graph is
    left (near-)empty — that must raise, not read as success."""
    from origin_pro_mcp.tools.style import ungroup_plots
    _graph3(fake_origin)
    fake_origin.execute_results["plotxy"] = False  # every rebuild fails
    with pytest.raises(ValueError, match="Ungroup FAILED"):
        ungroup_plots("G")


def test_ungroup_plots_partial_rebuild_is_not_success_toned(fake_origin):
    """Item 25: when some series re-plot and others don't, say exactly which are
    now missing instead of a plain 'Ungrouped' success."""
    from origin_pro_mcp.tools.style import ungroup_plots
    _graph3(fake_origin)
    fake_origin.execute_results["plotxy iy:=G_B"] = False  # G_B fails to rebuild
    out = ungroup_plots("G")
    assert "PARTIAL" in out
    assert "G_B" in out
    assert "rebuilt 2 of 3" in out
