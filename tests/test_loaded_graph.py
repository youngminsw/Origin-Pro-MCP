"""S2: loaded-from-.opju graph edit-freeze fix (no Origin needed).

The pathology: a graph loaded from a project file reports zero DataPlots and
silently no-ops style/axis commands until its page is activated and a FRESH
layer handle is taken. These tests drive the fake's loaded/frozen models to
prove:
  * activation-then-retry reveals the plots (loaded graphs become editable),
  * a graph that stays at zero plots raises an actionable error (no silent
    success),
  * the read-back verification catches an axis mutation that didn't take,
  * the GraphPages fallback acquisition path works when FindGraphLayer is blind.
"""
import pytest

from fakes import FakeGraph
from origin_pro_mcp import origin_connection


# --- activation reveals a loaded graph's plots -------------------------------

def test_get_plot_names_activates_loaded_graph(fake_origin):
    from origin_pro_mcp.tools.style_helpers import get_plot_names

    g = FakeGraph("L", plot_names=["L_A", "L_B"], loaded=True)
    fake_origin.graphs = [g]
    assert g.visible_plot_names() == []          # hidden before activation
    assert get_plot_names("L") == ["L_A", "L_B"]  # acquire activates, then sees them
    assert g.activated is True


def test_set_plot_style_on_loaded_graph_succeeds(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("L", plot_names=["L_A", "L_B"], loaded=True)]
    msg = set_plot_style("L", plot_index=2, rgb="10,20,30")
    assert "Updated style for plot 2" in msg
    layer = origin_connection.get_origin()._graph_layers["[L]Layer1"]
    assert any("set L_B -c color(10,20,30)" in s for s in layer.executed)


# --- a graph that stays frozen must RAISE, never silently succeed -------------

def test_set_plot_style_on_frozen_graph_raises(fake_origin):
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("F", plot_names=["F_A", "F_B"], frozen=True)]
    with pytest.raises(ValueError, match="reports zero plots even after activating"):
        set_plot_style("F", plot_index=1, rgb="1,2,3")


def test_ungroup_on_frozen_graph_raises(fake_origin):
    from origin_pro_mcp.tools.style import ungroup_plots

    fake_origin.graphs = [FakeGraph("F", plot_names=["F_A", "F_B"], frozen=True)]
    with pytest.raises(ValueError, match="reports zero plots even after activating"):
        ungroup_plots("F")


def test_apply_publication_style_frozen_reports_no_plots(fake_origin):
    from origin_pro_mcp.tools.style import apply_publication_style

    fake_origin.graphs = [FakeGraph("F", plot_names=["F_A"], frozen=True)]
    msg = apply_publication_style("F")
    # Axes/frame/labels still apply, but the return must NOT imply curves were
    # styled — it must say no data plots were found.
    assert "NO data plots were found to style" in msg


# --- read-back verification catches a silent axis no-op ----------------------

def test_axis_range_readback_detects_frozen_noop(fake_origin):
    from origin_pro_mcp.tools.graph import axis

    fake_origin.graphs = [FakeGraph("F", plot_names=["F_A"], frozen=True)]
    # The frozen layer accepts `layer.y.from = ...` but ignores it; the read-back
    # sees the value never changed and raises instead of reporting success.
    with pytest.raises(ValueError, match="did not take effect"):
        axis("F", op="range", axis="y", range_min=-2, range_max=28)


def test_axis_range_readback_passes_on_healthy_graph(fake_origin):
    from origin_pro_mcp.tools.graph import axis

    fake_origin.graphs = [FakeGraph("G", plot_names=["G_A"])]
    msg = axis("G", op="range", axis="y", range_min=-2, range_max=28)
    assert "Set axis range" in msg


# --- fallback acquisition via the GraphPages collection ----------------------

def test_acquire_uses_graphpages_fallback_when_findgraphlayer_blind(fake_origin):
    from origin_pro_mcp.tools.style_helpers import get_plot_names

    fake_origin.graphs = [FakeGraph("G", plot_names=["G_A", "G_B"])]
    fake_origin.find_graph_layer_returns_none = True  # force the fallback path
    assert get_plot_names("G") == ["G_A", "G_B"]
