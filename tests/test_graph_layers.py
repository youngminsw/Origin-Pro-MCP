"""Guard tests for layer/axis tools (no Origin needed).

These drive the consolidated ``axis``, ``annotate``, and ``colormap``
dispatchers.
"""
import pytest


def test_set_axis_scale_bad_scale(fake_origin):
    from origin_pro_mcp.tools.graph import axis

    with pytest.raises(ValueError, match="scale must be one of"):
        axis("Graph1", op="scale", axis="y", scale="logarithmic")


def test_set_axis_scale_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.graph import axis

    with pytest.raises(ValueError, match="not found"):
        axis("Ghost", op="scale", axis="y", scale="log10")


def test_set_axis_scale_runs(fake_origin):
    from origin_pro_mcp.tools.graph import axis

    msg = axis("Graph1", op="scale", axis="y", scale="log10")
    assert "log10" in msg
    assert any("layer.y.type = 2" in s for s in fake_origin.executed)


def test_add_second_y_axis_rejects_non_xy(fake_origin):
    from origin_pro_mcp.tools.graph import add_second_y_axis

    with pytest.raises(ValueError, match="only X,Y"):
        add_second_y_axis("Graph1", "Book1", "Sheet1", 1, 2, plot_type="histogram")


def test_add_layer_bad_type(fake_origin):
    from origin_pro_mcp.tools.graph import add_layer

    with pytest.raises(ValueError, match="layer_type must be one of"):
        add_layer("Graph1", "diagonal")


def test_add_reference_line_bad_orientation(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    with pytest.raises(ValueError, match="orientation must be one of"):
        annotate("Graph1", kind="reference_line", orientation="diagonal", value=5)


def test_add_reference_line_runs(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    annotate("Graph1", kind="reference_line", orientation="horizontal", value=5)
    assert any("draw -l -h 5.0" in s for s in fake_origin.executed)


def test_add_text_annotation_blocks_injection(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    with pytest.raises(ValueError, match="cannot be empty or contain"):
        annotate("Graph1", kind="text", text="hi; doc -s", x1=1, y1=2)


def test_add_text_annotation_runs(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    msg = annotate("Graph1", kind="text", text="Peak", x1=3.0, y1=70.0)
    assert "Peak" in msg
    assert any("label -p 3.0 70.0 -n anno Peak" in s for s in fake_origin.executed)


def test_apply_color_map_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.graph import colormap

    with pytest.raises(ValueError, match="not found"):
        colormap("Ghost", palette="Fire")


def test_apply_color_map_runs(fake_origin):
    from origin_pro_mcp.tools.graph import colormap

    colormap("Graph1", palette="Fire")
    assert any("layer.cmap.load(Fire.pal)" in s for s in fake_origin.executed)

def test_apply_color_map_bundled_viridis_full_path(fake_origin):
    """Bundled perceptually-uniform maps (viridis/cividis/...) are not Origin
    2020 built-ins, so they must be loaded from the bundled .pal by full,
    quoted path (regression for the colorblind-safe palette feature)."""
    from origin_pro_mcp.tools.graph import colormap

    colormap("Graph1", palette="Viridis")
    loads = [s for s in fake_origin.executed if "layer.cmap.load(" in s]
    assert loads, "no cmap load issued"
    assert any('opm_Viridis.pal"' in s and "load(\"" in s for s in loads), loads


def test_bundled_palettes_present():
    """The viridis-class .pal files must ship inside the package."""
    import os
    from origin_pro_mcp.tools.graph import _BUNDLED_PAL_DIR

    have = {f.lower() for f in os.listdir(_BUNDLED_PAL_DIR)}
    for name in ("viridis.pal", "cividis.pal"):
        assert name in have, (name, have)


def test_set_colormap_levels_bad_range(fake_origin):
    from origin_pro_mcp.tools.graph import colormap

    with pytest.raises(ValueError, match="z_max must be greater"):
        colormap("Graph1", z_min=5, z_max=5)


def test_colormap_requires_an_argument(fake_origin):
    from origin_pro_mcp.tools.graph import colormap

    with pytest.raises(ValueError, match="requires palette"):
        colormap("Graph1")


def test_colormap_requires_both_z_bounds(fake_origin):
    from origin_pro_mcp.tools.graph import colormap

    with pytest.raises(ValueError, match="both z_min and z_max"):
        colormap("Graph1", z_min=1)


def test_add_line_runs(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    annotate("Graph1", kind="line", x1=1, y1=2, x2=3, y2=4)
    assert any("draw -l {1.0,2.0,3.0,4.0}" in s for s in fake_origin.executed)


def test_create_graph_box_designates_y(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    create_graph("G", "Book1", "Sheet1", 1, 2, plot_type="box")
    assert any("wks.col2.type = 1" in s for s in fake_origin.executed)
    assert any("plot:=206" in s for s in fake_origin.executed)


def test_add_arrow_sets_arrowhead(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    msg = annotate("Graph1", kind="arrow", x1=1, y1=2, x2=3, y2=4)
    assert "single-headed" in msg
    assert any(".arrowEndShape = 1" in s for s in fake_origin.executed)
    assert not any("arrowBeginShape" in s for s in fake_origin.executed)


def test_add_arrow_double_headed(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    annotate("Graph1", kind="arrow", x1=1, y1=2, x2=3, y2=4, double_headed=True)
    assert any("arrowBeginShape = 1" in s for s in fake_origin.executed)


def test_add_arrow_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.graph import annotate

    with pytest.raises(ValueError, match="not found"):
        annotate("Ghost", kind="arrow", x1=1, y1=2, x2=3, y2=4)
