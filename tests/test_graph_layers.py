"""Guard tests for layer/axis tools (no Origin needed)."""
import pytest


def test_set_axis_scale_bad_scale(fake_origin):
    from origin_pro_mcp.tools.graph import set_axis_scale

    with pytest.raises(ValueError, match="scale must be one of"):
        set_axis_scale("Graph1", "y", "logarithmic")


def test_set_axis_scale_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.graph import set_axis_scale

    with pytest.raises(ValueError, match="not found"):
        set_axis_scale("Ghost", "y", "log10")


def test_set_axis_scale_runs(fake_origin):
    from origin_pro_mcp.tools.graph import set_axis_scale

    msg = set_axis_scale("Graph1", "y", "log10")
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
    from origin_pro_mcp.tools.graph import add_reference_line

    with pytest.raises(ValueError, match="orientation must be one of"):
        add_reference_line("Graph1", "diagonal", 5)


def test_add_reference_line_runs(fake_origin):
    from origin_pro_mcp.tools.graph import add_reference_line

    add_reference_line("Graph1", "horizontal", 5)
    assert any("draw -l -h 5.0" in s for s in fake_origin.executed)


def test_add_text_annotation_blocks_injection(fake_origin):
    from origin_pro_mcp.tools.graph import add_text_annotation

    with pytest.raises(ValueError, match="cannot be empty or contain"):
        add_text_annotation("Graph1", "hi; doc -s", 1, 2)


def test_add_text_annotation_runs(fake_origin):
    from origin_pro_mcp.tools.graph import add_text_annotation

    msg = add_text_annotation("Graph1", "Peak", 3.0, 70.0)
    assert "Peak" in msg
    assert any("label -p 3.0 70.0 -n anno Peak" in s for s in fake_origin.executed)


def test_apply_color_map_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.graph import apply_color_map

    with pytest.raises(ValueError, match="not found"):
        apply_color_map("Ghost", "Fire")


def test_apply_color_map_runs(fake_origin):
    from origin_pro_mcp.tools.graph import apply_color_map

    apply_color_map("Graph1", "Fire")
    assert any("layer.cmap.load(Fire.pal)" in s for s in fake_origin.executed)


def test_set_colormap_levels_bad_range(fake_origin):
    from origin_pro_mcp.tools.graph import set_colormap_levels

    with pytest.raises(ValueError, match="z_max must be greater"):
        set_colormap_levels("Graph1", 5, 5)


def test_add_line_runs(fake_origin):
    from origin_pro_mcp.tools.graph import add_line

    add_line("Graph1", 1, 2, 3, 4)
    assert any("draw -l {1.0,2.0,3.0,4.0}" in s for s in fake_origin.executed)


def test_create_graph_box_designates_y(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph

    create_graph("G", "Book1", "Sheet1", 1, 2, plot_type="box")
    assert any("wks.col2.type = 1" in s for s in fake_origin.executed)
    assert any("plot:=206" in s for s in fake_origin.executed)
