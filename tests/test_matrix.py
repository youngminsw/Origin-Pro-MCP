"""Guard/dispatch tests for the matrix tools (no Origin needed)."""
import json

import pytest

from conftest import FakeMatrix


def test_set_matrix_data_rejects_bad_json(fake_origin):
    from origin_pro_mcp.tools.matrix import set_matrix_data

    fake_origin.matrices = [FakeMatrix("Mtx")]
    with pytest.raises(ValueError, match="JSON array of arrays"):
        set_matrix_data("Mtx", "not json")


def test_set_matrix_data_rejects_ragged_rows(fake_origin):
    from origin_pro_mcp.tools.matrix import set_matrix_data

    fake_origin.matrices = [FakeMatrix("Mtx")]
    with pytest.raises(ValueError, match="same length"):
        set_matrix_data("Mtx", "[[1,2,3],[4,5]]")


def test_set_matrix_data_rejects_non_numeric(fake_origin):
    from origin_pro_mcp.tools.matrix import set_matrix_data

    fake_origin.matrices = [FakeMatrix("Mtx")]
    with pytest.raises(ValueError, match="non-numeric"):
        set_matrix_data("Mtx", '[["a","b"]]')


def test_set_matrix_data_unknown_matrix_lists_open(fake_origin):
    from origin_pro_mcp.tools.matrix import set_matrix_data

    with pytest.raises(ValueError, match="not found"):
        set_matrix_data("Ghost", "[[1,2],[3,4]]")


def test_set_and_get_matrix_roundtrip(fake_origin):
    from origin_pro_mcp.tools.matrix import set_matrix_data, get_matrix_data

    fake_origin.matrices = [FakeMatrix("Mtx")]
    msg = set_matrix_data("Mtx", "[[1,2,3],[4,5,6]]")
    assert "2x3" in msg
    out = json.loads(get_matrix_data("Mtx"))
    assert out["rows"] == [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]


def test_get_matrix_data_maps_missing_to_null(fake_origin):
    from origin_pro_mcp.tools.matrix import get_matrix_data

    m = FakeMatrix("Mtx")
    fake_origin.matrices = [m]
    fake_origin.matrix_data["[Mtx]MSheet1"] = [[1.0, 1e308], [2.0, 3.0]]
    out = json.loads(get_matrix_data("Mtx"))
    assert out["rows"] == [[1.0, None], [2.0, 3.0]]


def test_worksheet_to_matrix_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.matrix import worksheet_to_matrix

    with pytest.raises(ValueError, match="not found"):
        worksheet_to_matrix("Ghost", "Sheet1", 1, 2, 3)


def test_create_matrix_plot_unknown_matrix(fake_origin):
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    with pytest.raises(ValueError, match="not found"):
        create_matrix_plot("Ghost", plot_type="heatmap")


def test_create_matrix_plot_rejects_bad_type(fake_origin):
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    fake_origin.matrices = [FakeMatrix("Mtx")]
    with pytest.raises(ValueError, match="plot_type must be one of"):
        create_matrix_plot("Mtx", plot_type="wireframe")


def test_create_matrix_plot_z_label_sets_matrix_longname(fake_origin):
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    fake_origin.matrices = [FakeMatrix("Mtx")]
    create_matrix_plot("Mtx", plot_type="heatmap", z_label="Intensity (a.u.)")
    assert any(
        'wks.col1.lname$ = "Intensity (a.u.)"' in s for s in fake_origin.executed
    )


def test_create_matrix_plot_uses_colorscale_templates(fake_origin):
    """Each matrix plot type must plot from the system template that carries a
    data-linked color scale (regression guard for the missing-colorbar fix)."""
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    cases = {"contour": "CONTOUR", "heatmap": "HeatMap",
             "image": "image", "surface": "glcmap"}
    for plot_type, tmpl in cases.items():
        fake_origin.matrices = [FakeMatrix("Mtx")]
        fake_origin.executed = []
        create_matrix_plot("Mtx", plot_type=plot_type)
        assert any(
            f"template:={tmpl}" in s and "plotm" in s
            for s in fake_origin.executed
        ), f"{plot_type} did not plot via template {tmpl}"


def test_create_matrix_plot_surface_omits_plot_id(fake_origin):
    """Surface is a Z-colored OpenGL mesh from glcmap and must NOT pass a 2D
    plot id; the 2D types must pass theirs."""
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    fake_origin.matrices = [FakeMatrix("Mtx")]
    fake_origin.executed = []
    create_matrix_plot("Mtx", plot_type="surface")
    plotm = next(s for s in fake_origin.executed if "plotm" in s)
    assert "plot:=" not in plotm

    fake_origin.matrices = [FakeMatrix("Mtx")]
    fake_origin.executed = []
    create_matrix_plot("Mtx", plot_type="contour")
    plotm = next(s for s in fake_origin.executed if "plotm" in s)
    assert "plot:=226" in plotm
