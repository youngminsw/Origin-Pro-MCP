"""Guard/dispatch tests for the matrix tools (no Origin needed)."""
import json

import pytest

from conftest import FakeMatrix


def test_create_matrix_returns_parseable_name(fake_origin):
    from origin_pro_mcp.tools.matrix import create_matrix

    # _create_matrix_book reads the assigned name back via LTStr("page.name$");
    # the default fake always returns "", so simulate Origin echoing the name.
    fake_origin.LTStr = lambda name: "NewMtx" if name == "page.name$" else ""
    out = json.loads(create_matrix("NewMtx", rows=5, cols=7))
    assert out == {
        "name": "NewMtx",
        "requested_name": "NewMtx",
        "renamed": False,
        "rows": 5,
        "cols": 7,
    }


def test_get_matrix_data_unknown_matrix_raises(fake_origin):
    from origin_pro_mcp.tools.matrix import get_matrix_data

    with pytest.raises(ValueError, match="not found"):
        get_matrix_data("Ghost")


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


def test_worksheet_to_matrix_returns_json_name(fake_origin):
    from origin_pro_mcp.tools.matrix import worksheet_to_matrix

    # _create_matrix_book reads the assigned name back via LTStr("page.name$").
    fake_origin.LTStr = lambda name: "MyGrid" if name == "page.name$" else ""
    out = json.loads(
        worksheet_to_matrix("Book1", "Sheet1", 1, 2, 3, matrix_book="MyGrid")
    )
    assert out["name"] == "MyGrid"
    assert out["requested_name"] == "MyGrid"
    assert out["renamed"] is False
    assert out["source"] == "[Book1]Sheet1"
    assert out["rows"] == 20
    assert out["cols"] == 20


def test_worksheet_to_matrix_no_name_requested_is_null(fake_origin):
    from origin_pro_mcp.tools.matrix import worksheet_to_matrix

    fake_origin.LTStr = lambda name: "Matrix" if name == "page.name$" else ""
    out = json.loads(worksheet_to_matrix("Book1", "Sheet1", 1, 2, 3))
    assert out["name"] == "Matrix"
    assert out["requested_name"] is None
    assert out["renamed"] is False


def test_create_matrix_plot_unknown_matrix(fake_origin):
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    with pytest.raises(ValueError, match="not found"):
        create_matrix_plot("Ghost", plot_type="heatmap")


def test_create_matrix_plot_returns_json_name(fake_origin):
    from origin_pro_mcp.tools.matrix import create_matrix_plot

    fake_origin.matrices = [FakeMatrix("Mtx")]
    # No new graph appears in fake.graphs (plotm is a no-op Execute call), so
    # the tool falls back to reading the active window name.
    fake_origin.LTStr = lambda name: "Graph2" if name == "page.name$" else ""
    out = json.loads(create_matrix_plot("Mtx", plot_type="heatmap"))
    assert out["name"] == "Graph2"
    assert out["requested_name"] is None
    assert out["renamed"] is False
    assert out["plot_type"] == "heatmap"


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
