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
