"""Guard tests for worksheet data operations (no Origin needed).

Column add/delete/properties/formula operations are driven through the
consolidated ``manage_columns`` dispatcher.
"""
import json

import pytest


def test_create_worksheet_returns_parseable_name(fake_origin):
    from origin_pro_mcp.tools.worksheet import create_worksheet

    out = json.loads(create_worksheet("NewBook"))
    assert out == {
        "name": "NewBook",
        "requested_name": "NewBook",
        "renamed": False,
        "sheet": "Sheet1",
        "added_to_existing_book": False,
    }


def test_create_worksheet_reports_rename(fake_origin):
    from origin_pro_mcp.tools.worksheet import create_worksheet

    # Simulate Origin uniquifying the name on collision.
    fake_origin.CreatePage = lambda kind, name, template: "NewBook2"
    out = json.loads(create_worksheet("NewBook"))
    assert out["name"] == "NewBook2"
    assert out["requested_name"] == "NewBook"
    assert out["renamed"] is True


def test_create_worksheet_adds_sheet_to_existing_book(fake_origin):
    from origin_pro_mcp.tools.worksheet import create_worksheet

    # fake_origin.books already has "Book1" with "Sheet1".
    out = json.loads(create_worksheet("Book1", "Sheet2"))
    assert out == {
        "name": "Book1",
        "requested_name": "Book1",
        "renamed": False,
        "sheet": "Sheet2",
        "added_to_existing_book": True,
    }
    assert any(
        s.startswith('newsheet name:="Sheet2"') for s in fake_origin.executed
    )


def test_create_worksheet_existing_sheet_raises(fake_origin):
    from origin_pro_mcp.tools.worksheet import create_worksheet

    with pytest.raises(ValueError, match="already has a sheet"):
        create_worksheet("Book1", "Sheet1")


def test_set_column_formula_blocks_injection(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="cannot contain"):
        manage_columns("Book1", "Sheet1", op="formula", col=2, formula="col(1); doc -s")


def test_set_column_formula_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="not found"):
        manage_columns("Ghost", "Sheet1", op="formula", col=2, formula="col(1)^2")


def test_set_column_formula_runs(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    fake_origin.lt_vars["wks.ncols"] = 5  # avoid the grow-loop
    msg = manage_columns("Book1", "Sheet1", op="formula", col=2, formula="col(1)^2")
    assert "col(1)^2" in msg
    assert any("col(2) = col(1)^2" in s for s in fake_origin.executed)


def test_formula_op_requires_formula(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="requires col and formula"):
        manage_columns("Book1", "Sheet1", op="formula", col=2)


def test_manage_columns_bad_op(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="op must be one of"):
        manage_columns("Book1", "Sheet1", op="rename")


def test_add_columns_runs(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    msg = manage_columns("Book1", "Sheet1", op="add", count=2)
    assert "2 column" in msg
    assert sum(1 for s in fake_origin.executed if s == "wks.addCol();") == 2


def test_sort_worksheet_descending_flag(fake_origin):
    from origin_pro_mcp.tools.worksheet import sort_worksheet

    sort_worksheet("Book1", "Sheet1", 1, descending=True)
    assert any("wsort bycol:=1 descending:=1" in s for s in fake_origin.executed)


def test_sort_worksheet_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import sort_worksheet

    with pytest.raises(ValueError, match="not found"):
        sort_worksheet("Ghost", "Sheet1", 1)


def test_delete_columns_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="not found"):
        manage_columns("Ghost", "Sheet1", op="delete", col=2)


def test_delete_op_requires_col(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="op 'delete' requires col"):
        manage_columns("Book1", "Sheet1", op="delete")


def test_set_column_properties_requires_one_field(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="at least one"):
        manage_columns("Book1", "Sheet1", op="properties", col=1)


def test_set_column_properties_bad_designation(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="designation must be one of"):
        manage_columns("Book1", "Sheet1", op="properties", col=1, designation="vector")


def test_set_column_properties_sets_units_singular(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    manage_columns("Book1", "Sheet1", op="properties", col=1, units="s", long_name="Time")
    assert any("wks.col1.unit$" in s for s in fake_origin.executed)
    assert any("wks.col1.lname$" in s for s in fake_origin.executed)


def test_transpose_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import transpose_worksheet

    with pytest.raises(ValueError, match="not found"):
        transpose_worksheet("Ghost", "Sheet1")


def test_get_worksheet_data_maps_missing_to_null(fake_origin):
    from origin_pro_mcp.tools.worksheet import get_worksheet_data

    # Simulate a ragged column: Origin pads the shorter column with its
    # missing-value sentinel (~1.2e308).
    fake_origin.worksheet_data = ((1.0, 4.0), (2.0, 1.23e308))
    out = json.loads(get_worksheet_data("Book1", "Sheet1"))
    assert out["columns"] == [[1.0, 2.0], [4.0, None]]


def test_set_worksheet_data_accepts_null_as_missing_value(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    set_worksheet_data("Book1", "Sheet1", "[[1,2,null,4]]")
    # Exactly the null position (0-based row 2 -> LabTalk 1-based row 3)
    # gets the missing-value write; no other row is touched this way.
    assert "col(1)[3]=0/0;" in fake_origin.executed
    assert not any(
        s.startswith("col(1)[1]=") or s.startswith("col(1)[2]=") or s.startswith("col(1)[4]=")
        for s in fake_origin.executed
    )


def test_set_worksheet_data_accepts_nan_as_missing_value(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    # A bare NaN token in the JSON payload (Python's json module parses it
    # to float('nan')) is treated the same as JSON null.
    set_worksheet_data("Book1", "Sheet1", "[[1,NaN,3]]")
    assert "col(1)[2]=0/0;" in fake_origin.executed


def test_set_worksheet_data_no_missing_cells_emits_no_labtalk(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    set_worksheet_data("Book1", "Sheet1", "[[1,2,3]]")
    assert not any("=0/0;" in s for s in fake_origin.executed)


def test_set_worksheet_data_raises_naming_cell_when_missing_write_fails(fake_origin):
    """Item 2: if the per-cell missing-value write fails, the cell still holds
    the bulk 0.0 placeholder — the tool must raise naming that exact cell, not
    return success over corrupt data."""
    import pytest

    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    # Two columns; the null is at col 2 row 2 -> LabTalk col(2)[2]=0/0.
    fake_origin.execute_results["col(2)[2]=0/0"] = False
    with pytest.raises(ValueError, match="col 2 row 2"):
        set_worksheet_data("Book1", "Sheet1", "[[1,2],[3,null]]")
