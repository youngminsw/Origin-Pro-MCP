"""Guard tests for worksheet data operations (no Origin needed).

Column add/delete/properties/formula operations are driven through the
consolidated ``manage_columns`` dispatcher.
"""
import pytest


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
