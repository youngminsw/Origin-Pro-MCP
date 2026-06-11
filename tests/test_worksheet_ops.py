"""Guard tests for worksheet data operations (no Origin needed)."""
import pytest


def test_set_column_formula_blocks_injection(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_formula

    with pytest.raises(ValueError, match="cannot contain"):
        set_column_formula("Book1", "Sheet1", 2, "col(1); doc -s")


def test_set_column_formula_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_formula

    with pytest.raises(ValueError, match="not found"):
        set_column_formula("Ghost", "Sheet1", 2, "col(1)^2")


def test_set_column_formula_runs(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_formula

    fake_origin.lt_vars["wks.ncols"] = 5  # avoid the grow-loop
    msg = set_column_formula("Book1", "Sheet1", 2, "col(1)^2")
    assert "col(1)^2" in msg
    assert any("col(2) = col(1)^2" in s for s in fake_origin.executed)


def test_sort_worksheet_descending_flag(fake_origin):
    from origin_pro_mcp.tools.worksheet import sort_worksheet

    sort_worksheet("Book1", "Sheet1", 1, descending=True)
    assert any("wsort bycol:=1 descending:=1" in s for s in fake_origin.executed)


def test_sort_worksheet_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import sort_worksheet

    with pytest.raises(ValueError, match="not found"):
        sort_worksheet("Ghost", "Sheet1", 1)


def test_delete_columns_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import delete_columns

    with pytest.raises(ValueError, match="not found"):
        delete_columns("Ghost", "Sheet1", 2)


def test_set_column_properties_requires_one_field(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_properties

    with pytest.raises(ValueError, match="at least one"):
        set_column_properties("Book1", "Sheet1", 1)


def test_set_column_properties_bad_designation(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_properties

    with pytest.raises(ValueError, match="designation must be one of"):
        set_column_properties("Book1", "Sheet1", 1, designation="vector")


def test_set_column_properties_sets_units_singular(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_properties

    set_column_properties("Book1", "Sheet1", 1, units="s", long_name="Time")
    assert any("wks.col1.unit$" in s for s in fake_origin.executed)
    assert any("wks.col1.lname$" in s for s in fake_origin.executed)


def test_transpose_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import transpose_worksheet

    with pytest.raises(ValueError, match="not found"):
        transpose_worksheet("Ghost", "Sheet1")
