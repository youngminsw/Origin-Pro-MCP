import json

import pytest

from origin_pro_mcp.origin_connection import get_origin, get_lt_var


def test_line_fit_parameter_keys_match_regardless_of_plot_on_graph(fake_origin):
    """function="line" must return the same parameter keys whether it takes
    the fast fitlr path (no graph) or the NLFit path (plot_on_graph set) —
    both must report "intercept"/"slope", never "A"/"B"."""
    from origin_pro_mcp.tools.fitting import curve_fit, list_fitting_functions

    fake_origin.lt_vars["fitlr.r"] = 0.99
    fake_origin.lt_vars["fitlr.a"] = 1.0
    fake_origin.lt_vars["fitlr.b"] = 2.0
    no_graph = json.loads(curve_fit("Book1", "Sheet1", 1, 2, function="line"))

    fake_origin.lt_vars["__mcpread"] = 3.0
    with_graph = json.loads(
        curve_fit("Book1", "Sheet1", 1, 2, function="line", plot_on_graph="Graph1")
    )

    assert set(no_graph["parameters"].keys()) == {"intercept", "slope"}
    assert set(with_graph["parameters"].keys()) == {"intercept", "slope"}

    functions = json.loads(list_fitting_functions())
    assert functions["linear"]["line"] == ["intercept", "slope"]
    assert set(functions["linear"]["line"]) == set(no_graph["parameters"].keys())
    assert set(functions["linear"]["line"]) == set(with_graph["parameters"].keys())


def test_curve_fit_x_range_emits_row_subrange(fake_origin):
    """x_min/x_max resolve to a 1-based row block and append `[i1:i2]` to the
    NLFit input range so the fit is restricted to those rows."""
    from origin_pro_mcp.tools.fitting import curve_fit

    # X in column 1 = 0,1,2,...,10 (rows 1..11). x in [2,8] => rows 3..9.
    fake_origin.worksheet_data = tuple((float(i), float(i * i)) for i in range(0, 11))
    fake_origin.lt_vars["__mcpread"] = 5.0
    curve_fit("Book1", "Sheet1", 1, 2, function="gauss", x_min=2.0, x_max=8.0)
    nlbegin = [c for c in fake_origin.executed if c.startswith("nlbegin")]
    assert nlbegin, fake_origin.executed
    assert "[3:9]" in nlbegin[0], nlbegin[0]


def test_curve_fit_x_range_line_uses_nlfit_not_fitlr(fake_origin):
    """A range-restricted line fit must take the NLFit path (which accepts the
    row subrange), not the fast fitlr path."""
    from origin_pro_mcp.tools.fitting import curve_fit

    fake_origin.worksheet_data = tuple((float(i), float(2 * i)) for i in range(0, 11))
    fake_origin.lt_vars["__mcpread"] = 1.0
    curve_fit("Book1", "Sheet1", 1, 2, function="line", x_min=2.0, x_max=8.0)
    joined = " ".join(fake_origin.executed)
    assert "nlbegin" in joined
    assert "fitlr" not in joined


def test_curve_fit_bad_x_range_raises(fake_origin):
    from origin_pro_mcp.tools.fitting import curve_fit

    with pytest.raises(ValueError):
        curve_fit("Book1", "Sheet1", 1, 2, function="gauss", x_min=8.0, x_max=2.0)


def test_curve_fit_x_range_no_rows_raises(fake_origin):
    from origin_pro_mcp.tools.fitting import curve_fit

    fake_origin.worksheet_data = tuple((float(i), float(i)) for i in range(0, 11))
    with pytest.raises(ValueError):
        curve_fit("Book1", "Sheet1", 1, 2, function="gauss", x_min=100.0, x_max=200.0)


@pytest.mark.requires_origin
def test_linear_fit():
    o = get_origin()
    o.Execute("doc -s; doc -n;")
    o.CreatePage(2, "FitData", "origin")
    o.PutWorksheet("[FitData]Sheet1", [1.0, 2.0, 3.0, 4.0, 5.0], 0, 0)
    o.PutWorksheet("[FitData]Sheet1", [2.1, 3.9, 6.1, 7.9, 10.1], 0, 1)

    # Set column designations: col(1)=X, col(2)=Y
    o.Execute("win -a FitData;")
    o.Execute("[FitData]Sheet1!col(1).type = 4;")
    o.Execute("[FitData]Sheet1!col(2).type = 1;")

    # Use fitlr for linear regression; fitlr.r is the correlation coefficient
    o.Execute("fitlr col(2);")
    r = get_lt_var("fitlr.r")
    r2 = r * r
    assert r2 > 0.99
