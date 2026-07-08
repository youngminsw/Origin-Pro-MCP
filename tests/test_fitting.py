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
