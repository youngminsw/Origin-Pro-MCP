"""Tests for the direct-control CLI dispatcher.

These run without Origin: they exercise tool discovery, argument
parsing/coercion, and exit codes. COM-backed tools are checked through
the shared fake Origin fixture.
"""
import json

from origin_pro_mcp import cli


def test_list_returns_all_tools(capsys):
    assert cli.main(["list"]) == 0
    out = capsys.readouterr().out
    # A representative spread of the 23 registered tools
    for name in ("apply_publication_style", "curve_fit", "export_graph", "run_labtalk"):
        assert name in out


def test_no_args_prints_list(capsys):
    assert cli.main([]) == 0
    assert "Available tools" in capsys.readouterr().out


def test_unknown_tool_exits_2(capsys):
    assert cli.main(["does_not_exist"]) == 2
    assert "Unknown tool" in capsys.readouterr().err


def test_unknown_arg_exits_2(capsys):
    assert cli.main(["apply_publication_style", "--bogus", "x"]) == 2
    assert "unknown argument --bogus" in capsys.readouterr().err


def test_list_fitting_functions_no_origin(capsys):
    # This tool does not touch COM, so it runs anywhere
    assert cli.main(["list_fitting_functions"]) == 0
    data = json.loads(capsys.readouterr().out)
    assert "peak" in data and "gauss" in data["peak"]


def test_coerce_handles_optional_float():
    import inspect
    from origin_pro_mcp.tools.style import apply_publication_style

    sig = inspect.signature(apply_publication_style)
    kwargs = cli._parse_kv(["--x_min", "280", "--graph_name", "FigA"], sig)
    assert kwargs == {"x_min": 280.0, "graph_name": "FigA"}
    assert isinstance(kwargs["x_min"], float)


def test_coerce_handles_int_and_bool():
    import inspect
    from origin_pro_mcp.tools.style import set_plot_style
    from origin_pro_mcp.tools.graph import axis

    kv = cli._parse_kv(["--plot_index", "2"], inspect.signature(set_plot_style))
    assert kv == {"plot_index": 2} and isinstance(kv["plot_index"], int)

    # axis op="tick" exposes show_minor as bool | None; --show_minor must coerce
    kv2 = cli._parse_kv(["--show_minor", "false"], inspect.signature(axis))
    assert kv2 == {"show_minor": False}


def test_consolidated_tool_coerces_typed_params(fake_origin):
    # The transform dispatcher uses explicit typed params (no **kwargs), so the
    # CLI must coerce each arg by annotation and dispatch without "unknown
    # argument". window_size/x_col/y_col are ints; method/smooth_method are str.
    argv = ["transform", "--data_book", "Book1", "--data_sheet", "Sheet1",
            "--x_col", "1", "--y_col", "2", "--method", "smooth",
            "--window_size", "5", "--smooth_method", "adjacent"]
    fake_origin.lt_vars["wks.ncols"] = 4
    assert cli.main(argv) == 0


def test_consolidated_tool_parse_kv_types():
    import inspect
    from origin_pro_mcp.tools.analysis import transform

    sig = inspect.signature(transform)
    kwargs = cli._parse_kv(
        ["--data_book", "B", "--data_sheet", "S", "--x_col", "1", "--y_col", "2",
         "--method", "smooth", "--window_size", "5", "--smooth_method", "binomial"],
        sig,
    )
    assert kwargs == {
        "data_book": "B", "data_sheet": "S", "x_col": 1, "y_col": 2,
        "method": "smooth", "window_size": 5, "smooth_method": "binomial",
    }
    assert all(isinstance(kwargs[k], int) for k in ("x_col", "y_col", "window_size"))
    assert isinstance(kwargs["method"], str)
    assert isinstance(kwargs["smooth_method"], str)


def test_json_args_dispatch(fake_origin, capsys):
    # --json path: list_worksheets takes no args but must accept {} and run
    assert cli.main(["list_worksheets", "--json", "{}"]) == 0
    data = json.loads(capsys.readouterr().out)
    assert data["graphs"] == ["Graph1"]


def test_tool_runtime_error_exits_1(fake_origin, capsys):
    # curve_fit with an unknown function raises ValueError -> exit 1
    rc = cli.main(["curve_fit", "--json",
                   '{"data_book":"Book1","data_sheet":"Sheet1","x_col":1,"y_col":2,"function":"quadratic"}'])
    assert rc == 1
    assert "ERROR" in capsys.readouterr().err


def test_bad_json_exits_2(capsys):
    assert cli.main(["list_worksheets", "--json", "{not json"]) == 2
    assert "Argument error" in capsys.readouterr().err
