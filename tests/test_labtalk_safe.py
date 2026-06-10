import pytest

from origin_pro_mcp.labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_path,
    labtalk_string,
    positive_column,
    safe_labtalk_script,
)


def test_labtalk_name_rejects_statement_injection() -> None:
    with pytest.raises(ValueError, match="graph_name"):
        labtalk_name("Fig1;doc_s", "graph_name")


def test_labtalk_string_rejects_quote_breakout() -> None:
    with pytest.raises(ValueError, match="title"):
        labtalk_string('safe"; doc -s;', "title")


def test_labtalk_path_escapes_windows_backslashes() -> None:
    assert labtalk_path(r"C:\Users\me\figure.png", "file_path") == (
        r'"C:\\Users\\me\\figure.png"'
    )


def test_labtalk_choice_rejects_unknown_function() -> None:
    with pytest.raises(ValueError, match="function"):
        labtalk_choice("line;doc -s", {"line", "gauss"}, "function")


def test_positive_column_rejects_zero() -> None:
    with pytest.raises(ValueError, match="x_col"):
        positive_column(0, "x_col")


def test_safe_labtalk_script_allows_styling_commands() -> None:
    script = 'win -a Fig1; xb.text$ = "Temperature (K)"; layer.x.grid = 0;'

    assert safe_labtalk_script(script) == script


def test_safe_labtalk_script_blocks_project_reset() -> None:
    with pytest.raises(ValueError, match="doc -s"):
        safe_labtalk_script("doc -s; doc -n;")


def test_safe_labtalk_script_blocks_file_overwrite_commands() -> None:
    with pytest.raises(ValueError, match="save"):
        safe_labtalk_script('save "C:\\\\data\\\\old.opju";')


def test_safe_labtalk_script_blocks_window_delete() -> None:
    with pytest.raises(ValueError, match="win -cd"):
        safe_labtalk_script("win -cd Book1;")
