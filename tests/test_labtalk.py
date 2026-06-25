import pytest

from origin_pro_mcp.origin_connection import execute_labtalk, get_lt_var, get_lt_str
from origin_pro_mcp.tools.labtalk import run_labtalk


# --- COM-backed primitives (Windows-only) ---


@pytest.mark.requires_origin
def test_execute_labtalk_simple():
    execute_labtalk("double __mcp_test = 42;")
    assert get_lt_var("__mcp_test") == 42.0


@pytest.mark.requires_origin
def test_execute_labtalk_string():
    execute_labtalk('string __mcp_str$ = "hello";')
    assert get_lt_str("__mcp_str$") == "hello"


# --- run_labtalk confirm-gate (runs anywhere via fake Origin) ---


def test_run_labtalk_executes_allowed_script(fake_origin):
    result = run_labtalk("col(1) = col(2) * 2;")
    assert fake_origin.executed == ["col(1) = col(2) * 2;"]
    assert "Executed successfully" in result


def test_run_labtalk_allows_save_without_confirm(fake_origin):
    result = run_labtalk('save "C:\\\\data\\\\out.opju";')
    assert fake_origin.executed == ['save "C:\\\\data\\\\out.opju";']
    assert "Executed successfully" in result


def test_run_labtalk_allows_string_literal_keyword(fake_origin):
    # 'save' / 'run' inside a string literal must not gate.
    script = 'col(1) = "save the day";'
    result = run_labtalk(script)
    assert fake_origin.executed == [script]
    assert "Executed successfully" in result


def test_run_labtalk_gated_token_not_executed_without_confirm(fake_origin):
    result = run_labtalk("doc -s; doc -n;")
    # Must NOT execute and must NOT touch Origin.
    assert fake_origin.executed == []
    assert "NOT EXECUTED" in result
    assert "doc -s" in result
    assert "confirm=True" in result


def test_run_labtalk_gated_token_executes_with_confirm(fake_origin):
    script = "doc -s; doc -n;"
    result = run_labtalk(script, confirm=True)
    assert fake_origin.executed == [script]
    assert "Executed successfully" in result


@pytest.mark.parametrize(
    "script",
    [
        "del C:\\temp\\scratch.txt;",
        "delete dataset1;",
        "win -cd Book1;",
        "label -r obj NewText;",
        "system.path$ = 1;",
        "getsavename(f$);",
        "run.section(file, main);",
    ],
)
def test_run_labtalk_each_gated_token_blocked_without_confirm(fake_origin, script):
    result = run_labtalk(script)
    assert fake_origin.executed == []
    assert "NOT EXECUTED" in result
    assert "confirm=True" in result


def test_run_labtalk_reports_errors_from_origin(fake_origin):
    fake_origin.execute_results["boom"] = False
    result = run_labtalk("boom;")
    assert fake_origin.executed == ["boom;"]
    assert "with errors" in result


def test_run_labtalk_default_confirm_is_false():
    import inspect

    sig = inspect.signature(run_labtalk)
    assert sig.parameters["confirm"].default is False
