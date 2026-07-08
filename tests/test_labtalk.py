import pytest

from origin_pro_mcp.origin_connection import execute_labtalk, get_lt_var, get_lt_str
from origin_pro_mcp.tools.labtalk import get_labtalk_variable, run_labtalk


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

def test_run_labtalk_capture_returns_json_values(fake_origin):
    import json
    fake_origin.lt_vars["mean"] = 3.5
    result = run_labtalk("stats col(1); mean = stats.mean;", capture=["mean"])
    payload = json.loads(result)
    assert payload["status"] == "ok"
    assert payload["values"] == {"mean": 3.5}
    assert fake_origin.executed == ["stats col(1); mean = stats.mean;"]


def test_run_labtalk_capture_reports_error_status(fake_origin):
    import json
    fake_origin.execute_results["boom"] = False
    payload = json.loads(run_labtalk("boom;", capture=["x"]))
    assert payload["status"] == "error"


def test_run_labtalk_without_capture_keeps_plain_string(fake_origin):
    result = run_labtalk("col(1) = 1;")
    assert result == "Executed successfully: col(1) = 1;"


def test_run_labtalk_capture_still_gated(fake_origin):
    result = run_labtalk("del scratch;", capture=["x"])
    assert fake_origin.executed == []
    assert "NOT EXECUTED" in result


def test_run_labtalk_capture_rejects_bad_variable_name(fake_origin):
    with pytest.raises(ValueError):
        run_labtalk("x = 1;", capture=["bad name; del all"])


# --- get_labtalk_variable existence check ---


def test_get_labtalk_variable_returns_value_when_defined(fake_origin):
    fake_origin.lt_vars["_opm_lt_exist"] = 1.0
    fake_origin.lt_vars["x"] = 42.0
    assert get_labtalk_variable("x") == "42.0"


def test_get_labtalk_variable_raises_when_undefined(fake_origin):
    fake_origin.lt_vars["_opm_lt_exist"] = 0.0
    with pytest.raises(ValueError, match="is not defined"):
        get_labtalk_variable("nope")


def test_get_labtalk_variable_string_unaffected_by_exist_check(fake_origin):
    fake_origin.lt_vars["str$"] = "unused"
    # String variables bypass the exist() check entirely.
    get_labtalk_variable("str$")
    assert not any("exist(" in s for s in fake_origin.executed)
