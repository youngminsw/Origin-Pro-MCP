import pytest

from origin_pro_mcp.labtalk_safe import (
    classify_labtalk_script,
    labtalk_choice,
    labtalk_name,
    labtalk_path,
    labtalk_string,
    positive_column,
    safe_labtalk_script,
)


# --- Unchanged helper guards (signatures/behavior must be preserved) ---


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


# --- New tokenizer-aware classifier: allow-by-default ---


def test_classify_returns_three_tuple_shape() -> None:
    ok, requires_confirm, reason = classify_labtalk_script("col(1) = 5;")
    assert ok is True
    assert requires_confirm is False
    assert reason == ""


def test_classify_never_hard_blocks() -> None:
    # ok is essentially always True now: nothing is hard-blocked.
    for script in ("doc -s;", "del everything;", "col(1) = 1;", "system.path$;"):
        ok, _requires_confirm, _reason = classify_labtalk_script(script)
        assert ok is True


@pytest.mark.parametrize(
    "script",
    [
        'save "C:\\\\data\\\\old.opju";',
        'saveAs "C:\\\\data\\\\new.opju";',
        'open -w "C:\\\\data\\\\book.ogwu";',
        'expGraph type:=png path:="C:\\\\out.png";',
        'file -c "C:\\\\data\\\\notes.txt";',
        "col(1) = col(2) * 2;",
        'win -a Fig1; xb.text$ = "Temperature (K)"; layer.x.grid = 0;',
        "double tmp = 5; tmp = tmp + 1;",
        "type -b The fit converged;",
        "running = 1; mysum = running + 2;",
    ],
)
def test_classify_allows_core_ops_without_confirm(script: str) -> None:
    ok, requires_confirm, reason = classify_labtalk_script(script)
    assert ok is True
    assert requires_confirm is False, f"unexpectedly gated: {reason!r}"
    assert reason == ""


# --- False-positive regression: strings and comments are stripped ---


@pytest.mark.parametrize(
    "script",
    [
        'col(1) = "save the day";',
        'type "please run this";',
        'note$ = "remember to doc -s and del the file";',
        'msg$ = "system.path$ and getfilename()";',
        'label$ = "win -cd then label -r";',
    ],
)
def test_classify_ignores_keywords_inside_string_literals(script: str) -> None:
    _ok, requires_confirm, reason = classify_labtalk_script(script)
    assert requires_confirm is False, f"string literal triggered gate: {reason!r}"


@pytest.mark.parametrize(
    "script",
    [
        "col(1) = 5; // run.section later, also doc -s and del me",
        "// system.path$ = 1; getfilename();",
        "/* doc -n then win -ct */ col(1) = 1;",
        "col(1) = 1; /* multi\n line del delete\n system( */ col(2) = 2;",
    ],
)
def test_classify_ignores_keywords_inside_comments(script: str) -> None:
    _ok, requires_confirm, reason = classify_labtalk_script(script)
    assert requires_confirm is False, f"comment triggered gate: {reason!r}"


# --- Each CONFIRM-list token requires confirmation ---


CONFIRM_CASES = [
    ('system.path$ = "x";', "system."),
    ("y = system(1);", "system("),
    ("run.section(myfile, main);", "run.section"),
    ("run -e externalprog;", "run -"),
    ("dll loadlibrary mylib;", "dll"),
    ("dde connect server topic;", "dde"),
    ("getfilename(fname$);", "getfilename"),
    ("getsavename(fname$);", "getsavename"),
    ("doc -s;", "doc -s"),
    ("doc -n;", "doc -n"),
    ("del C:\\temp\\scratch.txt;", "del/delete"),
    ("delete dataset1;", "del/delete"),
    ("win -c Book1;", "win -c/-cd/-ct"),
    ("win -cd Book1;", "win -c/-cd/-ct"),
    ("win -ct Book1;", "win -c/-cd/-ct"),
    ("label -r obj NewText;", "label -r"),
]


@pytest.mark.parametrize("script, expected_label", CONFIRM_CASES)
def test_classify_gates_each_confirm_token(script: str, expected_label: str) -> None:
    ok, requires_confirm, reason = classify_labtalk_script(script)
    assert ok is True
    assert requires_confirm is True, f"token not gated: {script!r}"
    assert reason == expected_label


# --- Substring non-trigger: identifiers merely containing a gated word ---


@pytest.mark.parametrize(
    "script",
    [
        "system_id = 5;",
        "deleted_flag = 0;",
        "delete_me = col(1);",
        "double myrun = 2;",
        "myrun -e = 5;",
        'dllname$ = "lib.dll";',
        "ddevalue = 3;",
        "running = 1;",
        "labeling = 2;",
        "windowcount = 4;",
        "documents = 7;",
    ],
)
def test_classify_does_not_trigger_on_substrings(script: str) -> None:
    _ok, requires_confirm, reason = classify_labtalk_script(script)
    assert requires_confirm is False, f"substring falsely gated: {reason!r}"


def test_classify_reports_earliest_gated_token() -> None:
    # Two gated tokens present; the earliest one is reported.
    _ok, requires_confirm, reason = classify_labtalk_script("doc -s; del scratch;")
    assert requires_confirm is True
    assert reason == "doc -s"


# --- Backward-compatible shim: allow-by-default pass-through ---


def test_safe_labtalk_script_passes_through_allowed() -> None:
    script = 'win -a Fig1; xb.text$ = "Temperature (K)";'
    assert safe_labtalk_script(script) == script


def test_safe_labtalk_script_no_longer_hard_blocks() -> None:
    # The old denylist raised here; the shim now allows by default.
    assert safe_labtalk_script("doc -s; doc -n;") == "doc -s; doc -n;"
    assert safe_labtalk_script('save "C:\\\\data\\\\old.opju";') == (
        'save "C:\\\\data\\\\old.opju";'
    )
