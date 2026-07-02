"""N6: unsupported Origin text-markup escapes (notably \\q(), which pops a
blocking LaTeX/MiKTeX modal that wedges the daemon) must be rejected BEFORE
reaching Origin — in axis labels, titles, annotations, and column labels."""
import pytest

from origin_pro_mcp.labtalk_safe import labtalk_text, validate_text_escapes


@pytest.mark.parametrize("text", [
    r"\b(bold)", r"\i(it)", r"\u(u)", r"\+(2)", r"\-(2)", r"\g(a)",
    r"\f:Arial(x)", "Temperature (K)", "no escapes here", "",
    r"prefix \b(x) \g(m) suffix",
])
def test_supported_escapes_pass(text):
    validate_text_escapes(text, "f")          # no raise
    assert labtalk_text(text, "f") == f'"{text}"'


@pytest.mark.parametrize("text", [
    r"\q(d)", r"Mean Al depth, \q(d)", r"\z(x)", r"\c1(x)", r"\ab(x)",
    r"good \b(x) then bad \q(y)",
])
def test_unsupported_escapes_rejected(text):
    with pytest.raises(ValueError, match="unsupported Origin text escape"):
        validate_text_escapes(text, "f")
    with pytest.raises(ValueError, match="unsupported Origin text escape"):
        labtalk_text(text, "f")


def test_q_error_names_latex_hazard():
    with pytest.raises(ValueError, match="LaTeX"):
        validate_text_escapes(r"\q(d)", "y_label")


def test_axis_labels_reject_q_before_reaching_origin(fake_origin):
    from origin_pro_mcp.tools.graph import _set_axis_labels_impl
    with pytest.raises(ValueError, match="unsupported Origin text escape"):
        _set_axis_labels_impl("Graph1", y_label=r"Mean Al depth, \q(d) (A)")
    # nothing was pushed to Origin (no yl.text$ / xb.text$ dispatched)
    for script in fake_origin.executed:
        assert "yl.text$" not in script
        assert "xb.text$" not in script


def test_annotate_rejects_q(fake_origin):
    from origin_pro_mcp.tools.graph import annotate
    with pytest.raises(ValueError, match="unsupported Origin text escape"):
        annotate("Graph1", "text", x1=1.0, y1=2.0, text=r"peak \q(x)")
    for script in fake_origin.executed:
        assert "label -p" not in script  # never dispatched the annotation


def test_supported_escape_reaches_origin(fake_origin):
    from origin_pro_mcp.tools.graph import _set_axis_labels_impl
    _set_axis_labels_impl("Graph1", y_label=r"\g(m) (kg)")
    assert any("yl.text$" in s and r"\g(m)" in s for s in fake_origin.executed)
