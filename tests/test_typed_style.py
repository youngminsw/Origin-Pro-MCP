"""Guard tests for the typed error-bar / column / style tools (no Origin)."""
import pytest


# --- set_column_designation --------------------------------------------------

def test_set_column_designation_yerr(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_designation

    msg = set_column_designation("Book1", "Sheet1", 3, "yerr")
    assert "yerr" in msg
    assert any("wks.col3.type = 3" in s for s in fake_origin.executed)


def test_set_column_designation_xerr_is_seven(fake_origin):
    # Regression: xerr used to collide with label (both 5); it must be 7.
    from origin_pro_mcp.tools.worksheet import set_column_designation

    set_column_designation("Book1", "Sheet1", 4, "xerr")
    assert any("wks.col4.type = 7" in s for s in fake_origin.executed)


def test_set_column_designation_bad_role(fake_origin):
    from origin_pro_mcp.tools.worksheet import set_column_designation

    with pytest.raises(ValueError, match="role must be one of"):
        set_column_designation("Book1", "Sheet1", 1, "abscissa")


# --- set_error_bars ----------------------------------------------------------

def test_set_error_bars_y_uses_set_o(fake_origin):
    from origin_pro_mcp.tools.graph import set_error_bars

    msg = set_error_bars("Graph1", "Book1", "Sheet1", y_col=2, err_col=3)
    assert "y-error bars" in msg
    joined = " ".join(fake_origin.executed)
    assert "plotxy iy:=[Book1]Sheet1!col(3)" in joined
    assert "set __mcp_er -o __mcp_yr" in joined
    # N1 fix: designate the err column as Y Error (3) and rebuild the legend so
    # no stray "SD" curve/entry is left behind.
    assert "wks.col3.type = 3" in joined
    assert "legend -r" in joined


def test_set_error_bars_x_direction(fake_origin):
    from origin_pro_mcp.tools.graph import set_error_bars

    set_error_bars("Graph1", "Book1", "Sheet1", y_col=2, err_col=3, direction="x")
    assert any("set __mcp_er -ox __mcp_yr" in s for s in fake_origin.executed)


def test_set_error_bars_same_column_rejected(fake_origin):
    from origin_pro_mcp.tools.graph import set_error_bars

    with pytest.raises(ValueError, match="must be different"):
        set_error_bars("Graph1", "Book1", "Sheet1", y_col=2, err_col=2)


def test_set_error_bars_unknown_graph(fake_origin):
    from origin_pro_mcp.tools.graph import set_error_bars

    with pytest.raises(ValueError, match="not found"):
        set_error_bars("Ghost", "Book1", "Sheet1", y_col=2, err_col=3)


# --- set_plot_style open/solid markers ---------------------------------------

def test_set_plot_style_open_symbol(fake_origin):
    from fakes import FakeGraph
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    fake_origin.lt_vars["__mcpk"] = 1  # make _plot_has_symbols() report a symbol plot
    set_plot_style("Graph1", plot_index=1, open_symbol=True)
    assert any("-kf 1" in s for s in fake_origin.executed)


def test_set_plot_style_solid_symbol_default(fake_origin):
    from fakes import FakeGraph
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    fake_origin.lt_vars["__mcpk"] = 1
    set_plot_style("Graph1", plot_index=1)
    assert any("-kf 0" in s for s in fake_origin.executed)

# --- set_graph_font bold -----------------------------------------------------

def test_set_graph_font_bold_axes(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style as S

    # Origin 2020 has no xb.bold; bold wraps the title text in \b(...).
    monkeypatch.setattr(S, "get_lt_str", lambda name: "Temperature (K)")
    S.set_graph_font("Graph1", target="axes", bold=True)
    joined = " ".join(fake_origin.executed)
    assert r'xb.text$ = "\b(Temperature (K))"' in joined
    assert r'yl.text$ = "\b(Temperature (K))"' in joined
    assert not any("xb.bold" in s for s in fake_origin.executed)


def test_set_graph_font_bold_skips_already_bold(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style as S

    monkeypatch.setattr(S, "get_lt_str", lambda name: r"\b(Already)")
    S.set_graph_font("Graph1", target="axes", bold=True)
    assert not any(".text$ =" in s for s in fake_origin.executed)  # no double-wrap


def test_set_graph_font_no_bold_by_default(fake_origin):
    from origin_pro_mcp.tools.style import set_graph_font

    set_graph_font("Graph1", target="axes")
    # No title rewrap and no (nonexistent) .bold command on axis titles.
    assert not any(".text$ =" in s for s in fake_origin.executed)
    assert not any("xb.bold" in s for s in fake_origin.executed)


# --- set_tick_labels ---------------------------------------------------------

def _capture_layer(monkeypatch):
    """set_tick_labels routes layer.* props through graph_layer_execute (the
    graph layer's Execute, not Origin's) — capture those commands."""
    from origin_pro_mcp.tools import style as S

    sink = []

    def fake(graph_name, script):
        sink.append(script)
        return True

    monkeypatch.setattr(S, "graph_layer_execute", fake)
    return sink


def test_set_tick_labels_scientific_both(fake_origin, monkeypatch):
    from origin_pro_mcp.tools.style import set_tick_labels

    sink = _capture_layer(monkeypatch)
    set_tick_labels("Graph1", axis="both", format="scientific")
    joined = " ".join(sink)
    assert "layer.x.label.numFormat = 2" in joined
    assert "layer.y.label.numFormat = 2" in joined


def test_set_tick_labels_bold_and_decimals(fake_origin, monkeypatch):
    from origin_pro_mcp.tools.style import set_tick_labels

    sink = _capture_layer(monkeypatch)
    set_tick_labels("Graph1", axis="x", bold=True, decimal_places=2)
    joined = " ".join(sink)
    assert "layer.x.label.bold = 1" in joined
    assert "layer.x.label.decPlaces = 2" in joined


def test_set_tick_labels_requires_an_argument(fake_origin):
    from origin_pro_mcp.tools.style import set_tick_labels

    with pytest.raises(ValueError, match="at least one of"):
        set_tick_labels("Graph1", axis="x")


def test_set_tick_labels_bad_format(fake_origin):
    from origin_pro_mcp.tools.style import set_tick_labels

    with pytest.raises(ValueError, match="format must be one of"):
        set_tick_labels("Graph1", format="hex")


# --- set_layer_geometry ------------------------------------------------------

def test_set_layer_geometry_sets_fields(fake_origin):
    from origin_pro_mcp.tools.graph import set_layer_geometry

    set_layer_geometry("Graph1", left=15, top=12, width=75, height=75)
    joined = " ".join(fake_origin.executed)
    assert "layer.left = 15.0" in joined
    assert "layer.width = 75.0" in joined


def test_set_layer_geometry_partial(fake_origin):
    from origin_pro_mcp.tools.graph import set_layer_geometry

    set_layer_geometry("Graph1", width=80)
    joined = " ".join(fake_origin.executed)
    assert "layer.width = 80.0" in joined
    assert "layer.left" not in joined


def test_set_layer_geometry_requires_an_argument(fake_origin):
    from origin_pro_mcp.tools.graph import set_layer_geometry

    with pytest.raises(ValueError, match="at least one of"):
        set_layer_geometry("Graph1")
