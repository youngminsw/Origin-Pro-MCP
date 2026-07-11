"""Guard tests for the typed error-bar / column / style tools (no Origin)."""
import pytest


# --- column designation via manage_columns(op="properties") ------------------

def test_manage_columns_designation_yerr(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    msg = manage_columns("Book1", "Sheet1", op="properties", col=3, designation="yerr")
    assert "designation" in msg
    assert any("wks.col3.type = 3" in s for s in fake_origin.executed)


def test_manage_columns_designation_xerr_is_seven(fake_origin):
    # Regression: xerr used to collide with label (both 5); it must be 7.
    from origin_pro_mcp.tools.worksheet import manage_columns

    manage_columns("Book1", "Sheet1", op="properties", col=4, designation="xerr")
    assert any("wks.col4.type = 7" in s for s in fake_origin.executed)


def test_manage_columns_designation_bad_role(fake_origin):
    from origin_pro_mcp.tools.worksheet import manage_columns

    with pytest.raises(ValueError, match="designation must be one of"):
        manage_columns("Book1", "Sheet1", op="properties", col=1, designation="abscissa")


# --- set_error_bars ----------------------------------------------------------

def _fake_attach_success(fake_origin, monkeypatch, *, desig, err_is_error=True):
    """Wire the graph module so set_error_bars sees a successful attach: settle
    is a no-op, get_plot_info returns [data] before then [data, err] after, and
    the err column's designation reads back as `desig`."""
    from origin_pro_mcp.tools import graph

    monkeypatch.setattr(graph, "settle_new_plots", lambda *a, **k: None)
    seq = [
        [{"name": "G_B", "is_error": False}],
        [{"name": "G_B", "is_error": False},
         {"name": "G_C", "is_error": err_is_error}],
    ]
    calls = {"n": 0}

    def _infos(_g):
        i = min(calls["n"], len(seq) - 1)
        calls["n"] += 1
        return seq[i]

    monkeypatch.setattr(graph, "get_plot_info", _infos)
    fake_origin.lt_vars["__mcp_ct"] = float(desig)


def test_set_error_bars_y_settles_and_attaches(fake_origin, monkeypatch):
    """Item 22: the plotxy is followed by a settle, then the reassignment; a
    verified attach (col.type==3, error plot present, no stray data curve)
    returns success."""
    from origin_pro_mcp.tools.graph import set_error_bars

    _fake_attach_success(fake_origin, monkeypatch, desig=3)
    msg = set_error_bars("Graph1", "Book1", "Sheet1", y_col=2, err_col=3)
    assert "y-error bars" in msg
    joined = " ".join(fake_origin.executed)
    assert "plotxy iy:=[Book1]Sheet1!col(3)" in joined
    assert "set __mcp_er -o __mcp_yr" in joined
    assert "wks.col3.type = 3" in joined
    assert "legend -r" in joined


def test_set_error_bars_x_direction(fake_origin, monkeypatch):
    from origin_pro_mcp.tools.graph import set_error_bars

    # X error: verification keys on the X-Error designation (7) reading back.
    _fake_attach_success(fake_origin, monkeypatch, desig=7, err_is_error=False)
    set_error_bars("Graph1", "Book1", "Sheet1", y_col=2, err_col=3, direction="x")
    assert any("set __mcp_er -ox __mcp_yr" in s for s in fake_origin.executed)


def test_set_error_bars_fails_honestly_and_removes_stray(fake_origin, monkeypatch):
    """Item 22: when the error column does NOT convert (designation didn't
    stick, leaving a stray data curve), the tool removes the stray and raises —
    no false success."""
    from origin_pro_mcp.tools import graph
    from origin_pro_mcp.tools.graph import set_error_bars
    from conftest import FakeGraph

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["G_B", "G_C"])]
    monkeypatch.setattr(graph, "settle_new_plots", lambda *a, **k: None)
    # get_plot_info: [data] before, [data, stray-DATA] after (err not converted).
    seq = [
        [{"name": "G_B", "is_error": False}],
        [{"name": "G_B", "is_error": False}, {"name": "G_C", "is_error": False}],
    ]
    calls = {"n": 0}

    def _infos(_g):
        i = min(calls["n"], len(seq) - 1)
        calls["n"] += 1
        return seq[i]

    monkeypatch.setattr(graph, "get_plot_info", _infos)
    # __mcp_ct stays 0 (designation did not stick).
    with pytest.raises(ValueError, match="did not convert to error bars"):
        set_error_bars("Graph1", "Book1", "Sheet1", y_col=2, err_col=3)
    # The stray plot (last, index 1) was destroyed.
    assert fake_origin.graphs[0].plot_names == ["G_B"]


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
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    fake_origin.lt_vars["__mcpk"] = 1  # make _plot_has_symbols() report a symbol plot
    set_plot_style("Graph1", plot_index=1, open_symbol=True)
    layer = origin_connection.get_origin()._graph_layers["[Graph1]Layer1"]
    assert any("-kf 1" in s for s in layer.executed)


def test_set_plot_style_solid_symbol_explicit(fake_origin):
    from fakes import FakeGraph
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    fake_origin.lt_vars["__mcpk"] = 1
    set_plot_style("Graph1", plot_index=1, open_symbol=False)
    layer = origin_connection.get_origin()._graph_layers["[Graph1]Layer1"]
    assert any("-kf 0" in s for s in layer.executed)


def test_set_plot_style_nothing_requested_raises(fake_origin):
    from fakes import FakeGraph
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    with pytest.raises(ValueError, match="nothing to change"):
        set_plot_style("Graph1", plot_index=1)


# --- set_plot_style: None-defaults + P8 one-flag-per-call rule ---------------

def test_set_plot_style_line_width_only_emits_only_dash_w(fake_origin):
    """(a) line_width alone must emit ONLY `-w`; no -k/-z/-kf leak in from
    stale defaults."""
    from fakes import FakeGraph
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    fake_origin.lt_vars["__mcpk"] = 1  # this IS a symbol plot
    set_plot_style("Graph1", plot_index=1, line_width=5)
    layer = origin_connection.get_origin()._graph_layers["[Graph1]Layer1"]
    assert any("set Book1_B -w 1000;" == s for s in layer.executed)
    assert not any("-k " in s or "-z " in s or "-kf " in s for s in layer.executed)


def test_set_plot_style_rgb_sends_c_and_cf_as_separate_calls(fake_origin):
    """(b) P8 hard rule: -c and -cf must be two SEPARATE `set` commands, never
    combined in one string (combining wipes the plot to black on Origin 2020)."""
    from fakes import FakeGraph
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.style import set_plot_style

    fake_origin.graphs = [FakeGraph("Graph1", plot_names=["Book1_B"])]
    fake_origin.lt_vars["__mcpk"] = 1
    set_plot_style("Graph1", plot_index=1, rgb="255,0,0")
    layer = origin_connection.get_origin()._graph_layers["[Graph1]Layer1"]
    assert any("-c color(255,0,0)" in s for s in layer.executed)
    assert any("-cf color(255,0,0)" in s for s in layer.executed)
    assert not any("-c " in s and "-cf " in s for s in layer.executed)


# --- set_plot_style: error-bar width/cap -------------------------------------

def _eb_graph(fake, plot_names, columns):
    from fakes import FakeBook, FakeGraph, FakeSheet

    fake.books = [FakeBook("EB", sheets=[FakeSheet("Sheet1", columns=columns)])]
    fake.graphs = [FakeGraph("EB", plot_names=plot_names)]


def test_set_plot_style_error_bar_width_adjacent(fake_origin):
    """(d) error_bar_width/error_cap_width target the error plot immediately
    following the data plot in get_plot_info order (P6-confirmed adjacency)."""
    from fakes import FakeColumn
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.style import set_plot_style

    _eb_graph(fake_origin, ["EB_B", "EB_C"], [
        FakeColumn("A"), FakeColumn("B"), FakeColumn("C", col_type=2),
    ])
    set_plot_style("EB", plot_index=1, error_bar_width=2.5, error_cap_width=12)
    layer = origin_connection.get_origin()._graph_layers["[EB]Layer1"]
    assert any("set EB_C -erw 2.5;" == s for s in layer.executed)
    assert any("set EB_C -erwc 12;" == s for s in layer.executed)


def test_set_plot_style_error_bar_falls_back_when_not_adjacent(fake_origin):
    """When no error plot is directly adjacent, fall back to ALL error plots
    on the layer and note it in the return message."""
    from fakes import FakeColumn
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.style import set_plot_style

    _eb_graph(fake_origin, ["EB_B", "EB_D", "EB_C"], [
        FakeColumn("A"), FakeColumn("B"), FakeColumn("C", col_type=2), FakeColumn("D"),
    ])
    msg = set_plot_style("EB", plot_index=1, error_bar_width=3.0)
    layer = origin_connection.get_origin()._graph_layers["[EB]Layer1"]
    assert any("set EB_C -erw 3.0;" == s for s in layer.executed)
    assert "not directly adjacent" in msg or "ALL error plots" in msg


def test_set_plot_style_error_bar_requires_error_plots(fake_origin):
    """(e) error_bar_width/error_cap_width on a graph with NO error bars must
    raise, naming set_error_bars/y_error_col."""
    from fakes import FakeColumn
    from origin_pro_mcp.tools.style import set_plot_style

    _eb_graph(fake_origin, ["EB_B"], [FakeColumn("A"), FakeColumn("B")])
    with pytest.raises(ValueError, match="set_error_bars|y_error_col"):
        set_plot_style("EB", plot_index=1, error_bar_width=2.0)

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


def test_set_graph_font_tick_honors_bold_false(fake_origin):
    """Item 12: tick labels honor bold=False (previously always-bold)."""
    from origin_pro_mcp.tools.style import set_graph_font

    set_graph_font("Graph1", target="tick", bold=False)
    joined = " ".join(fake_origin.executed)
    assert "layer.x.label.bold = 0" in joined
    assert "layer.y.label.bold = 0" in joined


def test_set_graph_font_tick_bold_true(fake_origin):
    """Item 12: bold=True still bolds tick labels."""
    from origin_pro_mcp.tools.style import set_graph_font

    set_graph_font("Graph1", target="tick", bold=True)
    joined = " ".join(fake_origin.executed)
    assert "layer.x.label.bold = 1" in joined
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


def test_set_tick_labels_offset_x_uses_offsetV(fake_origin, monkeypatch):
    """The x (bottom) axis's perpendicular gap knob is the VERTICAL offset."""
    from origin_pro_mcp.tools.style import set_tick_labels

    sink = _capture_layer(monkeypatch)
    set_tick_labels("Graph1", axis="x", offset_pct=150)
    joined = " ".join(sink)
    assert "layer.x.label.offsetV = 150" in joined
    assert "offsetH" not in joined


def test_set_tick_labels_offset_y_uses_offsetH(fake_origin, monkeypatch):
    """The y (left) axis's perpendicular gap knob is the HORIZONTAL offset."""
    from origin_pro_mcp.tools.style import set_tick_labels

    sink = _capture_layer(monkeypatch)
    set_tick_labels("Graph1", axis="y", offset_pct=150)
    joined = " ".join(sink)
    assert "layer.y.label.offsetH = 150" in joined
    assert "offsetV" not in joined


def test_set_tick_labels_offset_both_and_negative(fake_origin, monkeypatch):
    from origin_pro_mcp.tools.style import set_tick_labels

    sink = _capture_layer(monkeypatch)
    set_tick_labels("Graph1", axis="both", offset_pct=-100)
    joined = " ".join(sink)
    assert "layer.x.label.offsetV = -100" in joined
    assert "layer.y.label.offsetH = -100" in joined


def test_set_tick_labels_offset_alone_satisfies_guard(fake_origin, monkeypatch):
    """offset_pct on its own is a valid edit (no 'provide at least one' raise)."""
    from origin_pro_mcp.tools.style import set_tick_labels

    _capture_layer(monkeypatch)
    msg = set_tick_labels("Graph1", axis="x", offset_pct=0)
    assert "offset" in msg


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


def test_set_layer_geometry_raises_on_frozen_graph(fake_origin):
    """Item 26a: routed through the rebind-safe path with a read-back — a frozen
    (loaded-.opju) layer silently drops the write, so this must raise instead of
    reporting success."""
    from fakes import FakeGraph
    from origin_pro_mcp.tools.graph import set_layer_geometry

    fake_origin.graphs = [FakeGraph("F", frozen=True)]
    with pytest.raises(ValueError, match="did not take"):
        set_layer_geometry("F", left=15, top=12, width=75, height=75)


def test_apply_publication_style_raises_when_frame_fails(fake_origin):
    """Item 26b: a FRAME failure is structural — apply_publication_style must
    raise, not silently omit the frame."""
    from origin_pro_mcp.tools.style import apply_publication_style

    fake_origin.execute_results["layer.x.opposite = 1; layer.y.opposite = 1;"] = False
    with pytest.raises(ValueError, match="close/thicken the frame"):
        apply_publication_style("Graph1")


def test_create_matrix_raises_on_dim_mismatch(fake_origin):
    """Item 26d: create_matrix reads wks dims back; a non-zero mismatch raises."""
    from origin_pro_mcp.tools.matrix import create_matrix

    fake_origin.LTStr = lambda name: "MX" if name == "page.name$" else ""
    fake_origin.lt_vars["wks.nrows"] = 3.0  # asked for 10, got 3 -> mismatch
    fake_origin.lt_vars["wks.ncols"] = 10.0
    with pytest.raises(ValueError, match="did not take"):
        create_matrix("MX", rows=10, cols=10)


# --- setter execute-result checks (consistency with raise-on-failure) -------

def test_set_graph_font_raises_on_execute_failure(fake_origin):
    from origin_pro_mcp.tools import style as S

    fake_origin.execute_results["xb.font$"] = False
    with pytest.raises(ValueError, match="x-axis title font"):
        S.set_graph_font("Graph1", target="axes")


def test_set_graph_font_axes_read_back_gate_raises_on_mismatch(fake_origin):
    """Item 4: the axis-title font is read back — a mismatching size (a silent
    wrong-window / frozen no-op that LabTalk still reports as success) raises."""
    from origin_pro_mcp.tools import style as S

    fake_origin.lt_vars["__mcp_fs"] = 99.0  # read-back != requested 24
    with pytest.raises(ValueError, match="x-axis title font did not take"):
        S.set_graph_font("Graph1", target="axes", font_size=24)


def test_set_legend_warns_when_legend_unreadable(fake_origin):
    """Item 4: after the rebuild, an unreadable legend object (frozen loaded
    graph) yields a WARNING, not a clean success."""
    from origin_pro_mcp.tools.style import set_legend

    fake_origin.execute_results["__mcp_fs = legend.fsize"] = False
    msg = set_legend("Graph1")
    assert "WARNING" in msg
    assert "could not confirm the legend rendered" in msg


def test_set_tick_labels_raises_on_execute_failure(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style as S

    monkeypatch.setattr(S, "graph_layer_execute", lambda graph_name, script: False)
    with pytest.raises(ValueError, match="Could not update tick labels"):
        S.set_tick_labels("Graph1", format="decimal")


def test_axis_tick_raises_on_execute_failure(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import style as S

    monkeypatch.setattr(S, "graph_layer_execute", lambda graph_name, script: False)
    with pytest.raises(ValueError, match="Could not update .* tick style"):
        S._set_tick_style_impl("Graph1")


def test_set_legend_hide_raises_on_execute_failure(fake_origin):
    from origin_pro_mcp.tools.style import set_legend

    fake_origin.execute_results["legend.show"] = False
    with pytest.raises(ValueError, match="Could not hide the legend"):
        set_legend("Graph1", visible=False)
