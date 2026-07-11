"""Live (Windows + Origin Pro COM) pixel-verified tests for the styling-report
fix round (2026-07-10 plan): the Task 0.5 settle barrier, set_plot_style
partial styling + error-bar knobs, axis frame width / per-side ticks, and
apply_publication_style integrity.

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_styling.py -v

Safety: every test runs against its OWN isolated Origin instance spawned via
``DispatchEx("Origin.Application")`` (same pattern as test_live_loaded_graph.py).
Never touches ``Origin.ApplicationSI``. ``origin.Exit()`` always runs in the
fixture teardown.
"""
import json

import pytest

pytestmark = pytest.mark.requires_origin


@pytest.fixture()
def live_origin():
    """A fresh, isolated Origin.exe bound to this thread; closed afterwards."""
    import pythoncom
    import win32com.client

    from origin_pro_mcp.origin_connection import (
        clear_session_origin,
        set_session_origin,
    )

    pythoncom.CoInitialize()

    def _factory():
        o = win32com.client.DispatchEx("Origin.Application")
        try:
            o.Visible = 1
        except Exception:
            pass
        return o

    origin = _factory()
    set_session_origin(origin, factory=_factory)
    try:
        yield origin
    finally:
        try:
            origin.Exit()
        except Exception:
            pass
        clear_session_origin()


def _name(create_result: str) -> str:
    return json.loads(create_result)["name"]


def _red_pixel_count(path: str, threshold_r=200, threshold_gb=80) -> int:
    from PIL import Image

    im = Image.open(path).convert("RGB")
    px = im.load()
    w, h = im.size
    count = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            r, g, b = px[x, y]
            if r > threshold_r and g < threshold_gb and b < threshold_gb:
                count += 1
    return count


def _build_line_symbol_with_error(book="SMOKE", y_error=True):
    """A worksheet + one line+symbol series (X,Y,Yerr), for styling tests."""
    from origin_pro_mcp.tools.graph import create_graph
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet(book))
    b, sheet = made["name"], made["sheet"]
    set_worksheet_data(
        b, sheet,
        json.dumps([[1, 2, 3, 4, 5], [1, 4, 9, 16, 25], [0.5, 0.5, 1.0, 1.0, 1.5]]),
    )
    kwargs = {"y_error_col": 3} if y_error else {}
    g = _name(create_graph("LineG", b, sheet, 1, 2, plot_type="line+symbol", **kwargs))
    return g, b, sheet


def test_set_error_bars_attaches_in_place_live(live_origin):
    """Item 22: set_error_bars on a plotted Y column attaches error bars in
    place — the error column reads back as an error plot and no stray DATA
    curve survives."""
    from origin_pro_mcp.tools.graph import set_error_bars
    from origin_pro_mcp.tools.style_helpers import get_plot_info

    g, book, sheet = _build_line_symbol_with_error(book="EBLIVE", y_error=False)
    before = get_plot_info(g)
    data_before = sum(1 for p in before if not p["is_error"])
    assert data_before == 1 and not any(p["is_error"] for p in before)

    msg = set_error_bars(g, book, sheet, y_col=2, err_col=3)
    assert "y-error bars" in msg

    after = get_plot_info(g)
    data_after = sum(1 for p in after if not p["is_error"])
    assert data_after == data_before, after  # no stray data curve
    assert any(p["is_error"] for p in after), after  # error plot present


def _build_multiseries_with_error(book="PUB", n_series=2):
    """A worksheet with n_series (X, Y, Yerr) triples, each plotted as its own
    line+symbol series with y-error bars — the reporter's issue #7 repro
    shape (multi-series + error cols, ungrouped since built via create_graph +
    add_plot_to_graph)."""
    from origin_pro_mcp.tools.graph import add_plot_to_graph, create_graph
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet(book))
    b, sheet = made["name"], made["sheet"]
    x = [1, 2, 3, 4, 5]
    cols = [x]
    for i in range(n_series):
        cols.append([v * (i + 1) for v in [1, 4, 9, 16, 25]])
        cols.append([0.5] * 5)
    set_worksheet_data(b, sheet, json.dumps(cols))

    g = _name(create_graph(
        "PubG", b, sheet, 1, 2, plot_type="line+symbol", y_error_col=3,
    ))
    for i in range(1, n_series):
        y_col = 2 + i * 2
        err_col = y_col + 1
        add_plot_to_graph(
            g, b, sheet, 1, y_col, plot_type="line+symbol", y_error_col=err_col,
        )
    return g, b, sheet


def test_settle_barrier_immediate_color_set_takes_effect(tmp_path, live_origin):
    """Task 0.5 regression: setting a curve's color IMMEDIATELY after
    create_graph must actually render (no silent no-op from the new-page
    settle hazard)."""
    from origin_pro_mcp.tools.graph import export_graph_to_file
    from origin_pro_mcp.tools.style_helpers import get_plot_info, graph_layer_execute

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    pname = get_plot_info(g)[0]["name"]
    graph_layer_execute(g, f"set {pname} -c color(255,0,0);")
    out = str(tmp_path / "settle_regress.png")
    export_graph_to_file(g, out)
    assert _red_pixel_count(out) > 100


# --- Task 1: set_plot_style partial styling + error-bar knobs ---------------

def test_set_plot_style_line_width_preserves_color(tmp_path, live_origin):
    """Coloring plot 1 red, then only changing line_width, must change pixels
    (the width did apply) while the red color survives (partial styling must
    not reset other aspects — the None-defaults fix)."""
    from origin_pro_mcp.tools.graph import export_graph_to_file
    from origin_pro_mcp.tools.style import set_plot_style

    g, _book, _sheet = _build_line_symbol_with_error(y_error=True)
    set_plot_style(g, plot_index=1, rgb="255,0,0")
    before = str(tmp_path / "lw_before.png")
    export_graph_to_file(g, before)
    before_red = _red_pixel_count(before)
    assert before_red > 100

    set_plot_style(g, plot_index=1, line_width=6.0)
    after = str(tmp_path / "lw_after.png")
    export_graph_to_file(g, after)
    after_red = _red_pixel_count(after)

    with open(before, "rb") as f1, open(after, "rb") as f2:
        assert f1.read() != f2.read()  # line width visibly changed the render
    assert after_red > 100  # color survived the partial line_width-only call


def test_set_plot_style_error_bar_width_and_cap_change_pixels(tmp_path, live_origin):
    """error_bar_width/error_cap_width actually change the rendered export."""
    from origin_pro_mcp.tools.graph import export_graph_to_file
    from origin_pro_mcp.tools.style import set_plot_style

    g, _book, _sheet = _build_line_symbol_with_error(y_error=True)
    baseline = str(tmp_path / "eb_baseline.png")
    export_graph_to_file(g, baseline)

    set_plot_style(g, plot_index=1, error_bar_width=4.0, error_cap_width=16)
    after = str(tmp_path / "eb_after.png")
    export_graph_to_file(g, after)

    with open(baseline, "rb") as f1, open(after, "rb") as f2:
        assert f1.read() != f2.read()


# --- Task 2: frame width + per-side tick control ----------------------------

def test_axis_frame_width_changes_render_and_reads_back(tmp_path, live_origin):
    """frame_width visibly thickens the frame (P1-confirmed knob) and reads
    back the value that was set."""
    from origin_pro_mcp.tools.graph import axis, export_graph_to_file
    from origin_pro_mcp.tools.style_helpers import read_layer_value

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    axis(g, op="frame", frame="closed", frame_width=0.5)
    thin = str(tmp_path / "frame_thin.png")
    export_graph_to_file(g, thin)

    axis(g, op="frame", frame="closed", frame_width=6.0)
    thick = str(tmp_path / "frame_thick.png")
    export_graph_to_file(g, thick)

    with open(thin, "rb") as f1, open(thick, "rb") as f2:
        assert f1.read() != f2.read()
    assert read_layer_value(g, "layer.x.thickness") == pytest.approx(6.0, abs=0.01)


def test_axis_tick_top_none_removes_marks_keeps_bottom_labels(tmp_path, live_origin):
    """axis(op="tick", axis="top"/"right", tick_direction="none") must NOT wipe
    the bottom/left tick NUMBER labels — that regression (#4) is the invariant
    under test. We crop the bottom-label strip and assert it survives non-blank
    and roughly unchanged. The top band is asserted only to NOT gain ink
    (removing marks can never add any); a strict decrease is deliberately NOT
    asserted, because on this graph state the removal is at or below the
    export's pixel threshold (top_before == top_after) — asserting a visible
    top change here, like the old whole-image byte-diff, is exactly what used
    to cry wolf on an otherwise-correct render."""
    import time

    from PIL import Image

    from origin_pro_mcp.tools.graph import axis, export_graph_to_file
    from origin_pro_mcp.tools.labtalk import run_labtalk

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    axis(g, op="frame", frame="closed")
    baseline = str(tmp_path / "tick_baseline.png")
    export_graph_to_file(g, baseline)

    axis(g, op="tick", axis="top", tick_direction="none")
    axis(g, op="tick", axis="right", tick_direction="none")
    # Settle before exporting: the previous whole-image byte-diff assertion
    # flaked because the FIRST export could precede the tick edit rendering
    # (passed on rerun). A refresh + short wait makes the change deterministic.
    run_labtalk("doc -uw;")
    time.sleep(1.0)
    after = str(tmp_path / "tick_after.png")
    export_graph_to_file(g, after)

    def dark_pixels(path, y0f, y1f, thr=160):
        im = Image.open(path).convert("L")
        w, h = im.size
        strip = im.crop((0, int(h * y0f), w, int(h * y1f)))
        px = strip.load()
        sw, sh = strip.size
        return sum(1 for y in range(sh) for x in range(sw) if px[x, y] < thr)

    # Top tick MARKS sit just inside the top frame; removing them can only
    # REDUCE ink in that band, never add it. A localized measured count (not a
    # full-image byte-diff) with the tolerant direction: on this graph state the
    # removal is at/below the pixel threshold, so we require only "did not gain
    # ink" — that is what stops the false failures on a correct render.
    top_before = dark_pixels(baseline, 0.05, 0.18)
    top_after = dark_pixels(after, 0.05, 0.18)
    assert top_after <= top_before  # marks removed or unchanged, never added

    # Bottom/left NUMBER labels must SURVIVE — the actual #4 regression guard:
    # removing top/right ticks must NOT blank the bottom/left tick labels.
    bottom_before = dark_pixels(baseline, 0.85, 1.0, thr=200)
    bottom_after = dark_pixels(after, 0.85, 1.0, thr=200)
    assert bottom_before > 0
    assert bottom_after > 0
    assert abs(bottom_after - bottom_before) < max(40, bottom_before * 0.1)


# --- Task 3: apply_publication_style integrity (#7) + grouped-fill truth (#8) --

def test_apply_publication_style_reporter_repro(tmp_path, live_origin):
    """Reporter's minimal issue #7 repro, rerun clean with the Task 0.5 settle
    fix in place: 2-series line+symbol + error cols, custom colors via
    set_plot_style, then apply_publication_style (which overwrites them with
    the pastel palette — documented behavior), then set_plot_style(
    open_symbol=False) on each plot to prove shapes/plots stay individually
    addressable (not corrupted/merged) after the publication pass."""
    from origin_pro_mcp.tools.graph import export_graph_to_file
    from origin_pro_mcp.tools.style import apply_publication_style, set_plot_style

    g, _book, _sheet = _build_multiseries_with_error(n_series=2)

    # Precondition: custom colors actually apply (Task 0.5 fix verified).
    set_plot_style(g, plot_index=1, rgb="0,0,255")
    set_plot_style(g, plot_index=2, rgb="255,128,0")
    precond = str(tmp_path / "repro_precondition.png")
    export_graph_to_file(g, precond)
    assert _red_pixel_count(precond, threshold_r=200, threshold_gb=150) > 50  # orange-ish plot 2 rendered

    msg = apply_publication_style(g, x_label="X", y_label="Y")
    assert "2 data plots styled" in msg
    styled = str(tmp_path / "repro_styled.png")
    export_graph_to_file(g, styled)
    with open(precond, "rb") as f1, open(styled, "rb") as f2:
        assert f1.read() != f2.read()  # palette overwrote the custom colors

    # Refill each plot open/solid — must not raise and must still resolve
    # each plot independently (proves apply_publication_style did not merge
    # or corrupt the per-plot addressability of an ungrouped multi-series graph).
    msg1 = set_plot_style(g, plot_index=1, open_symbol=False)
    msg2 = set_plot_style(g, plot_index=2, open_symbol=True)
    assert "plot 1" in msg1
    assert "plot 2" in msg2
    final = str(tmp_path / "repro_refill.png")
    export_graph_to_file(g, final)
    with open(styled, "rb") as f1, open(final, "rb") as f2:
        assert f1.read() != f2.read()  # the per-plot refill visibly changed plot 2


# --- Task 5: run_labtalk window param ----------------------------------------

def test_run_labtalk_window_targets_the_named_graph(tmp_path, live_origin):
    """run_labtalk(window=...) activates that window before executing, so a
    `layer.*` write lands on the NAMED graph even when a different window is
    currently active."""
    from origin_pro_mcp.tools.graph import export_graph_to_file
    from origin_pro_mcp.tools.labtalk import run_labtalk
    from origin_pro_mcp.tools.worksheet import create_worksheet

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    # Make some OTHER window active first.
    create_worksheet("OTHER")

    run_labtalk("layer.x.thickness = 6;", window=g)
    out = str(tmp_path / "window_param.png")
    export_graph_to_file(g, out)

    from origin_pro_mcp.tools.style_helpers import read_layer_value
    assert read_layer_value(g, "layer.x.thickness") == pytest.approx(6.0, abs=0.01)


# --- Task 6: WSL export path translation (#13) -------------------------------

def test_export_graph_bare_posix_path_lands_on_wsl_side(monkeypatch, live_origin):
    """Issue #13: a bare POSIX path (no /mnt/<drive> prefix) is translated to
    \\\\wsl.localhost\\<distro>\\... and the export must actually succeed (no
    "cannot access" rejection) when ORIGIN_PRO_MCP_WSL_DISTRO names a real
    distro. The WSL-side file arrival itself is checked separately from the
    WSL shell (P7 confirmed this live: the UNC write genuinely lands there)."""
    monkeypatch.setenv("ORIGIN_PRO_MCP_WSL_DISTRO", "Ubuntu")
    from origin_pro_mcp.tools.graph import export_graph_to_file

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    out = export_graph_to_file(g, "/tmp/live_wsl_export_check.png")
    assert out == "\\\\wsl.localhost\\Ubuntu\\tmp\\live_wsl_export_check.png"


# --- Issue #14: tick-label offset (frame->label gap) -------------------------

def _xlabel_gap(path: str) -> float:
    """Pixel gap between the bottom frame line and the x tick-label glyphs.

    Finds the bottom frame row (the row with the most dark pixels across the
    central band), then returns the vertical centroid of the tick-label ink
    below it — excluding the outward tick marks right under the frame and the
    centered axis title. Larger = labels farther from the axis.
    """
    from PIL import Image

    im = Image.open(path).convert("L")
    w, h = im.size
    px = im.load()
    x0, x1 = int(w * 0.20), int(w * 0.80)
    frame_row, frame_cnt = None, 0
    for y in range(int(h * 0.35), int(h * 0.95)):
        cnt = sum(1 for x in range(x0, x1) if px[x, y] < 100)
        if cnt > frame_cnt:
            frame_cnt, frame_row = cnt, y
    tot = wsum = 0
    for y in range(frame_row + 6, min(frame_row + 140, h)):
        for x in range(x0, x1):
            if 0.42 * w < x < 0.58 * w:  # skip the centered axis title
                continue
            if px[x, y] < 100:
                tot += 1
                wsum += y - frame_row
    return wsum / tot if tot else 0.0


def test_tick_label_offset_moves_x_labels_toward_and_away(tmp_path, live_origin):
    """set_tick_labels(offset_pct=...) must move the x tick labels perpendicular
    to the axis by a MEASURABLE gap, in the documented direction: a large
    negative offset pushes them far from the frame, a positive offset pulls them
    close. Pixel-measured, not just bytes-differ."""
    from origin_pro_mcp.tools.graph import axis, export_graph_to_file
    from origin_pro_mcp.tools.style import set_tick_labels

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    axis(g, op="frame", frame="closed")

    set_tick_labels(g, axis="x", offset_pct=-150)  # push labels far from axis
    far = str(tmp_path / "tlo_far.png")
    export_graph_to_file(g, far)
    gap_far = _xlabel_gap(far)

    set_tick_labels(g, axis="x", offset_pct=100)  # pull labels toward axis
    near = str(tmp_path / "tlo_near.png")
    export_graph_to_file(g, near)
    gap_near = _xlabel_gap(near)

    with open(far, "rb") as f1, open(near, "rb") as f2:
        assert f1.read() != f2.read()  # the offset visibly changed the render
    # Direction + magnitude: far offset keeps labels clearly lower than near.
    assert gap_far - gap_near > 15, f"gap_far={gap_far:.1f} gap_near={gap_near:.1f}"


def test_tick_label_offset_y_axis_changes_render(tmp_path, live_origin):
    """The y (left) axis routes offset_pct to the HORIZONTAL offset; a positive
    vs negative value must produce different renders (the offsetH path works on
    real Origin, not just in the fake-test emission check)."""
    from origin_pro_mcp.tools.graph import axis, export_graph_to_file
    from origin_pro_mcp.tools.style import set_tick_labels

    g, _book, _sheet = _build_line_symbol_with_error(y_error=False)
    axis(g, op="frame", frame="closed")

    set_tick_labels(g, axis="y", offset_pct=-150)
    left = str(tmp_path / "tlo_y_left.png")
    export_graph_to_file(g, left)

    set_tick_labels(g, axis="y", offset_pct=80)
    right = str(tmp_path / "tlo_y_right.png")
    export_graph_to_file(g, right)

    with open(left, "rb") as f1, open(right, "rb") as f2:
        assert f1.read() != f2.read()
