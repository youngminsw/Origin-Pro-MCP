"""Live (Windows + Origin Pro COM) smoke tests for the S2 loaded-graph fix.

These run only where a real Origin is available (requires_origin; skipped in
WSL/CI). They build a multi-series graph in-session, round-trip it through a
saved .opju, then prove that per-curve styling and axis edits actually take
effect on the RELOADED graph — the exact case that used to return success while
changing nothing.

Run on the Windows side with:  pytest -m requires_origin tests/test_live_loaded_graph.py
"""
import json

import pytest

pytestmark = pytest.mark.requires_origin


def _name(create_result: str) -> str:
    return json.loads(create_result)["name"]


def _build_multiseries(book="Book1", sheet="Sheet1"):
    """Two ungrouped line series (X, YA, YB) on one graph. Returns its name."""
    from origin_pro_mcp.tools.graph import add_plot_to_graph, create_graph
    from origin_pro_mcp.tools.worksheet import set_worksheet_data

    set_worksheet_data(
        book, sheet,
        json.dumps([[1, 2, 3, 4], [10, 20, 15, 30], [12, 22, 17, 33]]),
        "X,YA,YB",
    )
    name = _name(create_graph("MULTI", book, sheet, x_col=1, y_col=2, plot_type="line"))
    add_plot_to_graph(name, book, sheet, x_col=1, y_col=3, plot_type="line")
    return name


def test_per_curve_color_changes_reloaded_graph(tmp_path):
    """Set a per-curve color on a graph LOADED from a .opju and confirm the
    exported pixels actually change (no silent no-op)."""
    from origin_pro_mcp.tools.graph import export_graph
    from origin_pro_mcp.tools.project import load_project, new_project, save_project
    from origin_pro_mcp.tools.style import set_plot_style

    new_project()
    name = _build_multiseries()
    before = str(tmp_path / "before.png")
    export_graph(name, before)
    proj = str(tmp_path / "roundtrip.opju")
    save_project(proj)

    # Reload from disk into a fresh session — this is the frozen-edit scenario.
    new_project()
    load_project(proj)

    msg = set_plot_style(name, plot_index=2, rgb="255,0,0", line_width=6)
    assert "Updated style" in msg  # must not raise / no-op on the loaded graph

    after = str(tmp_path / "after.png")
    export_graph(name, after)
    with open(before, "rb") as a, open(after, "rb") as b:
        assert a.read() != b.read(), "per-curve color edit did not change the reloaded graph"


def test_axis_range_takes_effect_on_reloaded_graph(tmp_path):
    """Setting an axis range on a reloaded graph must succeed (the read-back
    verification would raise if it silently no-oped)."""
    from origin_pro_mcp.tools.graph import axis
    from origin_pro_mcp.tools.project import load_project, new_project, save_project

    new_project()
    name = _build_multiseries()
    proj = str(tmp_path / "roundtrip_axis.opju")
    save_project(proj)

    new_project()
    load_project(proj)

    # Would raise "did not take effect" if the loaded layer froze the change.
    msg = axis(name, op="range", axis="y", range_min=0, range_max=100)
    assert "Set axis range" in msg


def test_style_reloaded_multiseries_via_ungroup_then_color(tmp_path):
    """End-to-end: ungroup + per-curve color on a reloaded multi-series graph."""
    from origin_pro_mcp.tools.project import load_project, new_project, save_project
    from origin_pro_mcp.tools.style import set_plot_style, ungroup_plots

    new_project()
    name = _build_multiseries()
    proj = str(tmp_path / "roundtrip_ungroup.opju")
    save_project(proj)

    new_project()
    load_project(proj)

    ungroup_plots(name)  # must find the plots on the reloaded graph, not raise
    set_plot_style(name, plot_index=1, rgb="0,0,255")
    set_plot_style(name, plot_index=2, rgb="255,128,0")
