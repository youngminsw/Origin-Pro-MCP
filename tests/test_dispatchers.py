"""Behavior-preservation tests for the Phase-2 consolidated dispatchers.

The 27 old tools were turned into unchanged private impls (``_<name>_impl``)
and replaced by 8 dispatchers. These tests assert, for EVERY op/method/kind
branch, that the dispatcher issues exactly the same sequence of LabTalk
commands (``fake.executed``) and returns the same value as calling the original
impl directly — proving the refactor preserved behavior. They also lock the
registry at 37 tools and cover the bad-selector and missing-required-param
errors.
"""
from types import SimpleNamespace

import pytest

from origin_pro_mcp.tools import analysis as A
from origin_pro_mcp.tools import graph as G
from origin_pro_mcp.tools import worksheet as W
from origin_pro_mcp.tools import style as S


# --- equivalence harness ------------------------------------------------------

def _run(fake, fn, *args, **kwargs):
    fake.executed.clear()
    err = None
    result = None
    try:
        result = fn(*args, **kwargs)
    except Exception as exc:  # capture type+message for comparison
        err = (type(exc).__name__, str(exc))
    return result, err, list(fake.executed)


def _assert_equiv(fake, impl, impl_a, disp, disp_a, impl_kw=None, disp_kw=None):
    r1, e1, x1 = _run(fake, impl, *impl_a, **(impl_kw or {}))
    r2, e2, x2 = _run(fake, disp, *disp_a, **(disp_kw or {}))
    assert x1 == x2, (x1, x2)
    assert e1 == e2, (e1, e2)
    assert r1 == r2, (r1, r2)


# --- registry shape -----------------------------------------------------------

NEW_TOOLS = {
    "axis", "transform", "stats", "annotate", "colormap",
    "export_graph", "manage_columns", "import_data",
}
REMOVED_TOOLS = {
    "set_axis_labels", "set_axis_range", "set_axis_scale", "set_tick_style",
    "integrate", "differentiate", "smooth", "interpolate", "fft", "find_peaks",
    "column_statistics", "compare_means", "frequency_count",
    "add_reference_line", "add_text_annotation", "add_line", "add_arrow",
    "apply_color_map", "set_colormap_levels", "export_graph_sized",
    "add_columns", "delete_columns", "set_column_properties",
    "set_column_formula", "import_csv_to_worksheet", "import_excel",
}


def _registry():
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp
    return mcp._tool_manager._tools


def test_registry_has_exactly_48_tools():
    # 43 Origin tools + list_skills + get_skill = 45, plus the N7 additions
    # set_plot_color, set_plot_line_width, ungroup_plots = 48. (Skills are
    # exposed as first-class MCP tools via origin_pro_mcp.skills.register_skills.)
    assert len(_registry()) == 48


def test_eight_new_dispatchers_registered():
    names = set(_registry())
    assert NEW_TOOLS <= names


def test_twenty_six_old_names_deregistered():
    names = set(_registry())
    assert not (REMOVED_TOOLS & names)


# --- transform: 6 method branches --------------------------------------------

def test_transform_integrate_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._integrate_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  disp_kw={"method": "integrate"})


def test_transform_differentiate_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._differentiate_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  disp_kw={"method": "differentiate"})


def test_transform_smooth_default_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._smooth_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  disp_kw={"method": "smooth"})


def test_transform_smooth_window_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._smooth_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  impl_kw={"window": 7},
                  disp_kw={"method": "smooth", "window_size": 7})


def test_transform_interpolate_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._interpolate_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  impl_kw={"num_points": 50, "method": "spline"},
                  disp_kw={"method": "interpolate", "num_points": 50,
                           "interp_method": "spline"})


def test_transform_fft_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._fft_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  disp_kw={"method": "fft"})


def test_transform_find_peaks_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._find_peaks_impl, ("Book1", "Sheet1", 1, 2),
                  A.transform, ("Book1", "Sheet1", 1, 2),
                  impl_kw={"direction": "both", "local_points": 6},
                  disp_kw={"method": "find_peaks", "direction": "both",
                           "local_points": 6})


def test_transform_bad_method(fake_origin):
    with pytest.raises(ValueError, match="method must be one of"):
        A.transform("Book1", "Sheet1", 1, 2, method="wavelet")


# --- stats: 3 op branches -----------------------------------------------------

def test_stats_column_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._column_statistics_impl, ("Book1", "Sheet1", 1),
                  A.stats, ("Book1", "Sheet1", "column", 1))


def test_stats_compare_means_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._compare_means_impl, ("Book1", "Sheet1", 1, 2),
                  A.stats, ("Book1", "Sheet1", "compare_means", 1),
                  impl_kw={"equal_variance": True},
                  disp_kw={"col2": 2, "equal_variance": True})


def test_stats_frequency_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  A._frequency_count_impl, ("Book1", "Sheet1", 1, 0.0, 10.0, 1.0),
                  A.stats, ("Book1", "Sheet1", "frequency", 1),
                  disp_kw={"bin_min": 0.0, "bin_max": 10.0, "bin_size": 1.0})


def test_stats_bad_op(fake_origin):
    with pytest.raises(ValueError, match="op must be one of"):
        A.stats("Book1", "Sheet1", "median", 1)


def test_stats_compare_means_missing_col2(fake_origin):
    with pytest.raises(ValueError, match="requires col2"):
        A.stats("Book1", "Sheet1", "compare_means", 1)


def test_stats_frequency_missing_bins(fake_origin):
    with pytest.raises(ValueError, match="requires bin_min"):
        A.stats("Book1", "Sheet1", "frequency", 1)


# --- axis: 4 op branches ------------------------------------------------------

def test_axis_labels_x_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._set_axis_labels_impl, ("Graph1",),
                  G.axis, ("Graph1", "labels"),
                  impl_kw={"x_label": "Temperature"},
                  disp_kw={"axis": "x", "label": "Temperature"})


def test_axis_labels_y_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._set_axis_labels_impl, ("Graph1",),
                  G.axis, ("Graph1", "labels"),
                  impl_kw={"y_label": "Signal"},
                  disp_kw={"axis": "y", "label": "Signal"})


def test_axis_range_x_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._set_axis_range_impl, ("Graph1",),
                  G.axis, ("Graph1", "range"),
                  impl_kw={"x_min": 0.0, "x_max": 10.0},
                  disp_kw={"axis": "x", "range_min": 0.0, "range_max": 10.0})


def test_axis_scale_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._set_axis_scale_impl, ("Graph1",),
                  G.axis, ("Graph1", "scale"),
                  impl_kw={"axis": "y", "scale": "log10"},
                  disp_kw={"axis": "y", "scale": "log10"})


def test_axis_scale_default_axis_defaults_to_y(fake_origin):
    """MEDIUM 3: axis(op="scale") with NO axis arg must default to y (the old
    set_axis_scale default) and issue the SAME command — not forward
    axis="both", which the impl rejects."""
    _assert_equiv(fake_origin,
                  G._set_axis_scale_impl, ("Graph1",),
                  G.axis, ("Graph1", "scale"),
                  impl_kw={"scale": "log10"},   # impl default axis="y"
                  disp_kw={"scale": "log10"})   # dispatcher: no axis -> must be y


def test_axis_tick_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  S._set_tick_style_impl, ("Graph1",),
                  G.axis, ("Graph1", "tick"),
                  impl_kw={"tick_direction": "out", "major_length": 6,
                           "minor_count": 2, "show_minor": False},
                  disp_kw={"tick_direction": "out", "major_length": 6,
                           "minor_count": 2, "show_minor": False})


def test_axis_tick_default_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  S._set_tick_style_impl, ("Graph1",),
                  G.axis, ("Graph1", "tick"))


def test_axis_bad_op(fake_origin):
    with pytest.raises(ValueError, match="op must be one of"):
        G.axis("Graph1", "grid")


def test_axis_labels_missing_label(fake_origin):
    with pytest.raises(ValueError, match="op 'labels' requires label"):
        G.axis("Graph1", "labels")


def test_axis_scale_missing_scale(fake_origin):
    with pytest.raises(ValueError, match="op 'scale' requires scale"):
        G.axis("Graph1", "scale", axis="y")


# --- annotate: 4 kind branches ------------------------------------------------

def test_annotate_reference_line_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._add_reference_line_impl, ("Graph1", "horizontal", 5.0),
                  G.annotate, ("Graph1", "reference_line"),
                  disp_kw={"orientation": "horizontal", "value": 5.0})


def test_annotate_text_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._add_text_annotation_impl, ("Graph1", "Peak", 3.0, 70.0),
                  G.annotate, ("Graph1", "text"),
                  disp_kw={"text": "Peak", "x1": 3.0, "y1": 70.0})


def test_annotate_line_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._add_line_impl, ("Graph1", 1.0, 2.0, 3.0, 4.0),
                  G.annotate, ("Graph1", "line"),
                  disp_kw={"x1": 1.0, "y1": 2.0, "x2": 3.0, "y2": 4.0})


def test_annotate_arrow_equiv(fake_origin, monkeypatch):
    # add_arrow embeds a random uuid in the object name; pin it so the two
    # calls produce byte-identical LabTalk.
    monkeypatch.setattr(G.uuid, "uuid4", lambda: SimpleNamespace(hex="deadbeef00"))
    _assert_equiv(fake_origin,
                  G._add_arrow_impl, ("Graph1", 1.0, 2.0, 3.0, 4.0),
                  G.annotate, ("Graph1", "arrow"),
                  impl_kw={"double_headed": True, "head_size": 12},
                  disp_kw={"x1": 1.0, "y1": 2.0, "x2": 3.0, "y2": 4.0,
                           "double_headed": True, "head_size": 12})


def test_annotate_bad_kind(fake_origin):
    with pytest.raises(ValueError, match="kind must be one of"):
        G.annotate("Graph1", "circle")


def test_annotate_reference_line_missing_params(fake_origin):
    with pytest.raises(ValueError, match="requires orientation and value"):
        G.annotate("Graph1", "reference_line", orientation="horizontal")


def test_annotate_line_missing_params(fake_origin):
    with pytest.raises(ValueError, match="kind 'line' requires"):
        G.annotate("Graph1", "line", x1=1, y1=2)


# --- colormap: 2 branches -----------------------------------------------------

def test_colormap_palette_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._apply_color_map_impl, ("Graph1", "Fire"),
                  G.colormap, ("Graph1",),
                  disp_kw={"palette": "Fire"})


def test_colormap_levels_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  G._set_colormap_levels_impl, ("Graph1", 0.0, 10.0),
                  G.colormap, ("Graph1",),
                  disp_kw={"z_min": 0.0, "z_max": 10.0})


def test_colormap_requires_argument(fake_origin):
    with pytest.raises(ValueError, match="requires palette"):
        G.colormap("Graph1")


def test_colormap_requires_both_bounds(fake_origin):
    with pytest.raises(ValueError, match="both z_min and z_max"):
        G.colormap("Graph1", z_max=10)


# --- export_graph: 2 branches -------------------------------------------------

def test_export_graph_plain_equiv(fake_origin, tmp_path):
    # The fake can't produce a real file, so the file-existence check fails;
    # the unknown-graph guard path exercises the same code in both impl and
    # dispatcher (sized=False).
    p = str(tmp_path / "fig.png")
    _assert_equiv(fake_origin,
                  G._export_graph_impl, ("Ghost", p),
                  G.export_graph, ("Ghost", p))


def test_export_graph_default_is_clipboard_free(fake_origin, tmp_path, monkeypatch):
    # The default export (sized=False) must issue an expGraph LabTalk command
    # (clipboard-free) and never touch CopyPage. Pretend the file was written
    # so the impl's existence/size checks pass under the fake.
    monkeypatch.setattr(G.os.path, "exists", lambda _p: True)
    monkeypatch.setattr(G.os.path, "getsize", lambda _p: 12345)
    p = str(tmp_path / "fig.png")
    fake_origin.executed.clear()
    msg = G.export_graph("Graph1", p)
    cmds = " ".join(fake_origin.executed)
    assert "expGraph" in cmds
    assert "tr1.width" in cmds
    assert "CopyPage" not in cmds
    assert "Exported to:" in msg


def test_export_graph_sized_equiv(fake_origin, tmp_path):
    p = str(tmp_path / "fig.png")
    _assert_equiv(fake_origin,
                  G._export_graph_sized_impl, ("Graph1", p),
                  G.export_graph, ("Graph1", p),
                  impl_kw={"width": 1200, "height": 0, "format": "png"},
                  disp_kw={"sized": True, "width": 1200, "height": 0,
                           "format": "png"})


# --- manage_columns: 4 op branches -------------------------------------------

def test_manage_columns_add_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  W._add_columns_impl, ("Book1", "Sheet1"),
                  W.manage_columns, ("Book1", "Sheet1", "add"),
                  impl_kw={"count": 3}, disp_kw={"count": 3})


def test_manage_columns_delete_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  W._delete_columns_impl, ("Book1", "Sheet1", 2),
                  W.manage_columns, ("Book1", "Sheet1", "delete"),
                  impl_kw={"count": 2}, disp_kw={"col": 2, "count": 2})


def test_manage_columns_properties_equiv(fake_origin):
    _assert_equiv(fake_origin,
                  W._set_column_properties_impl, ("Book1", "Sheet1", 1),
                  W.manage_columns, ("Book1", "Sheet1", "properties"),
                  impl_kw={"long_name": "Time", "units": "s"},
                  disp_kw={"col": 1, "long_name": "Time", "units": "s"})


def test_manage_columns_formula_equiv(fake_origin):
    fake_origin.lt_vars["wks.ncols"] = 5  # avoid the grow-loop
    _assert_equiv(fake_origin,
                  W._set_column_formula_impl, ("Book1", "Sheet1", 2, "col(1)^2"),
                  W.manage_columns, ("Book1", "Sheet1", "formula"),
                  disp_kw={"col": 2, "formula": "col(1)^2"})


def test_manage_columns_bad_op(fake_origin):
    with pytest.raises(ValueError, match="op must be one of"):
        W.manage_columns("Book1", "Sheet1", "rename")


def test_manage_columns_delete_missing_col(fake_origin):
    with pytest.raises(ValueError, match="op 'delete' requires col"):
        W.manage_columns("Book1", "Sheet1", "delete")


# --- import_data: 2 branches --------------------------------------------------

def test_import_data_csv_equiv(fake_origin, tmp_path):
    f = tmp_path / "data.csv"
    f.write_text("1,2\n3,4\n")
    p = str(f)
    _assert_equiv(fake_origin,
                  W._import_csv_to_worksheet_impl, (p,),
                  W.import_data, (p,),
                  disp_kw={"format": "csv"})


def test_import_data_excel_equiv(fake_origin, tmp_path):
    f = tmp_path / "data.xlsx"
    f.write_bytes(b"PK\x03\x04stub")
    p = str(f)
    _assert_equiv(fake_origin,
                  W._import_excel_impl, (p,),
                  W.import_data, (p,),
                  disp_kw={"format": "excel"})


def test_import_data_auto_detects_excel(fake_origin, tmp_path):
    f = tmp_path / "data.xlsx"
    f.write_bytes(b"PK\x03\x04stub")
    p = str(f)
    # format="auto" must route .xlsx through the Excel impl (impExcel command).
    _assert_equiv(fake_origin,
                  W._import_excel_impl, (p,),
                  W.import_data, (p,))


def test_import_data_auto_detects_csv(fake_origin, tmp_path):
    f = tmp_path / "data.csv"
    f.write_text("1,2\n")
    p = str(f)
    _assert_equiv(fake_origin,
                  W._import_csv_to_worksheet_impl, (p,),
                  W.import_data, (p,))


def test_import_data_bad_format(fake_origin, tmp_path):
    with pytest.raises(ValueError, match="format must be one of"):
        W.import_data(str(tmp_path / "x.csv"), format="parquet")
