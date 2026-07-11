"""Guard tests for analysis tools (no Origin needed).

These drive the consolidated ``transform`` and ``stats`` dispatchers; the
underlying private impls are exercised directly only where a sub-option
(e.g. the smoothing sub-method) is not reachable through the dispatcher.
"""
import json

import pytest


def test_integrate_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    with pytest.raises(ValueError, match="not found"):
        transform("Ghost", "Sheet1", 1, 2, method="integrate")


def test_integrate_returns_area(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    fake_origin.lt_vars["integ1.area"] = 42.0
    out = json.loads(transform("Book1", "Sheet1", 1, 2, method="integrate"))
    assert out["area"] == 42.0


def test_smooth_rejects_bad_method(fake_origin):
    # smooth_method is now reachable through the transform dispatcher (item 9).
    from origin_pro_mcp.tools.analysis import transform

    with pytest.raises(ValueError, match="must be one of"):
        transform("Book1", "Sheet1", 1, 2, method="smooth", smooth_method="bilateral")


def test_interpolate_reuses_stable_book_when_present(fake_origin):
    # Item 11: a pre-existing "Interp" book is reused (rows cleared), not a new
    # Interp1/Interp2... spawned.
    from fakes import FakeBook
    from origin_pro_mcp.tools.analysis import transform

    fake_origin.books.append(FakeBook("Interp"))
    transform("Book1", "Sheet1", 1, 2, method="interpolate", num_points=10)
    joined = " ".join(fake_origin.executed)
    assert "wks.nrows = 0" in joined  # the reuse/overwrite path


def test_interpolate_creates_book_when_absent(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    transform("Book1", "Sheet1", 1, 2, method="interpolate", num_points=10)
    joined = " ".join(fake_origin.executed)
    assert "wks.nrows = 0" not in joined  # first-time create path, no clear


def test_find_peaks_deletes_prior_output_columns(fake_origin, monkeypatch):
    # Item 11: on a repeat call, find_peaks deletes the "Peak X"/"Peak Y"
    # columns it left last time (highest index first) rather than appending
    # two more.
    from origin_pro_mcp.tools.analysis import transform

    def fake_ltstr(name):
        if name == "wks.col3.lname$":
            return "Peak X"
        if name == "wks.col4.lname$":
            return "Peak Y"
        return ""

    monkeypatch.setattr(fake_origin, "LTStr", fake_ltstr)
    fake_origin.lt_vars["wks.ncols"] = 4
    fake_origin.worksheet_data = tuple((float(i), float(i)) for i in range(1, 12))
    transform("Book1", "Sheet1", 1, 2, method="find_peaks")
    joined = " ".join(fake_origin.executed)
    assert "delete col(4)" in joined
    assert "delete col(3)" in joined
    # delete the higher index before the lower so the second isn't shifted
    assert joined.index("delete col(4)") < joined.index("delete col(3)")


def test_find_peaks_no_prior_columns_no_delete(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    fake_origin.lt_vars["wks.ncols"] = 2
    fake_origin.worksheet_data = tuple((float(i), float(i)) for i in range(1, 12))
    transform("Book1", "Sheet1", 1, 2, method="find_peaks")
    assert "delete col" not in " ".join(fake_origin.executed)


def test_find_peaks_rejects_bad_direction(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    with pytest.raises(ValueError, match="direction must be one of"):
        transform("Book1", "Sheet1", 1, 2, method="find_peaks", direction="up")


def test_find_peaks_clamps_local_points_to_data_length(fake_origin):
    """Item 28: local_points is clamped to (n-1)//2 so a default 10 doesn't
    fail on a short spectrum; the used value is reported and the pkfind command
    carries it."""
    from origin_pro_mcp.tools.analysis import transform

    # 11 rows (Y = col 2). max window = (11-2)//2 = 4, so 10 clamps to 4.
    fake_origin.worksheet_data = tuple((float(i), float(i)) for i in range(11))
    out = json.loads(
        transform("Book1", "Sheet1", 1, 2, method="find_peaks", local_points=10)
    )
    assert out["local_points_used"] == 4
    assert any("npts:=4" in s for s in fake_origin.executed)


def test_find_peaks_raises_actionably_on_too_few_points(fake_origin):
    """Item 28: fewer than 3 points gives an actionable error, not pkfind's
    opaque failure."""
    from origin_pro_mcp.tools.analysis import transform

    fake_origin.worksheet_data = ((1.0, 2.0), (2.0, 4.0))  # only 2 rows
    with pytest.raises(ValueError, match="at least 3 numeric points"):
        transform("Book1", "Sheet1", 1, 2, method="find_peaks")


def test_interpolate_rejects_bad_method(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    with pytest.raises(ValueError, match="method must be one of"):
        transform("Book1", "Sheet1", 1, 2, method="interpolate", interp_method="quadratic")


def test_differentiate_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    with pytest.raises(ValueError, match="not found"):
        transform("Ghost", "Sheet1", 1, 2, method="differentiate")


def test_column_statistics_computes_se_and_variance(fake_origin):
    from origin_pro_mcp.tools.analysis import stats

    fake_origin.lt_vars.update({"stats.mean": 10.0, "stats.sd": 2.0, "stats.n": 4.0,
                                "stats.min": 6.0, "stats.max": 14.0, "stats.sum": 40.0,
                                "__mcp_med": 9.5})
    out = json.loads(stats("Book1", "Sheet1", op="column", col=1))
    assert out["mean"] == 10.0 and out["variance"] == 4.0
    assert out["se"] == 1.0  # sd / sqrt(n) = 2 / 2
    assert out["median"] == 9.5


def test_compare_means_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.analysis import stats

    with pytest.raises(ValueError, match="not found"):
        stats("Ghost", "Sheet1", op="compare_means", col=1, col2=2)


def test_compare_means_requires_col2(fake_origin):
    from origin_pro_mcp.tools.analysis import stats

    with pytest.raises(ValueError, match="requires col2"):
        stats("Book1", "Sheet1", op="compare_means", col=1)


def test_frequency_count_rejects_bad_bins(fake_origin):
    from origin_pro_mcp.tools.analysis import stats

    with pytest.raises(ValueError, match="bin_size must be positive"):
        stats("Book1", "Sheet1", op="frequency", col=1, bin_min=0, bin_max=10, bin_size=0)
    with pytest.raises(ValueError, match="bin_max must be greater"):
        stats("Book1", "Sheet1", op="frequency", col=1, bin_min=10, bin_max=5, bin_size=1)


def test_frequency_count_requires_bins(fake_origin):
    from origin_pro_mcp.tools.analysis import stats

    with pytest.raises(ValueError, match="requires bin_min"):
        stats("Book1", "Sheet1", op="frequency", col=1)
