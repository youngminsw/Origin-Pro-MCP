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
    # The smoothing sub-method is reached only through the private impl.
    from origin_pro_mcp.tools.analysis import _smooth_impl

    with pytest.raises(ValueError, match="method must be one of"):
        _smooth_impl("Book1", "Sheet1", 1, 2, method="bilateral")


def test_find_peaks_rejects_bad_direction(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    with pytest.raises(ValueError, match="direction must be one of"):
        transform("Book1", "Sheet1", 1, 2, method="find_peaks", direction="up")


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
