"""Guard tests for analysis tools (no Origin needed)."""
import json

import pytest


def test_integrate_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.analysis import integrate

    with pytest.raises(ValueError, match="not found"):
        integrate("Ghost", "Sheet1", 1, 2)


def test_integrate_returns_area(fake_origin):
    from origin_pro_mcp.tools.analysis import integrate

    fake_origin.lt_vars["integ1.area"] = 42.0
    out = json.loads(integrate("Book1", "Sheet1", 1, 2))
    assert out["area"] == 42.0


def test_smooth_rejects_bad_method(fake_origin):
    from origin_pro_mcp.tools.analysis import smooth

    with pytest.raises(ValueError, match="method must be one of"):
        smooth("Book1", "Sheet1", 1, 2, method="bilateral")


def test_find_peaks_rejects_bad_direction(fake_origin):
    from origin_pro_mcp.tools.analysis import find_peaks

    with pytest.raises(ValueError, match="direction must be one of"):
        find_peaks("Book1", "Sheet1", 1, 2, direction="up")


def test_interpolate_rejects_bad_method(fake_origin):
    from origin_pro_mcp.tools.analysis import interpolate

    with pytest.raises(ValueError, match="method must be one of"):
        interpolate("Book1", "Sheet1", 1, 2, method="quadratic")


def test_differentiate_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.analysis import differentiate

    with pytest.raises(ValueError, match="not found"):
        differentiate("Ghost", "Sheet1", 1, 2)


def test_column_statistics_computes_se_and_variance(fake_origin):
    from origin_pro_mcp.tools.analysis import column_statistics

    fake_origin.lt_vars.update({"stats.mean": 10.0, "stats.sd": 2.0, "stats.n": 4.0,
                                "stats.min": 6.0, "stats.max": 14.0, "stats.sum": 40.0,
                                "__mcp_med": 9.5})
    out = json.loads(column_statistics("Book1", "Sheet1", 1))
    assert out["mean"] == 10.0 and out["variance"] == 4.0
    assert out["se"] == 1.0  # sd / sqrt(n) = 2 / 2
    assert out["median"] == 9.5


def test_compare_means_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.analysis import compare_means

    with pytest.raises(ValueError, match="not found"):
        compare_means("Ghost", "Sheet1", 1, 2)


def test_frequency_count_rejects_bad_bins(fake_origin):
    from origin_pro_mcp.tools.analysis import frequency_count

    with pytest.raises(ValueError, match="bin_size must be positive"):
        frequency_count("Book1", "Sheet1", 1, 0, 10, 0)
    with pytest.raises(ValueError, match="bin_max must be greater"):
        frequency_count("Book1", "Sheet1", 1, 10, 5, 1)
