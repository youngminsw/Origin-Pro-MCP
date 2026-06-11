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
