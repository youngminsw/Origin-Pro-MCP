"""Live (Windows + Origin Pro COM) tests for curve_fit's X-range restriction
(item 8).

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_fitting.py -v

Safety: an isolated ``DispatchEx("Origin.Application")`` per test; never
``Origin.ApplicationSI``; ``origin.Exit()`` in teardown.
"""
import json
import math

import pytest

pytestmark = pytest.mark.requires_origin


@pytest.fixture()
def live_origin():
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


def _two_peak_spectrum():
    """X 0..20 step .5; a small peak at xc=5 and a LARGER interfering peak at
    xc=15. A single-Gaussian fit of the whole curve is pulled to the big peak
    (~15); restricting to [2,8] must recover the small in-range peak (~5)."""
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("FITSPEC"))
    b, sh = made["name"], made["sheet"]
    xs = [i * 0.5 for i in range(0, 41)]

    def g(x, xc, A, w=1.0):
        return A * math.exp(-((x - xc) ** 2) / (2 * w * w))

    ys = [g(x, 5.0, 10.0) + g(x, 15.0, 30.0) for x in xs]
    set_worksheet_data(b, sh, json.dumps([xs, ys]))
    return b, sh


def test_curve_fit_x_range_isolates_in_range_peak(live_origin):
    b, sh = _two_peak_spectrum()
    from origin_pro_mcp.tools.fitting import curve_fit

    full = json.loads(curve_fit(b, sh, 1, 2, function="gauss"))
    restricted = json.loads(
        curve_fit(b, sh, 1, 2, function="gauss", x_min=2.0, x_max=8.0)
    )

    xc_full = full["parameters"]["xc"]["value"]
    xc_restricted = restricted["parameters"]["xc"]["value"]

    # The restricted fit must land on the in-range peak at x=5.
    assert abs(xc_restricted - 5.0) < 0.5, f"restricted xc={xc_restricted}"
    # And the restriction must actually change the outcome (full fit is pulled
    # toward the bigger out-of-range peak near 15).
    assert xc_full > 10.0, f"full xc={xc_full} (expected ~15)"


def test_curve_fit_x_range_no_points_raises(live_origin):
    b, sh = _two_peak_spectrum()
    from origin_pro_mcp.tools.fitting import curve_fit

    with pytest.raises(ValueError):
        curve_fit(b, sh, 1, 2, function="gauss", x_min=100.0, x_max=200.0)
