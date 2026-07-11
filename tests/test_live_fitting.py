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


# --- multi-peak deconvolution (item 7) ---------------------------------------

def _three_gaussian_spectrum(book="MP3G"):
    """y0=0; three well-separated Gaussians at xc=20/35/50 (sigma 4/3/5,
    amp 10/15/8); x 0..70 step 0.35 = 201 pts + tiny noise. Origin's `gauss`
    reports w=2*sigma and A=area, so the fit recovers xc dead-on."""
    import math
    import random

    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet(book))
    b, sh = made["name"], made["sheet"]
    random.seed(7)
    xs = [round(i * 0.35, 4) for i in range(0, 201)]

    def g(x, xc, s, A):
        return A * math.exp(-((x - xc) ** 2) / (2 * s * s))

    ys = [
        g(x, 20, 4, 10) + g(x, 35, 3, 15) + g(x, 50, 5, 8) + random.uniform(-0.05, 0.05)
        for x in xs
    ]
    set_worksheet_data(b, sh, json.dumps([xs, ys]))
    return b, sh


def test_multipeak_three_gaussians_recovers_centers(live_origin):
    b, sh = _three_gaussian_spectrum()
    from origin_pro_mcp.tools.fitting import curve_fit

    out = json.loads(curve_fit(b, sh, 1, 2, function="gauss", peaks=3))
    assert out["peaks"] == 3
    xcs = sorted(out["parameters"][f"peak_{k}"]["xc"]["value"] for k in (1, 2, 3))
    for got, truth in zip(xcs, (20.0, 35.0, 50.0)):
        assert abs(got - truth) < 0.5, f"recovered xc={xcs} vs 20/35/50"
    assert out["statistics"]["r_squared"] > 0.99, out["statistics"]
    # every peak reports a width, an area and their std errors
    for k in (1, 2, 3):
        blk = out["parameters"][f"peak_{k}"]
        assert "value" in blk["w"] and "value" in blk["A"]
        assert "std_error" in blk["xc"]


def test_multipeak_lorentz_with_supplied_centers(live_origin):
    """Two Lorentzians at xc=20/40 with caller-supplied initial centres."""
    import math

    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("MP2L"))
    b, sh = made["name"], made["sheet"]
    xs = [round(i * 0.35, 4) for i in range(0, 201)]

    def lz(x, xc, w, A):
        return (2 * A / math.pi) * (w / (4 * (x - xc) ** 2 + w * w))

    ys = [lz(x, 20, 4, 40) + lz(x, 40, 5, 60) for x in xs]
    set_worksheet_data(b, sh, json.dumps([xs, ys]))

    from origin_pro_mcp.tools.fitting import curve_fit

    out = json.loads(
        curve_fit(b, sh, 1, 2, function="lorentz", peaks=2, peak_centers="20,40")
    )
    xcs = sorted(out["parameters"][f"peak_{k}"]["xc"]["value"] for k in (1, 2))
    assert abs(xcs[0] - 20.0) < 0.5 and abs(xcs[1] - 40.0) < 0.5, xcs


def test_multipeak_composes_with_x_range(live_origin):
    """Restrict to [10,42] (drops the xc=50 peak), then fit the two remaining
    peaks — deconvolution within a sub-range, the real XPS workflow."""
    b, sh = _three_gaussian_spectrum(book="MP3R")
    from origin_pro_mcp.tools.fitting import curve_fit

    out = json.loads(
        curve_fit(b, sh, 1, 2, function="gauss", peaks=2, x_min=10.0, x_max=42.0)
    )
    xcs = sorted(out["parameters"][f"peak_{k}"]["xc"]["value"] for k in (1, 2))
    assert abs(xcs[0] - 20.0) < 0.5 and abs(xcs[1] - 35.0) < 0.5, xcs


def test_multipeak_nonconvergence_raises(live_origin):
    """Flat data + centres far outside the range cannot converge; the R^2<=0
    guard must raise actionably rather than return a frozen fit."""
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("MPBAD"))
    b, sh = made["name"], made["sheet"]
    xs = [float(i) for i in range(0, 50)]
    ys = [1.0 for _ in xs]
    set_worksheet_data(b, sh, json.dumps([xs, ys]))

    from origin_pro_mcp.tools.fitting import curve_fit

    with pytest.raises(ValueError, match="did not converge"):
        curve_fit(
            b, sh, 1, 2, function="gauss", peaks=3, peak_centers="1000,2000,3000"
        )


def test_multipeak_plot_draws_cumulative_and_components(live_origin):
    b, sh = _three_gaussian_spectrum(book="MP3P")
    from origin_pro_mcp.tools.fitting import curve_fit
    from origin_pro_mcp.tools.graph import create_graph, export_graph

    g = json.loads(create_graph("MPG", b, sh, x_col=1, y_col=2, plot_type="line"))
    graph = g["name"]
    out = json.loads(
        curve_fit(b, sh, 1, 2, function="gauss", peaks=3, plot_on_graph=graph)
    )
    assert "fit_curve" in out, out
    assert out["fit_curve"]["components_drawn"] == 3, out["fit_curve"]
    assert out["fit_curve"]["cumulative_column"] > 0

    import os
    import tempfile

    path = os.path.join(tempfile.gettempdir(), "mp_fit_plot.png")
    export_graph(graph, path)
    assert os.path.exists(path) and os.path.getsize(path) > 1000
