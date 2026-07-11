"""Live (Windows + Origin Pro COM) tests for the dual-Y polish round:
right-axis title typed path (item 29) and legend re-placement after
add_second_y_axis (item 30).

Run on the Windows side with:
    pytest -m requires_origin tests/test_live_dualy.py -v

Safety: an isolated ``DispatchEx("Origin.Application")`` per test; never
``Origin.ApplicationSI``; ``origin.Exit()`` in teardown.
"""
import json

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


def _dual_y_graph(name="DUALG"):
    from origin_pro_mcp.tools.graph import add_second_y_axis, create_graph
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("DUALB"))
    b, sh = made["name"], made["sheet"]
    set_worksheet_data(
        b, sh,
        json.dumps([[1, 2, 3, 4, 5], [2, 4, 6, 8, 10], [50, 40, 30, 20, 10]]),
    )
    g = json.loads(create_graph(name, b, sh, 1, 2, plot_type="line+symbol"))["name"]
    msg = add_second_y_axis(g, b, sh, 1, 3, plot_type="line+symbol")
    return g, b, sh, msg


def test_right_axis_title_typed_path_sets_and_reads_back(live_origin):
    """Item 29: axis(op='labels', axis='right') sets the yr title on layer 2
    and it reads back."""
    from origin_pro_mcp.origin_connection import execute_labtalk, get_lt_str
    from origin_pro_mcp.tools.graph import axis

    g, _, _, _ = _dual_y_graph()
    axis(g, "labels", axis="right", label="Magnetization")
    execute_labtalk(f"win -a {g}; page.active=2;")
    got = get_lt_str("yr.text$")
    assert "Magnetization" in got, f"yr.text$={got!r}"


def test_right_axis_title_raises_without_second_layer(live_origin):
    """Item 29: on a single-layer graph, axis='right' must raise (the false
    success the usability agent hit), not silently claim it worked."""
    from origin_pro_mcp.tools.graph import create_graph, axis
    from origin_pro_mcp.tools.worksheet import create_worksheet, set_worksheet_data

    made = json.loads(create_worksheet("SINGLEB"))
    b, sh = made["name"], made["sheet"]
    set_worksheet_data(b, sh, json.dumps([[1, 2, 3], [4, 5, 6]]))
    g = json.loads(create_graph("SINGLEG", b, sh, 1, 2, plot_type="line+symbol"))["name"]
    with pytest.raises(ValueError):
        axis(g, "labels", axis="right", label="Should Not Apply")
