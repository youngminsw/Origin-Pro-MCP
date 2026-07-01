"""D3 usage guard: typed plot/delete tools must REJECT a cross-kind window
name (a graph passed where a worksheet is expected, or vice versa) with a
clear ValueError BEFORE any window-scoped LabTalk command is dispatched.

This is the hang-trigger-surface reduction from the downgrade plan: the
reported hang was a worksheet-scoped command (`worksheet -p 201`) reaching a
GRAPH window. The require_worksheet / require_graph guards (COM property
lookups, not synchronous LabTalk) turn that into an immediate friendly error
instead of a command that can wedge the COM apartment. These tests lock the
invariant in COM-free via the shared FakeOrigin (Book1 workbook, Graph1 graph).
"""
import pytest



def _no_plot_dispatched(fake):
    """No plot/worksheet-scoped LabTalk command reached Origin."""
    for script in fake.executed:
        assert not script.startswith("plotxy"), f"plot dispatched: {script!r}"
        assert not script.startswith("plotxyz"), f"plot dispatched: {script!r}"
        assert "worksheet -p" not in script, f"worksheet-plot dispatched: {script!r}"


def test_create_graph_rejects_graph_name_as_source(fake_origin):
    from origin_pro_mcp.tools.graph import create_graph
    # data_book "Graph1" is a GRAPH, not a worksheet
    with pytest.raises(ValueError, match="not found"):
        create_graph("G", "Graph1", "Layer1", 1, 2)
    _no_plot_dispatched(fake_origin)


def test_add_plot_rejects_worksheet_as_target_graph(fake_origin):
    from origin_pro_mcp.tools.graph import add_plot_to_graph
    # graph_name "Book1" is a WORKBOOK, not a graph
    with pytest.raises(ValueError, match="not found"):
        add_plot_to_graph("Book1", "Book1", "Sheet1", 1, 2)
    _no_plot_dispatched(fake_origin)


def test_add_plot_rejects_graph_name_as_source_book(fake_origin):
    from origin_pro_mcp.tools.graph import add_plot_to_graph
    # target graph is valid (Graph1) but data_book "Graph1" is not a worksheet
    with pytest.raises(ValueError, match="not found"):
        add_plot_to_graph("Graph1", "Graph1", "Layer1", 1, 2)
    _no_plot_dispatched(fake_origin)


def test_delete_graph_rejects_workbook_name(fake_origin):
    from origin_pro_mcp.tools.graph import delete_graph
    # deleting "Book1" (a workbook) via the graph tool must not fire win -cd
    with pytest.raises(ValueError, match="not found"):
        delete_graph("Book1")
    for script in fake_origin.executed:
        assert "win -cd" not in script, f"destructive win -cd dispatched: {script!r}"


def test_remove_plot_rejects_workbook_name(fake_origin):
    from origin_pro_mcp.tools.graph import remove_plot
    with pytest.raises(ValueError, match="not found"):
        remove_plot("Book1", 1)
    for script in fake_origin.executed:
        assert "layer -e" not in script
        assert "layer -d" not in script


def test_valid_create_graph_still_dispatches_plot(fake_origin):
    """Positive control: a correct worksheet source DOES reach plotxy."""
    from origin_pro_mcp.tools.graph import create_graph
    out = create_graph("G", "Book1", "Sheet1", 1, 2)
    assert "Created graph" in out
    assert any(s.startswith("plotxy") for s in fake_origin.executed)
