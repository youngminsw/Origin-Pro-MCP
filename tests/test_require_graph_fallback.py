"""require_graph's GraphPages fallback (origin2018-improvements).

Origin 2018 (9.5) rejects ``FindGraphLayer("[Graph]Layer1")`` (returns None),
which made export_graph report "not found" for graphs that exist.
"""
import pytest
from fakes import FakeGraph, FakeLayer, FakePages

from origin_pro_mcp.origin_connection import require_graph


def test_primary_path_uses_find_graph_layer(fake_origin):
    layer = require_graph("Graph1")
    assert isinstance(layer, FakeLayer)
    assert layer is fake_origin.FindGraphLayer("[Graph1]Layer1")


def test_fallback_walks_graph_pages_when_find_graph_layer_returns_none(fake_origin):
    asked = []

    def find(target):
        asked.append(target)  # and return None, like Origin 2018

    fake_origin.FindGraphLayer = find
    fake_origin.graphs = [FakeGraph("Other"), FakeGraph("Graph1"), FakeGraph("Graph2")]
    layer = require_graph("Graph1")
    assert asked == ["[Graph1]Layer1"]
    assert isinstance(layer, FakeLayer)
    assert layer._graph is fake_origin.graphs[1]   # the SAME-NAME page


class _MultiLayerGraph:
    def __init__(self, name, layers):
        self.Name = name
        self._layers = layers

    @property
    def Layers(self):
        return FakePages(self._layers)


def test_fallback_returns_first_layer(fake_origin):
    fake_origin.find_graph_layer_returns_none = True
    first, second = object(), object()
    fake_origin.graphs = [_MultiLayerGraph("G", [first, second])]
    assert require_graph("G") is first


def test_fallback_skips_same_name_page_without_layers(fake_origin):
    fake_origin.find_graph_layer_returns_none = True
    good = object()
    fake_origin.graphs = [_MultiLayerGraph("G", []), _MultiLayerGraph("G", [good])]
    assert require_graph("G") is good


def test_fallback_not_found_lists_open_graphs(fake_origin):
    fake_origin.find_graph_layer_returns_none = True
    fake_origin.graphs = [FakeGraph("Graph1"), FakeGraph("Fig 2")]
    with pytest.raises(ValueError) as exc:
        require_graph("Missing")
    msg = str(exc.value)
    assert "Graph 'Missing' not found" in msg
    assert "Graph1" in msg and "Fig 2" in msg


def test_not_found_with_no_graphs_says_none(fake_origin):
    fake_origin.find_graph_layer_returns_none = True
    fake_origin.graphs = []
    with pytest.raises(ValueError, match=r"Open graphs: \(none\)"):
        require_graph("Graph1")


def test_fallback_com_error_becomes_not_found(fake_origin):
    class _Boom:
        @property
        def Count(self):
            raise OSError("RPC failure")

    fake_origin.find_graph_layer_returns_none = True
    original = FakePages([FakeGraph("Graph1")])
    calls = {"n": 0}

    class _Origin(type(fake_origin)):
        @property
        def GraphPages(self):
            calls["n"] += 1
            # First access (the fallback walk) explodes; the error message's
            # graph_names() call then sees the real list.
            return _Boom() if calls["n"] == 1 else original

    fake_origin.__class__ = _Origin
    with pytest.raises(ValueError, match="not found") as exc:
        require_graph("Graph1")
    assert "Graph1" in str(exc.value)


def test_export_graph_works_through_fallback(fake_origin, tmp_path, monkeypatch):
    # End-to-end: the user-visible symptom was export_graph reporting
    # "not found" on Origin 2018. Pretend the file was written (as in
    # test_dispatchers) so the existence/size checks pass under the fake.
    from origin_pro_mcp.tools import graph as G

    fake_origin.find_graph_layer_returns_none = True
    monkeypatch.setattr(G.os.path, "exists", lambda _p: True)
    monkeypatch.setattr(G.os.path, "getsize", lambda _p: 12345)
    msg = G.export_graph("Graph1", str(tmp_path / "fig.png"))
    assert "Exported to:" in msg
    assert "expGraph" in " ".join(fake_origin.executed)


def test_export_graph_unknown_graph_still_not_found_through_fallback(fake_origin, tmp_path):
    from origin_pro_mcp.tools import graph as G

    fake_origin.find_graph_layer_returns_none = True
    with pytest.raises(ValueError, match="not found"):
        G.export_graph("Ghost", str(tmp_path / "fig.png"))
