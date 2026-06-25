"""The legend auto-fallback decision: corner when a corner is clear, outside
the frame when every corner still overlaps data. Validated visually against
real Origin 2020 (the outside-right LabTalk recipe); these tests pin the
decision logic without COM."""
import origin_pro_mcp.tools.style_helpers as sh


def test_clear_corner_uses_corner(monkeypatch):
    calls = []
    monkeypatch.setattr(sh, "choose_legend_corner_overlap",
                        lambda g, p=None: ("top-left", 0))
    monkeypatch.setattr(sh, "_place_legend",
                        lambda g, c: calls.append(("corner", g, c)))
    monkeypatch.setattr(sh, "_place_legend_outside",
                        lambda g: calls.append(("outside", g)) or True)
    result = sh.place_legend_avoiding_data("Graph1")
    assert result == "top-left"
    assert calls == [("corner", "Graph1", "top-left")]


def test_all_corners_overlap_falls_back_outside(monkeypatch):
    calls = []
    monkeypatch.setattr(sh, "choose_legend_corner_overlap",
                        lambda g, p=None: ("top-right", 4))
    monkeypatch.setattr(sh, "_place_legend",
                        lambda g, c: calls.append(("corner", g, c)))
    monkeypatch.setattr(sh, "_place_legend_outside",
                        lambda g: calls.append(("outside", g)) or True)
    result = sh.place_legend_avoiding_data("Graph1")
    assert result == "outside-right"
    assert calls == [("outside", "Graph1")]


def test_outside_unavailable_falls_back_to_corner(monkeypatch):
    # If geometry can't be read, outside-placement returns False and we still
    # place at the least-bad corner rather than leaving the legend unplaced.
    calls = []
    monkeypatch.setattr(sh, "choose_legend_corner_overlap",
                        lambda g, p=None: ("bottom-left", 3))
    monkeypatch.setattr(sh, "_place_legend",
                        lambda g, c: calls.append(("corner", g, c)))
    monkeypatch.setattr(sh, "_place_legend_outside", lambda g: False)
    result = sh.place_legend_avoiding_data("Graph1")
    assert result == "bottom-left"
    assert calls == [("corner", "Graph1", "bottom-left")]
