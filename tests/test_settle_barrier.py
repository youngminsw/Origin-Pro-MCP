"""Fake tests for the Task 0.5 settle barrier (style_helpers.settle_new_plots).

A freshly created graph page can silently ignore the FIRST styling/read/
export command issued against it (probe-confirmed) even though no exception
is raised. settle_new_plots polls get_plot_info until the expected plot count
appears, then adds a short settle tail — skipped when the plots were already
there on the very first poll, so it stays cheap in the common (settled) case.
"""
import time

import pytest


def test_settle_new_plots_fast_when_already_settled(fake_origin, monkeypatch):
    """First poll already sees the expected count -> no tail sleep, no delay."""
    from origin_pro_mcp.tools import style_helpers as H

    monkeypatch.setattr(H, "get_plot_info", lambda name: [{"name": "d1"}])
    start = time.monotonic()
    H.settle_new_plots("Graph1", expected_min_plots=1)
    assert time.monotonic() - start < 0.05


def test_settle_new_plots_bails_fast_on_unresolvable_graph(fake_origin, monkeypatch):
    """A graph that can't be resolved at all (e.g. a test double's CreatePage
    never registering the page) is not a timing issue — must not spin for the
    full timeout."""
    from origin_pro_mcp.tools import style_helpers as H

    def _raise(name):
        raise ValueError("not found")

    monkeypatch.setattr(H, "get_plot_info", _raise)
    start = time.monotonic()
    H.settle_new_plots("Ghost", expected_min_plots=1, timeout_s=4.0)
    assert time.monotonic() - start < 0.05


def test_settle_new_plots_polls_until_enumerated(fake_origin, monkeypatch):
    """Plots appear only after a couple of polls -> the tail sleep runs once
    plots are found (not settled on the first try)."""
    from origin_pro_mcp.tools import style_helpers as H

    calls = {"n": 0}

    def _delayed(name):
        calls["n"] += 1
        if calls["n"] < 3:
            return []
        return [{"name": "d1"}]

    monkeypatch.setattr(H, "get_plot_info", _delayed)
    H.settle_new_plots("Graph1", expected_min_plots=1, timeout_s=4.0)
    assert calls["n"] == 3


def test_settle_new_plots_gives_up_at_timeout(fake_origin, monkeypatch):
    """Plots never reach the expected count -> returns at the timeout instead
    of hanging forever."""
    from origin_pro_mcp.tools import style_helpers as H

    monkeypatch.setattr(H, "get_plot_info", lambda name: [])
    start = time.monotonic()
    H.settle_new_plots("Graph1", expected_min_plots=5, timeout_s=0.2)
    elapsed = time.monotonic() - start
    assert 0.15 <= elapsed < 1.0
