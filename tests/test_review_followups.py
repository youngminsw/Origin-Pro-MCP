"""Follow-ups from the 2026-11-23 review of the fork.

* transform(interpolate / find_peaks) accept sheet names with spaces, like the
  other transform methods; interpolate addresses its source as the active
  sheet, so the name never goes into a LabTalk range string.
* sheet names containing LabTalk substitution syntax (``%``, ``$(``) are
  rejected instead of being silently rewritten.
* require_graph's GraphPages fallback matches window names case-insensitively.
* winproc.process_running reports an access-denied process as running.
"""
import os
import sys

import pytest
from fakes import FakeBook, FakeGraph, FakeLayer

from origin_pro_mcp.labtalk_safe import labtalk_sheet_name
from origin_pro_mcp.origin_connection import require_graph

SPACED = "wpd datasets 1"


@pytest.fixture
def spaced_book(fake_origin):
    fake_origin.books.append(FakeBook("Data", sheets=(SPACED,)))
    return fake_origin


# --- transform: interpolate / find_peaks ------------------------------------ #


def test_interpolate_uses_active_sheet_not_a_range_string(fake_origin):
    from origin_pro_mcp.tools.analysis import transform

    transform("Book1", "Sheet1", 1, 2, method="interpolate", num_points=10)
    interp = [c for c in fake_origin.executed if c.startswith("interp1 ")]
    assert interp, fake_origin.executed
    assert "iy:=(1,2)" in interp[-1]
    assert "[Book1]Sheet1!" not in interp[-1]
    # The source sheet is activated before interp1 runs.
    before = fake_origin.executed[: fake_origin.executed.index(interp[-1])]
    assert any('page.active$ = "Sheet1"' in c and "Interp" not in c for c in before)


@pytest.mark.parametrize("method", ["interpolate", "find_peaks"])
def test_transform_accepts_sheet_names_with_spaces(spaced_book, method):
    from origin_pro_mcp.tools.analysis import transform

    try:
        transform("Data", SPACED, 1, 2, method=method)
    except ValueError as exc:
        # The fake has no real data, so a data-related error is fine; only the
        # name validation must not reject the sheet.
        assert "must start with a letter" not in str(exc), exc


# --- LabTalk substitution syntax in sheet names ----------------------------- #


@pytest.mark.parametrize("name", ["50% load", "a%Hb", "x$(1+1)"])
def test_sheet_name_with_substitution_syntax_is_rejected(name):
    with pytest.raises(ValueError, match="cannot contain"):
        labtalk_sheet_name(name, "sheet_name")


@pytest.mark.parametrize("name", ["Price $", "cost$", "run (1)", "wpd datasets 1"])
def test_plain_dollar_and_parentheses_are_still_allowed(name):
    assert labtalk_sheet_name(name, "sheet_name") == name


# --- require_graph fallback: case-insensitive ------------------------------- #


def test_graph_fallback_matches_name_case_insensitively(fake_origin):
    fake_origin.FindGraphLayer = lambda target: None      # Origin 2018 behavior
    fake_origin.graphs = [FakeGraph("Other"), FakeGraph("Graph1")]
    layer = require_graph("graph1")
    assert isinstance(layer, FakeLayer)
    assert layer._graph is fake_origin.graphs[1]


def test_graph_fallback_still_raises_for_a_missing_graph(fake_origin):
    fake_origin.FindGraphLayer = lambda target: None
    fake_origin.graphs = [FakeGraph("Graph1")]
    with pytest.raises(ValueError, match="not found"):
        require_graph("Graph2")


# --- winproc.process_running ------------------------------------------------ #

windows_only = pytest.mark.skipif(sys.platform != "win32", reason="Windows process API")


@windows_only
def test_process_running_own_process():
    from origin_pro_mcp.winproc import process_running

    assert process_running(os.getpid()) is True


@windows_only
def test_process_running_missing_pid_is_false():
    from origin_pro_mcp.winproc import process_running

    assert process_running(0) is False
    assert process_running(999_999) is False


@windows_only
def test_process_running_access_denied_counts_as_running():
    # PID 4 is the System process: it always exists, and opening it with the
    # rights used here fails with ERROR_ACCESS_DENIED for a normal user. The
    # old check treated that failure as "process gone".
    from origin_pro_mcp.winproc import process_running

    assert process_running(4) is True
