import importlib.util
import sys

import pytest

from origin_pro_mcp import origin_connection


WINDOWS_ORIGIN_AVAILABLE = (
    sys.platform == "win32" and importlib.util.find_spec("win32com") is not None
)


@pytest.fixture(autouse=True)
def _no_daemon_breakaway(monkeypatch):
    """Keep test-spawned daemons inside the test runner's process tree.

    When pytest itself runs inside a Job Object (e.g. launched from a shell or
    an agent harness that uses one), the shim would start real daemons through
    WMI, outside the job — and the tests' own cleanup could not reach them, so
    they would linger. Tests that exercise the WMI path delete this variable
    or patch ``shim._in_job_object`` themselves.
    """
    monkeypatch.setenv("ORIGIN_PRO_MCP_DAEMON_BREAKAWAY", "0")


def pytest_configure(config: pytest.Config) -> None:
    config.addinivalue_line(
        "markers",
        "requires_origin: test requires Windows Python, pywin32, and Origin Pro COM",
    )


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    if WINDOWS_ORIGIN_AVAILABLE:
        return
    skip_origin = pytest.mark.skip(
        reason="requires Windows Python with pywin32 and Origin Pro COM"
    )
    for item in items:
        if "requires_origin" in item.keywords:
            item.add_marker(skip_origin)


# --- Shared fake Origin COM doubles (used by test_tool_guards, test_cli) ---
# The fake classes live in tests/fakes.py so the daemon/transport tests can
# import the exact same COM surface; re-exported here for existing imports.
from fakes import (  # noqa: E402,F401
    FakeBook,
    FakeColumn,
    FakeGraph,
    FakeLayer,
    FakeMatrix,
    FakeOrigin,
    FakePages,
    FakePlot,
    FakeSheet,
)


@pytest.fixture
def fake_origin():
    fake = FakeOrigin()
    origin_connection.set_session_origin(fake)
    try:
        yield fake
    finally:
        origin_connection.clear_session_origin()
