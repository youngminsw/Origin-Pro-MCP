import importlib.util
import sys

import pytest

from origin_pro_mcp import origin_connection


WINDOWS_ORIGIN_AVAILABLE = (
    sys.platform == "win32" and importlib.util.find_spec("win32com") is not None
)


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
