import importlib.util
import sys

import pytest


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
