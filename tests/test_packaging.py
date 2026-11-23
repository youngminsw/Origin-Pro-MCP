"""Packaging: the mcp pin matches the 1.x API the code uses, and both
``python -m origin_pro_mcp`` and ``python -m origin_pro_mcp.app`` start the
server through the real ``origin_pro_mcp.app.main``."""
import runpy
import warnings
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


def test_pyproject_pins_mcp_1x():
    text = (REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert '"mcp>=1.8.0,<2"' in text
    assert '"mcp>=1.0.0"' not in text


def test_requirements_txt_pins_mcp_1x():
    lines = (REPO_ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()
    assert "mcp>=1.8.0,<2" in [ln.strip() for ln in lines]


def test_installed_mcp_is_1x():
    from importlib.metadata import version

    major = int(version("mcp").split(".")[0])
    assert major == 1
    from mcp.server.fastmcp import FastMCP  # noqa: F401 — the 1.x import path


@pytest.fixture
def recorded_main(monkeypatch):
    import origin_pro_mcp.app as app

    calls = []
    monkeypatch.setattr(app, "main", lambda: calls.append("main"))
    return calls


def test_python_dash_m_package_calls_app_main(recorded_main):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)  # already-imported notice
        runpy.run_module("origin_pro_mcp", run_name="__main__")
    assert recorded_main == ["main"]


def test_python_dash_m_app_delegates_to_package_module(recorded_main):
    # Run as __main__, app.py is a second copy of the module; it must call the
    # REAL origin_pro_mcp.app.main (whose mcp the tools register on), not its
    # own copy.
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)
        runpy.run_module("origin_pro_mcp.app", run_name="__main__")
    assert recorded_main == ["main"]


def test_importing_main_module_does_not_start_server(recorded_main):
    runpy.run_module("origin_pro_mcp", run_name="not_main")
    assert recorded_main == []
