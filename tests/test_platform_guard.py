"""The MCP entry point refuses to start outside Windows 10/11 (app.platform_problem).

Only ``app.main()`` checks — importing the package, the daemon module and the
CLI must keep working on Linux (these unit tests themselves run there).
"""
from __future__ import annotations

import platform
import sys
import types

import pytest

from origin_pro_mcp import app


def _winver(major, minor=0, build=19045):
    return lambda: types.SimpleNamespace(major=major, minor=minor, build=build)


@pytest.fixture(autouse=True)
def _no_bypass(monkeypatch):
    monkeypatch.delenv("ORIGIN_PRO_MCP_SKIP_PLATFORM_CHECK", raising=False)


@pytest.mark.parametrize("major,minor,build", [(10, 0, 19045), (10, 0, 22631), (10, 0, 26100)])
def test_windows_10_and_11_allowed(monkeypatch, major, minor, build):
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(sys, "getwindowsversion", _winver(major, minor, build), raising=False)
    assert app.platform_problem() is None


@pytest.mark.parametrize("major,minor", [(6, 1), (6, 3)])  # Windows 7, 8.1
def test_old_windows_rejected(monkeypatch, major, minor):
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(sys, "getwindowsversion", _winver(major, minor, 9600), raising=False)
    msg = app.platform_problem()
    assert msg and "Windows 10 or 11" in msg and f"NT {major}.{minor}" in msg


def test_unknown_windows_version_does_not_block(monkeypatch):
    def boom():
        raise RuntimeError("no version")

    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(sys, "getwindowsversion", boom, raising=False)
    assert app.platform_problem() is None


@pytest.mark.parametrize("plat", ["linux", "darwin"])
def test_non_windows_rejected(monkeypatch, plat):
    monkeypatch.setattr(sys, "platform", plat)
    monkeypatch.setattr(platform, "release", lambda: "6.8.0-generic")
    msg = app.platform_problem()
    assert msg and "Windows 10/11" in msg and plat in msg


@pytest.mark.parametrize("release", ["5.15.153.1-microsoft-standard-WSL2", "4.4.0-19041-Microsoft"])
def test_wsl_gets_specific_hint(monkeypatch, release):
    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setattr(platform, "release", lambda: release)
    msg = app.platform_problem()
    assert msg and "WSL" in msg and "origin-pro-mcp.exe" in msg


@pytest.mark.parametrize("value", ["1", "true", "YES", "on"])
def test_bypass_env(monkeypatch, value):
    monkeypatch.setenv("ORIGIN_PRO_MCP_SKIP_PLATFORM_CHECK", value)
    monkeypatch.setattr(sys, "platform", "linux")
    assert app.platform_problem() is None


def test_main_exits_before_starting_anything(monkeypatch, capsys):
    monkeypatch.setattr(app, "platform_problem", lambda: "not supported here")

    def must_not_run():
        raise AssertionError("server must not start on an unsupported platform")

    monkeypatch.setattr(app, "use_daemon", must_not_run)
    with pytest.raises(SystemExit) as exc:
        app.main()
    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "not supported here" in err and "ORIGIN_PRO_MCP_SKIP_PLATFORM_CHECK" in err


def test_main_proceeds_when_supported(monkeypatch):
    monkeypatch.setattr(app, "platform_problem", lambda: None)
    started = []
    monkeypatch.setattr(app, "use_daemon", lambda: True)
    fake_shim = types.ModuleType("origin_pro_mcp.shim")
    fake_shim.run_stdio = lambda: started.append(True)
    monkeypatch.setitem(sys.modules, "origin_pro_mcp.shim", fake_shim)
    app.main()
    assert started == [True]
