"""Daemon spawn outside the MCP client's Job Object (shim._spawn_daemon).

The Python MCP SDK client runs every server inside a Windows Job Object with
KILL_ON_JOB_CLOSE, so a Popen'd daemon dies with the client and never reaps its
sessions. When the shim is in a job it starts the daemon through WMI instead,
falling back to Popen on any failure. These tests fake the job check, WMI and
Popen, so they run on any platform.
"""
from __future__ import annotations

import sys
import types

import pytest

from origin_pro_mcp import shim


# --------------------------------------------------------------------------- #
# _spawn_daemon routing                                                       #
# --------------------------------------------------------------------------- #


@pytest.fixture
def spawn_env(monkeypatch, tmp_path):
    calls = {"popen": [], "wmi": []}

    def fake_popen(args, **kwargs):
        calls["popen"].append((args, kwargs))
        return types.SimpleNamespace(pid=1)

    monkeypatch.setattr(shim.subprocess, "Popen", fake_popen)
    monkeypatch.setattr(shim.sys, "platform", "win32")
    # CREATE_NEW_PROCESS_GROUP only exists on Windows' subprocess module.
    monkeypatch.setattr(shim.subprocess, "CREATE_NEW_PROCESS_GROUP", 0x200, raising=False)
    return calls, str(tmp_path / "daemon.json")


def test_in_job_and_wmi_succeeds_skips_popen(spawn_env, monkeypatch):
    calls, lockfile = spawn_env
    monkeypatch.setattr(shim, "_in_job_object", lambda: True)

    def fake_wmi(env, cwd):
        calls["wmi"].append((env, cwd))
        return True

    monkeypatch.setattr(shim, "_spawn_outside_job", fake_wmi)
    shim._spawn_daemon(lockfile)
    assert len(calls["wmi"]) == 1
    assert calls["popen"] == []
    env, _cwd = calls["wmi"][0]
    assert env["ORIGIN_PRO_MCP_LOCKFILE"] == lockfile
    assert "PYTHONPATH" in env


def test_in_job_but_wmi_fails_falls_back_to_popen(spawn_env, monkeypatch):
    calls, lockfile = spawn_env
    monkeypatch.setattr(shim, "_in_job_object", lambda: True)
    monkeypatch.setattr(shim, "_spawn_outside_job", lambda env, cwd: False)
    shim._spawn_daemon(lockfile)
    assert len(calls["popen"]) == 1
    args, kwargs = calls["popen"][0]
    assert args[-2:] == ["-m", "origin_pro_mcp.daemon"]
    assert kwargs["creationflags"] & 0x00000008  # DETACHED_PROCESS


def test_not_in_job_uses_popen_without_wmi(spawn_env, monkeypatch):
    calls, lockfile = spawn_env
    monkeypatch.setattr(shim, "_in_job_object", lambda: False)

    def must_not_call(env, cwd):
        raise AssertionError("WMI spawn must not be used outside a job")

    monkeypatch.setattr(shim, "_spawn_outside_job", must_not_call)
    shim._spawn_daemon(lockfile)
    assert len(calls["popen"]) == 1


def test_non_windows_never_checks_job(monkeypatch, tmp_path):
    popen_calls = []
    monkeypatch.setattr(shim.subprocess, "Popen",
                        lambda args, **kw: popen_calls.append(kw))
    monkeypatch.setattr(shim.sys, "platform", "linux")

    def must_not_call():
        raise AssertionError("job check is Windows-only")

    monkeypatch.setattr(shim, "_in_job_object", must_not_call)
    shim._spawn_daemon(str(tmp_path / "daemon.json"))
    assert len(popen_calls) == 1
    assert popen_calls[0].get("start_new_session") is True


# --------------------------------------------------------------------------- #
# _in_job_object                                                              #
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize("value", ["0", "off", "false", "no", "OFF"])
def test_in_job_object_opt_out(monkeypatch, value):
    monkeypatch.setenv("ORIGIN_PRO_MCP_DAEMON_BREAKAWAY", value)
    fake_job = types.SimpleNamespace(IsProcessInJob=lambda h, j: True)
    monkeypatch.setitem(sys.modules, "win32job", fake_job)
    monkeypatch.setitem(sys.modules, "win32api",
                        types.SimpleNamespace(GetCurrentProcess=lambda: -1))
    assert shim._in_job_object() is False


@pytest.mark.parametrize("in_job", [True, False])
def test_in_job_object_uses_win32job(monkeypatch, in_job):
    monkeypatch.delenv("ORIGIN_PRO_MCP_DAEMON_BREAKAWAY", raising=False)
    seen = {}

    def is_in_job(handle, job):
        seen["args"] = (handle, job)
        return in_job

    monkeypatch.setitem(sys.modules, "win32job", types.SimpleNamespace(IsProcessInJob=is_in_job))
    monkeypatch.setitem(sys.modules, "win32api",
                        types.SimpleNamespace(GetCurrentProcess=lambda: -1))
    assert shim._in_job_object() is in_job
    assert seen["args"] == (-1, None)


def test_in_job_object_never_raises(monkeypatch):
    """pywin32 missing AND no ctypes.windll (non-Windows) -> False, no exception."""
    monkeypatch.delenv("ORIGIN_PRO_MCP_DAEMON_BREAKAWAY", raising=False)
    monkeypatch.setitem(sys.modules, "win32job", None)  # import -> ImportError
    monkeypatch.setitem(sys.modules, "win32api", None)
    if sys.platform == "win32":
        pytest.skip("ctypes fallback is real on Windows; covered by live tests")
    assert shim._in_job_object() is False


# --------------------------------------------------------------------------- #
# _spawn_outside_job (fake WMI)                                               #
# --------------------------------------------------------------------------- #


class _Obj(types.SimpleNamespace):
    pass


class _FakeProcessClass:
    def __init__(self, return_value):
        self.return_value = return_value
        self.params = _Obj()
        self.executed = None

    def Methods_(self, name):
        assert name == "Create"
        params = self.params
        return _Obj(InParameters=_Obj(SpawnInstance_=lambda: params))

    def ExecMethod_(self, name, params):
        assert name == "Create"
        self.executed = params
        return _Obj(ReturnValue=self.return_value, ProcessId=4321)


class _FakeWmi:
    def __init__(self, return_value=0):
        self.startup = _Obj()
        self.process = _FakeProcessClass(return_value)

    def Get(self, cls):
        if cls == "Win32_ProcessStartup":
            startup = self.startup
            return _Obj(SpawnInstance_=lambda: startup)
        if cls == "Win32_Process":
            return self.process
        raise AssertionError(cls)


@pytest.fixture
def fake_com(monkeypatch):
    state = {"wmi": _FakeWmi(), "init": 0, "uninit": 0, "moniker": None}

    def get_object(moniker):
        state["moniker"] = moniker
        return state["wmi"]

    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: state.__setitem__("init", state["init"] + 1)
    pythoncom.CoUninitialize = lambda: state.__setitem__("uninit", state["uninit"] + 1)
    client = types.ModuleType("win32com.client")
    client.GetObject = get_object
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    return state


def test_spawn_outside_job_builds_wmi_call(fake_com):
    env = {"ORIGIN_PRO_MCP_LOCKFILE": r"C:\x\daemon.json", "PYTHONPATH": r"C:\src", "": "skip"}
    assert shim._spawn_outside_job(env, r"C:\Temp") is True
    wmi = fake_com["wmi"]
    assert "root\\cimv2" in fake_com["moniker"]
    params = wmi.process.executed
    assert params.CommandLine.endswith('-m origin_pro_mcp.daemon')
    assert params.CommandLine.startswith(f'"{sys.executable}"')
    assert params.CurrentDirectory == r"C:\Temp"
    startup = params.ProcessStartupInformation
    # WMI rejects CREATE_NO_WINDOW (ReturnValue 21); only DETACHED_PROCESS.
    assert startup.CreateFlags == 0x00000008
    assert startup.ShowWindow == 0
    assert r"ORIGIN_PRO_MCP_LOCKFILE=C:\x\daemon.json" in startup.EnvironmentVariables
    assert r"PYTHONPATH=C:\src" in startup.EnvironmentVariables
    assert not any(v.startswith("=") for v in startup.EnvironmentVariables)
    assert fake_com["init"] == 1 and fake_com["uninit"] == 1


def test_spawn_outside_job_nonzero_return_is_failure(fake_com):
    fake_com["wmi"] = _FakeWmi(return_value=21)
    assert shim._spawn_outside_job({"A": "1"}, r"C:\Temp") is False
    assert fake_com["uninit"] == 1  # COM is always uninitialized


def test_spawn_outside_job_exception_is_failure(fake_com, monkeypatch):
    def boom(moniker):
        raise RuntimeError("WMI unavailable")

    monkeypatch.setattr(sys.modules["win32com.client"], "GetObject", boom)
    assert shim._spawn_outside_job({"A": "1"}, r"C:\Temp") is False
    assert fake_com["uninit"] == 1


def test_spawn_outside_job_without_pywin32(monkeypatch):
    monkeypatch.setitem(sys.modules, "pythoncom", None)
    assert shim._spawn_outside_job({"A": "1"}, r"C:\Temp") is False
