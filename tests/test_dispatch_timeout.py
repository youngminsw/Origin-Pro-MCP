"""D2: per-dispatch timeout recovers a wedged session out-of-band (COM-free).

A tool dispatch whose COM Execute wedges past the dispatch budget must:
  - have the session's Origin PID force-killed by the watchdog,
  - reply the client one actionable reset error (not hang until the socket timeout),
  - free the pool slot so a fresh session works,
all driven by a fake clock + injected terminate hook (no real waits, no COM)."""
from __future__ import annotations

import threading
import time

import pytest

from fakes import FakeOrigin
from origin_pro_mcp import transport
from origin_pro_mcp.daemon import Daemon


class FakeClock:
    def __init__(self, start=0.0):
        self._t = start
        self._lock = threading.Lock()

    def __call__(self):
        with self._lock:
            return self._t

    def advance(self, dt):
        with self._lock:
            self._t += dt


def _wait_until(pred, timeout=3.0, interval=0.005):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if pred():
            return True
        time.sleep(interval)
    return pred()


def _real_registry():
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp
    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


class BlockingExecuteOrigin(FakeOrigin):
    """Its Execute BLOCKS on any script containing 'HANG' until released —
    models an Origin COM call that wedges. Killing the PID (test terminate hook)
    releases it, modeling the com_error that frees the wedged worker thread."""

    def __init__(self):
        super().__init__()
        self.unblock = threading.Event()
        self.execute_entered = threading.Event()

    def Execute(self, script):
        if "HANG" in script:
            self.execute_entered.set()
            self.unblock.wait(10)  # safety cap so a broken test can't wedge forever
            return True
        return super().Execute(script)


def _start(tmp_path, dispatch_timeout, clock):
    origins = []
    killed = []

    def factory():
        o = BlockingExecuteOrigin()
        o.fake_pid = 88001 + len(origins)
        origins.append(o)
        return o

    def terminate(pid):
        killed.append(pid)
        for o in origins:
            if getattr(o, "fake_pid", None) == pid:
                o.unblock.set()  # killing Origin unblocks the wedged Execute

    daemon = Daemon()
    ok = daemon.start(
        origin_factory=factory, registry=_real_registry(), max_size=3,
        host="127.0.0.1", port=0, get_pid=lambda i: i.fake_pid,
        terminate_process=terminate, clock=clock, dispatch_timeout=dispatch_timeout,
        monitor_tick=0.005, reconnect_grace=0.0,
        lockfile_path=str(tmp_path / "daemon.json"),
    )
    assert ok
    return daemon, origins, killed


def _req(conn, rid, name, kwargs):
    conn.send_frame({"type": "request", "request_id": rid, "name": name, "kwargs": kwargs})
    return conn.recv_frame()




def test_dispatch_timeout_recovers_wedged_session(tmp_path):
    clock = FakeClock()
    daemon, origins, killed = _start(tmp_path, dispatch_timeout=5.0, clock=clock)
    conns = []
    try:
        conn = transport.connect(daemon.host, daemon.port, daemon.token, session_id="A")
        conn.settimeout(5.0)
        conns.append(conn)
        # send a request whose Execute will wedge
        conn.send_frame({"type": "request", "request_id": "r1",
                         "name": "run_labtalk", "kwargs": {"script": "HANG;"}})
        # wait until the worker is actually blocked inside Execute, then trip the deadline
        assert _wait_until(lambda: origins and origins[0].execute_entered.is_set())
        clock.advance(6.0)  # past the 5s dispatch deadline
        resp = conn.recv_frame()
        assert resp["ok"] is False
        assert "dispatch timeout" in resp["error"]
        assert "force-reset" in resp["error"]
        # the wedged session's Origin PID was force-killed
        assert killed == [origins[0].fake_pid]
        # a fresh session is served (the pool slot was freed)
        conn2 = transport.connect(daemon.host, daemon.port, daemon.token, session_id="B")
        conn2.settimeout(5.0)
        conns.append(conn2)
        r = _req(conn2, "r2", "list_worksheets", {})
        assert r["ok"] is True
    finally:
        for c in conns:
            c.close()
        daemon.stop()


def test_dispatch_timeout_disabled_by_default_no_kill(tmp_path):
    """dispatch_timeout=0 (default) => no deadline armed; a fast op is unaffected
    and the terminate hook is never called for a normal request."""
    clock = FakeClock()
    daemon, origins, killed = _start(tmp_path, dispatch_timeout=0.0, clock=clock)
    conns = []
    try:
        conn = transport.connect(daemon.host, daemon.port, daemon.token, session_id="A")
        conn.settimeout(5.0)
        conns.append(conn)
        r = _req(conn, "r1", "list_worksheets", {})  # fast, non-hang
        assert r["ok"] is True
        clock.advance(100.0)  # would trip any deadline if one were armed
        time.sleep(0.05)
        assert killed == []  # nothing armed => nothing killed
    finally:
        for c in conns:
            c.close()
        daemon.stop()


def test_run_labtalk_per_call_timeout_override_arms_when_global_off(tmp_path):
    """run_labtalk(timeout=N) bounds a single call even when the global
    dispatch timeout is off (opt-in per call)."""
    clock = FakeClock()
    daemon, origins, killed = _start(tmp_path, dispatch_timeout=0.0, clock=clock)
    conns = []
    try:
        conn = transport.connect(daemon.host, daemon.port, daemon.token, session_id="A")
        conn.settimeout(5.0)
        conns.append(conn)
        conn.send_frame({"type": "request", "request_id": "r1", "name": "run_labtalk",
                         "kwargs": {"script": "HANG;", "timeout": 5}})
        assert _wait_until(lambda: origins and origins[0].execute_entered.is_set())
        clock.advance(6.0)  # past the per-call 5s override
        resp = conn.recv_frame()
        assert resp["ok"] is False
        assert "dispatch timeout" in resp["error"]
        assert "5s" in resp["error"]  # message reflects the per-call budget, not the global 0
        assert killed == [origins[0].fake_pid]
    finally:
        for c in conns:
            c.close()
        daemon.stop()


def test_reconcile_call_timeout_env(monkeypatch):
    from origin_pro_mcp.shim import _reconcile_call_timeout
    monkeypatch.delenv("ORIGIN_PRO_MCP_DISPATCH_TIMEOUT", raising=False)
    assert _reconcile_call_timeout(60.0) == 60.0
    for off in ("off", "false", "no", "0", ""):
        monkeypatch.setenv("ORIGIN_PRO_MCP_DISPATCH_TIMEOUT", off)
        assert _reconcile_call_timeout(60.0) == 60.0
    monkeypatch.setenv("ORIGIN_PRO_MCP_DISPATCH_TIMEOUT", "30")
    assert _reconcile_call_timeout(60.0) == 60.0     # 30+15 < 60 => keep larger
    monkeypatch.setenv("ORIGIN_PRO_MCP_DISPATCH_TIMEOUT", "90")
    assert _reconcile_call_timeout(60.0) == 105.0    # 90+15 > 60
    monkeypatch.setenv("ORIGIN_PRO_MCP_DISPATCH_TIMEOUT", "garbage")
    assert _reconcile_call_timeout(60.0) == 60.0     # unparseable => unchanged
