"""Rigorous, COM-free shim tests.

Two daemon lanes are exercised, both on WSL:

* an *in-process* :class:`Daemon` started with an injected :class:`FakeOrigin`
  factory over real loopback sockets (connect-existing / forward / reconnect /
  heartbeat), exactly like ``test_daemon.py``;
* the *auto-spawn* lane drives ``ensure_daemon`` to spawn a detached
  ``python -m origin_pro_mcp.daemon`` process whose in-package fake origin is
  selected via ``ORIGIN_PRO_MCP_FAKE_ORIGIN=1`` — proving the spawn path end to
  end without COM.
"""
from __future__ import annotations

import inspect
import os
import signal
import threading
import time

import pytest

from fakes import FakeOrigin

from origin_pro_mcp import shim
from origin_pro_mcp.daemon import Daemon, read_lockfile


# --------------------------------------------------------------------------- #
# Helpers / fixtures                                                           #
# --------------------------------------------------------------------------- #


class FakeFactory:
    """Hands out a fresh FakeOrigin (with a unique fake pid) on each call."""

    def __init__(self):
        self.created: list[FakeOrigin] = []
        self._next_pid = 80000
        self._lock = threading.Lock()

    def __call__(self) -> FakeOrigin:
        with self._lock:
            fake = FakeOrigin()
            fake.fake_pid = self._next_pid
            self._next_pid += 1
            self.created.append(fake)
            return fake

    @staticmethod
    def get_pid(instance) -> int:
        return instance.fake_pid


def _real_registry() -> dict:
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp

    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


@pytest.fixture
def in_process_daemon(tmp_path):
    """Factory that starts (and tears down) in-process daemons on one lockfile."""
    daemons: list[Daemon] = []
    lockfile = str(tmp_path / "daemon.json")

    def make(max_size: int = 3):
        factory = FakeFactory()
        daemon = Daemon()
        ok = daemon.start(
            origin_factory=factory,
            registry=_real_registry(),
            max_size=max_size,
            host="127.0.0.1",
            port=0,
            get_pid=FakeFactory.get_pid,
            lockfile_path=lockfile,
        )
        assert ok is True
        daemons.append(daemon)
        return daemon, factory, lockfile

    try:
        yield make
    finally:
        for daemon in daemons:
            daemon.stop()


def _forwarder(client: shim.ShimClient, name: str):
    server = shim.build_shim_server(client)
    return server._tool_manager._tools[name].fn


def _kill_lockfile_daemon(lockfile: str) -> None:
    try:
        data = read_lockfile(lockfile)
    except (OSError, ValueError):
        return
    pids = [data.get("pid"), *list(data.get("child_pids", []))]
    for pid in pids:
        if not pid:
            continue
        try:
            os.kill(int(pid), signal.SIGKILL)
        except (OSError, ProcessLookupError):
            pass


# --------------------------------------------------------------------------- #
# SCHEMA EQUALITY — all 37 tools                                               #
# --------------------------------------------------------------------------- #


def test_shim_schema_is_byte_identical_for_all_tools():
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp as real_mcp

    real_tools = real_mcp._tool_manager._tools
    # A client that never connects (heartbeat off) — schema build is offline.
    client = shim.ShimClient(heartbeat_interval=0)
    shim_server = shim.build_shim_server(client)
    shim_tools = shim_server._tool_manager._tools

    # Same exact set of tool names, and exactly 45 of them: 43 Origin
    # forwarders + list_skills + get_skill (skill tools are served locally by
    # the shim, not forwarded, but must still be schema-identical to the
    # in-process server).
    assert set(shim_tools) == set(real_tools)
    assert len(shim_tools) == 48

    for name, real in real_tools.items():
        fwd = shim_tools[name].fn
        assert inspect.signature(fwd) == inspect.signature(real.fn), name
        assert fwd.__doc__ == real.fn.__doc__, name
        assert fwd.__name__ == real.fn.__name__, name
        # The registered JSON input schema must be byte-identical.
        assert shim_tools[name].parameters == real.parameters, name
        assert shim_tools[name].description == real.description, name


# --------------------------------------------------------------------------- #
# END-TO-END FORWARD (in-process daemon, fake factory)                        #
# --------------------------------------------------------------------------- #


def test_forward_round_trips_to_fake_origin(in_process_daemon):
    _daemon, factory, lockfile = in_process_daemon(max_size=3)
    client = shim.ShimClient(lockfile_path=lockfile, heartbeat_interval=0,
                             call_timeout=5.0, allow_spawn=False)
    run = _forwarder(client, "run_labtalk")
    try:
        result = run(script="myvar = 7;")
        assert "Executed successfully" in result
        assert result.endswith("myvar = 7;")
        # The session's OWN FakeOrigin recorded the command (forward -> daemon
        # -> session worker -> real tool body -> fake).
        assert len(factory.created) == 1
        assert factory.created[0].executed == ["myvar = 7;"]
    finally:
        client.close()


def test_forward_surfaces_tool_errors(in_process_daemon):
    _daemon, _factory, lockfile = in_process_daemon(max_size=3)
    client = shim.ShimClient(lockfile_path=lockfile, heartbeat_interval=0,
                             call_timeout=5.0, allow_spawn=False)
    run = _forwarder(client, "run_labtalk")
    try:
        # A gated token without confirm is NOT executed -> actionable message.
        result = run(script="doc -s;")
        assert "NOT EXECUTED" in result
    finally:
        client.close()


# --------------------------------------------------------------------------- #
# AUTO-SPAWN (detached daemon, env-var fake seam)                             #
# --------------------------------------------------------------------------- #


def test_ensure_daemon_auto_spawns_and_reuses(tmp_path, monkeypatch):
    lockfile = str(tmp_path / "daemon.json")
    monkeypatch.setenv("ORIGIN_PRO_MCP_FAKE_ORIGIN", "1")
    assert not os.path.exists(lockfile)

    try:
        host, port, token = shim.ensure_daemon(lockfile, spawn_timeout=25.0)
        assert os.path.exists(lockfile)
        data = read_lockfile(lockfile)
        assert (host, port, token) == ("127.0.0.1", data["port"], data["token"])

        # A real forward over the spawned daemon succeeds.
        client = shim.ShimClient(lockfile_path=lockfile, heartbeat_interval=0,
                                 call_timeout=5.0, allow_spawn=False)
        try:
            run = _forwarder(client, "run_labtalk")
            assert "Executed successfully" in run(script="spawned = 1;")
        finally:
            client.close()

        # A SECOND ensure reuses the same daemon — no second spawn.
        host2, port2, token2 = shim.ensure_daemon(lockfile, spawn=False)
        assert (host2, port2, token2) == (host, port, token)
        assert read_lockfile(lockfile)["pid"] == data["pid"]
    finally:
        _kill_lockfile_daemon(lockfile)


def test_no_spawn_env_suppresses_autospawn(tmp_path, monkeypatch):
    """ORIGIN_PRO_MCP_NO_SPAWN keeps a killed daemon stopped: ensure_daemon
    refuses to auto-respawn and returns an actionable error instead."""
    lockfile = str(tmp_path / "daemon.json")
    monkeypatch.setenv("ORIGIN_PRO_MCP_NO_SPAWN", "1")
    assert not os.path.exists(lockfile)
    with pytest.raises(RuntimeError, match="auto-spawn is disabled"):
        shim.ensure_daemon(lockfile)  # spawn defaults True, but env forces off
    assert not os.path.exists(lockfile)  # nothing was spawned


# --------------------------------------------------------------------------- #
# RECONNECT / NO-HANG                                                          #
# --------------------------------------------------------------------------- #


def test_reconnect_failure_raises_actionable_error_no_hang(in_process_daemon):
    daemon, _factory, lockfile = in_process_daemon(max_size=3)
    client = shim.ShimClient(lockfile_path=lockfile, heartbeat_interval=0,
                             call_timeout=2.0, allow_spawn=False)
    run = _forwarder(client, "run_labtalk")
    try:
        assert "Executed successfully" in run(script="a = 1;")

        # Kill the daemon mid-session (removes the lockfile too).
        daemon.stop()

        started = time.monotonic()
        with pytest.raises(RuntimeError) as exc:
            run(script="b = 2;")
        elapsed = time.monotonic() - started
        assert elapsed < 10.0, f"call hung for {elapsed:.1f}s"
        # The message is actionable, not a bare socket error.
        assert "Origin daemon" in str(exc.value)
    finally:
        client.close()


def test_reconnect_succeeds_after_daemon_restart(in_process_daemon):
    daemon, _factory, lockfile = in_process_daemon(max_size=3)
    client = shim.ShimClient(lockfile_path=lockfile, heartbeat_interval=0,
                             call_timeout=2.0, allow_spawn=False)
    run = _forwarder(client, "run_labtalk")
    try:
        run(script="a = 1;")
        daemon.stop()

        # A fresh daemon on the same lockfile path -> the next call reconnects.
        _daemon2, factory2, _lockfile = in_process_daemon(max_size=3)
        result = run(script="b = 2;")
        assert "Executed successfully" in result
        assert any("b = 2;" in e
                   for f in factory2.created for e in f.executed)
    finally:
        client.close()


# --------------------------------------------------------------------------- #
# HEARTBEAT NON-CORRUPTION                                                     #
# --------------------------------------------------------------------------- #


def test_concurrent_forwards_with_heartbeats_correlate(in_process_daemon):
    _daemon, _factory, lockfile = in_process_daemon(max_size=3)
    # Aggressive heartbeats fire continuously during the forwards.
    client = shim.ShimClient(lockfile_path=lockfile, heartbeat_interval=0.01,
                             call_timeout=10.0, allow_spawn=False)
    run = _forwarder(client, "run_labtalk")

    results: dict[int, str] = {}
    errors: list = []
    lock = threading.Lock()

    def worker(i: int):
        try:
            r = run(script=f"hb{i} = {i};")
            with lock:
                results[i] = r
        except Exception as exc:  # pragma: no cover - failure path
            with lock:
                errors.append((i, exc))

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(24)]
    try:
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=20.0)
        assert not any(t.is_alive() for t in threads), "a forward hung"
        assert not errors, errors
        assert len(results) == 24
        # Each forward got back ITS OWN response (correct request_id correlation
        # despite heartbeats sharing the socket).
        for i, r in results.items():
            assert r.endswith(f"hb{i} = {i};"), (i, r)
    finally:
        client.close()
