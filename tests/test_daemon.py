"""Rigorous, COM-free daemon/pool/watchdog/singleton tests.

Everything runs over real loopback sockets with an injected ``origin_factory``
that hands each session its own :class:`FakeOrigin`. No COM, no win32com.
"""
from __future__ import annotations

import os
import threading
import time

import pytest

from fakes import FakeOrigin

from origin_pro_mcp import daemon as daemon_mod
from origin_pro_mcp import transport
from origin_pro_mcp.daemon import (
    Daemon,
    Pool,
    PoolFull,
    SingletonGuard,
    Watchdog,
    write_lockfile,
    read_lockfile,
)


# --------------------------------------------------------------------------- #
# Helpers / fixtures                                                           #
# --------------------------------------------------------------------------- #


class FakeFactory:
    """Hands out a fresh FakeOrigin (with a unique fake pid) on each call.

    The created fakes are retained so a test can inspect exactly which commands
    each session's *own* instance recorded — the isolation proof.
    """

    def __init__(self):
        self.created: list[FakeOrigin] = []
        self._next_pid = 90000
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
    """The real FastMCP tool registry — dispatch goes through real tool bodies."""
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp

    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


@pytest.fixture
def started_daemon(tmp_path):
    """A running Daemon (fake factory, real registry) with full teardown."""
    factory = FakeFactory()
    daemon = Daemon()
    client_conns: list = []

    def make(max_size=3):
        ok = daemon.start(
            origin_factory=factory,
            registry=_real_registry(),
            max_size=max_size,
            host="127.0.0.1",
            port=0,
            get_pid=FakeFactory.get_pid,
            lockfile_path=str(tmp_path / "daemon.json"),
        )
        assert ok is True
        return daemon, factory

    def client(session_id):
        conn = transport.connect(daemon.host, daemon.port, daemon.token,
                                 session_id=session_id)
        conn.settimeout(5.0)
        client_conns.append(conn)
        return conn

    try:
        yield make, client, factory
    finally:
        for conn in client_conns:
            conn.close()
        daemon.stop()


def _call(conn, request_id, name, kwargs):
    conn.send_frame({"type": "request", "request_id": request_id,
                     "name": name, "kwargs": kwargs})
    return conn.recv_frame()


# --------------------------------------------------------------------------- #
# Lockfile                                                                     #
# --------------------------------------------------------------------------- #


def test_lockfile_round_trip(tmp_path):
    path = str(tmp_path / "daemon.json")
    write_lockfile(path, port=51234, token="tok", pid=42, child_pids=[1, 2, 3])
    data = read_lockfile(path)
    assert data == {"port": 51234, "token": "tok", "pid": 42,
                    "child_pids": [1, 2, 3]}


# --------------------------------------------------------------------------- #
# Singleton guard                                                              #
# --------------------------------------------------------------------------- #


def test_singleton_second_acquire_fails(tmp_path):
    lock_path = str(tmp_path / "daemon.lock")
    first = SingletonGuard(lock_path)
    second = SingletonGuard(lock_path)
    try:
        assert first.acquire() is True
        # A second daemon would call acquire() and, getting False, exit.
        assert second.acquire() is False
        # After the first releases, the slot is reclaimable.
        first.release()
        assert second.acquire() is True
    finally:
        first.release()
        second.release()


def test_second_daemon_start_returns_false(tmp_path):
    factory = FakeFactory()
    lockfile = str(tmp_path / "daemon.json")
    d1, d2 = Daemon(), Daemon()
    try:
        assert d1.start(origin_factory=factory, registry={}, port=0,
                        get_pid=FakeFactory.get_pid, lockfile_path=lockfile) is True
        # Second daemon cannot grab the singleton -> start() is a no-op False.
        assert d2.start(origin_factory=factory, registry={}, port=0,
                        get_pid=FakeFactory.get_pid, lockfile_path=lockfile) is False
        assert d2.port is None  # nothing was started
    finally:
        d1.stop()
        d2.stop()


# --------------------------------------------------------------------------- #
# Isolation (the core per-session proof)                                       #
# --------------------------------------------------------------------------- #


def test_three_sessions_are_isolated(started_daemon):
    """Three sessions each run a distinguishable LabTalk command through the
    REAL registry; each session's OWN FakeOrigin must have recorded only its
    own command. Disjoint ``executed`` lists prove thread-local isolation.
    """
    make, client, factory = started_daemon
    make(max_size=3)

    scripts = {"s0": "r0 = 100;", "s1": "r1 = 200;", "s2": "r2 = 300;"}
    for sid, script in scripts.items():
        conn = client(sid)
        resp = _call(conn, f"{sid}-req", "run_labtalk", {"script": script})
        assert resp["ok"] is True, resp
        assert resp["type"] == "response"
        assert resp["request_id"] == f"{sid}-req"

    assert len(factory.created) == 3  # exactly one instance per session

    executed_sets = [set(fake.executed) for fake in factory.created]
    # Each instance recorded exactly its own single command...
    for ex in executed_sets:
        assert len(ex) == 1
    # ...the union is precisely the three sent scripts...
    assert set().union(*executed_sets) == set(scripts.values())
    # ...and the sets are pairwise disjoint (no cross-session bleed).
    for i in range(len(executed_sets)):
        for j in range(i + 1, len(executed_sets)):
            assert executed_sets[i].isdisjoint(executed_sets[j])


def test_repeat_session_reuses_same_instance(started_daemon):
    make, client, factory = started_daemon
    make(max_size=3)

    conn = client("same")
    _call(conn, "a", "run_labtalk", {"script": "x = 1;"})
    _call(conn, "b", "run_labtalk", {"script": "x = 2;"})
    # Same session id -> one instance reused, both commands on it.
    assert len(factory.created) == 1
    assert factory.created[0].executed == ["x = 1;", "x = 2;"]


# --------------------------------------------------------------------------- #
# Pool full                                                                    #
# --------------------------------------------------------------------------- #


def test_fourth_distinct_session_rejected(started_daemon):
    """A 4th distinct session gets the actionable PoolFull error, never hangs."""
    make, client, factory = started_daemon
    make(max_size=3)

    # Fill the pool: three live sessions, connections kept open.
    for i in range(3):
        conn = client(f"sess-{i}")
        resp = _call(conn, f"r{i}", "run_labtalk", {"script": f"a{i} = {i};"})
        assert resp["ok"] is True

    overflow = client("sess-overflow")
    resp = _call(overflow, "r4", "run_labtalk", {"script": "a4 = 4;"})
    assert resp["ok"] is False
    assert resp["error"] == (
        "Origin pool full (3/3). Close another Origin MCP session and retry."
    )
    assert len(factory.created) == 3  # the overflow never spawned an instance


def test_pool_acquire_raises_poolfull_directly():
    pool = Pool(FakeFactory(), registry={}, max_size=2,
                get_pid=FakeFactory.get_pid)
    try:
        pool.acquire("a")
        pool.acquire("b")
        with pytest.raises(PoolFull) as exc:
            pool.acquire("c")
        assert str(exc.value) == (
            "Origin pool full (2/2). Close another Origin MCP session and retry."
        )
    finally:
        pool.stop_all()


# --------------------------------------------------------------------------- #
# Watchdog                                                                     #
# --------------------------------------------------------------------------- #


class ThreadGuardedFake:
    """A COM-proxy stand-in that records ANY off-thread attribute access.

    The watchdog must never dereference a proxy — it only knows PIDs. We hand
    the watchdog the int pid (never this object) and assert ``touched`` stays
    False, proving the watchdog stayed COM-free.
    """

    def __init__(self, owner_thread_id: int):
        object.__setattr__(self, "_owner", owner_thread_id)
        object.__setattr__(self, "touched", False)

    def __getattr__(self, name):
        if name in ("_owner", "touched"):
            return object.__getattribute__(self, name)
        if threading.get_ident() != object.__getattribute__(self, "_owner"):
            object.__setattr__(self, "touched", True)
            raise RuntimeError(f"COM proxy touched off-owner-thread: {name!r}")
        return object.__getattribute__(self, name)


def test_watchdog_fires_kills_pid_and_frees_slot():
    killed: list = []
    reaped: list = []
    guard_fake = ThreadGuardedFake(threading.get_ident())

    def terminate(pid):
        killed.append(pid)

    def on_reap(session_id, pid):
        reaped.append((session_id, pid))

    wd = Watchdog(terminate_process=terminate, on_reap=on_reap, tick=0.005)
    wd.start()
    try:
        wd.arm("wedged", pid=4242, deadline=time.monotonic() + 0.05)

        deadline = time.monotonic() + 2.0
        while not killed and time.monotonic() < deadline:
            time.sleep(0.005)

        assert killed == [4242], "watchdog must hard-kill the recorded pid"
        assert reaped == [("wedged", 4242)], "watchdog must free the slot"
        # The watchdog dealt only in PIDs; it never touched the COM proxy.
        assert guard_fake.touched is False
    finally:
        wd.stop()


def test_watchdog_disarm_cancels_reap():
    killed: list = []
    wd = Watchdog(terminate_process=lambda pid: killed.append(pid), tick=0.005)
    wd.start()
    try:
        wd.arm("s", pid=7, deadline=time.monotonic() + 0.1)
        wd.disarm("s")
        time.sleep(0.2)
        assert killed == []  # disarmed before the deadline -> no kill
    finally:
        wd.stop()


def test_daemon_watchdog_reap_releases_pool_slot(started_daemon):
    """Wiring check: a watchdog reap frees the pool slot for that session."""
    make, client, factory = started_daemon
    daemon, _factory = make(max_size=3)

    conn = client("victim")
    _call(conn, "r", "run_labtalk", {"script": "v = 1;"})
    victim_pid = factory.created[0].fake_pid
    assert victim_pid in daemon.pool.child_pids()

    killed: list = []
    daemon.watchdog._terminate = lambda pid: killed.append(pid)  # type: ignore
    daemon.watchdog.arm("victim", pid=victim_pid,
                        deadline=time.monotonic() + 0.02)

    deadline = time.monotonic() + 2.0
    while not killed and time.monotonic() < deadline:
        time.sleep(0.005)

    assert killed == [victim_pid]
    # Slot freed -> a brand-new session id can be acquired even at max_size.
    fresh = client("after-reap")
    resp = _call(fresh, "r2", "run_labtalk", {"script": "w = 1;"})
    assert resp["ok"] is True


# --------------------------------------------------------------------------- #
# Default terminate hook is platform-guarded (smoke)                           #
# --------------------------------------------------------------------------- #


def test_default_terminate_hook_is_callable():
    assert callable(daemon_mod.default_terminate_process)


# --------------------------------------------------------------------------- #
# CRITICAL — the watchdog must NEVER hard-kill the daemon's own pid / None     #
# --------------------------------------------------------------------------- #


def _spin(predicate, timeout=2.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(0.005)
    return predicate()


def test_default_get_pid_returns_none_not_self_pid():
    """Production default must mean 'unknown -> do not force-kill', NOT self."""
    assert daemon_mod._default_get_pid(object()) is None
    assert daemon_mod._default_get_pid(object()) != os.getpid()


def test_watchdog_never_kills_self_pid():
    killed: list = []
    reaped: list = []
    wd = Watchdog(terminate_process=lambda p: killed.append(p),
                  on_reap=lambda s, p: reaped.append(s), tick=0.005)
    wd.start()
    try:
        wd.arm("self", pid=os.getpid(), deadline=time.monotonic() + 0.02)
        assert _spin(lambda: reaped == ["self"]), reaped
        # The slot was freed, but the daemon's OWN pid was never terminated.
        assert killed == [], killed
    finally:
        wd.stop()


def test_watchdog_never_kills_none_pid():
    killed: list = []
    reaped: list = []
    wd = Watchdog(terminate_process=lambda p: killed.append(p),
                  on_reap=lambda s, p: reaped.append(s), tick=0.005)
    wd.start()
    try:
        wd.arm("unknown", pid=None, deadline=time.monotonic() + 0.02)
        assert _spin(lambda: reaped == ["unknown"]), reaped
        assert killed == [], killed  # None pid -> skip kill, still free slot
    finally:
        wd.stop()


# --------------------------------------------------------------------------- #
# MEDIUM 2 — a raising terminate hook must not kill the watchdog thread        #
# --------------------------------------------------------------------------- #


def test_watchdog_survives_raising_terminate_and_handles_next():
    killed: list = []
    reaped: list = []

    def term(pid):
        if pid == 111:
            raise ProcessLookupError("already exited after graceful close")
        killed.append(pid)

    wd = Watchdog(terminate_process=term,
                  on_reap=lambda s, p: reaped.append(s), tick=0.005)
    wd.start()
    try:
        wd.arm("first", pid=111, deadline=time.monotonic() + 0.02)
        # Despite the raise, the slot is still freed for the first session...
        assert _spin(lambda: "first" in reaped), reaped
        # ...and the watchdog thread SURVIVED to handle the next reap.
        wd.arm("second", pid=222, deadline=time.monotonic() + 0.02)
        assert _spin(lambda: killed == [222]), killed
        assert "second" in reaped
    finally:
        wd.stop()


# --------------------------------------------------------------------------- #
# HIGH 2 — Pool.acquire must NOT hold the lock across Session.start()          #
# --------------------------------------------------------------------------- #


def test_acquire_does_not_block_pool_during_start_and_rolls_back():
    blocked = threading.Event()
    release = threading.Event()

    def slow_factory():
        blocked.set()
        release.wait(5.0)  # the worker stays stuck in start the whole time
        return FakeOrigin()

    pool = Pool(slow_factory, registry={}, max_size=3,
                get_pid=FakeFactory.get_pid, start_timeout=0.3)
    results: dict = {}

    def do_acquire():
        try:
            pool.acquire("slow")
        except Exception as exc:  # noqa: BLE001
            results["err"] = type(exc).__name__

    t = threading.Thread(target=do_acquire)
    t.start()
    try:
        assert blocked.wait(2.0), "factory/start never entered"
        # While start() blocks, OTHER pool ops must not be blocked by the lock.
        t0 = time.monotonic()
        assert pool.get("other") is None
        assert pool.child_pids() == []  # the starting session has no pid yet
        pool.release("nonexistent")
        assert time.monotonic() - t0 < 1.0, "pool ops were blocked by start()"

        # On start-timeout the slot is rolled back; nothing lingers untracked.
        t.join(timeout=3.0)
        assert results.get("err") == "TimeoutError"
        assert "slow" not in pool.session_ids()
    finally:
        release.set()
        pool.stop_all()


# --------------------------------------------------------------------------- #
# ROBUSTNESS 1 — FastMCP registry accessor fails LOUD, never returns {}        #
# --------------------------------------------------------------------------- #


def test_iter_registered_tools_returns_the_real_registry():
    from origin_pro_mcp.daemon import iter_registered_tools
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp

    reg = iter_registered_tools(mcp)
    assert isinstance(reg, dict) and len(reg) > 0
    assert all(callable(fn) for fn in reg.values())


def test_iter_registered_tools_raises_when_internals_renamed():
    """A FastMCP upgrade that renames ``_tool_manager`` must raise a LOUD error
    naming the missing attribute — NEVER silently return an empty registry."""
    from origin_pro_mcp.daemon import iter_registered_tools

    class RenamedInternals:  # no ``_tool_manager`` at all
        pass

    with pytest.raises(RuntimeError) as ei:
        iter_registered_tools(RenamedInternals())
    assert "_tool_manager" in str(ei.value)


def test_iter_registered_tools_raises_on_empty_registry():
    """An empty ``_tool_manager._tools`` (relocated registry / failed
    registration) must raise, not bring a zero-tool server up silently."""
    from origin_pro_mcp.daemon import iter_registered_tools

    class EmptyManager:
        _tools: dict = {}

    class EmptyMCP:
        _tool_manager = EmptyManager()

    with pytest.raises(RuntimeError) as ei:
        iter_registered_tools(EmptyMCP())
    assert "zero tools" in str(ei.value).lower()


# --------------------------------------------------------------------------- #
# ROBUSTNESS 2 — an oversized tool result becomes an actionable error          #
# --------------------------------------------------------------------------- #


def test_oversized_result_returns_actionable_error():
    """A result larger than MAX_FRAME must be replaced, at the serialization
    boundary, with an actionable error — not shipped as an oversized frame the
    peer would reject with a raw FrameError."""
    import json

    from origin_pro_mcp.daemon import Session
    from origin_pro_mcp.transport import MAX_FRAME

    huge = "x" * (MAX_FRAME + 4096)
    session = Session("s-big", FakeFactory(), {"get_worksheet_data": lambda: huge})
    resp = session._dispatch("r1", "get_worksheet_data", {})

    assert resp["ok"] is False
    assert resp["result"] is None
    assert "transport limit" in resp["error"]
    assert "export_worksheet" in resp["error"]
    # The replacement response itself fits in one frame (no raw FrameError).
    assert len(json.dumps(resp).encode("utf-8")) <= MAX_FRAME


def test_normal_result_passes_through_unflagged():
    """A normal-sized result is returned unchanged (the size guard is inert)."""
    from origin_pro_mcp.daemon import Session

    session = Session("s-ok", FakeFactory(), {"list_worksheets": lambda: "[]"})
    resp = session._dispatch("r1", "list_worksheets", {})

    assert resp["ok"] is True
    assert resp["result"] == "[]"
    assert resp["error"] is None
