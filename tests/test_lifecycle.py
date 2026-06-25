"""Rigorous, COM-free lifecycle / orphan self-cleanup tests (Phase 3d).

Everything here runs against an injected :class:`FakeOrigin` factory over real
loopback sockets, with a CONTROLLABLE fake clock and an INJECTED terminate hook
so every timeout (graceful reap, watchdog force-kill, heartbeat gap, idle exit)
fires deterministically — no real 5s/30s/10min waits, no COM.

The headline invariant proven below: slot reclamation NEVER waits on a wedged
worker — a session whose graceful save/close blocks forever is still reclaimed
out-of-band by the watchdog killing the recorded PID.
"""
from __future__ import annotations

import os
import threading
import time

import pytest

from fakes import FakeOrigin, ThreadGuardedFake

from origin_pro_mcp import transport
from origin_pro_mcp.daemon import Daemon, DEFAULT_RECONNECT_GRACE, recovery_path, write_lockfile


# --------------------------------------------------------------------------- #
# Deterministic time + helpers                                                 #
# --------------------------------------------------------------------------- #


class FakeClock:
    """A monotonic clock the test drives by hand (thread-safe)."""

    def __init__(self, start: float = 0.0):
        self._t = start
        self._lock = threading.Lock()

    def __call__(self) -> float:
        with self._lock:
            return self._t

    def advance(self, dt: float) -> None:
        with self._lock:
            self._t += dt


def _wait_until(predicate, timeout: float = 3.0, interval: float = 0.005):
    """Poll ``predicate`` in bounded REAL time (deadlines are fake-clock based)."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(interval)
    return predicate()


def _real_registry() -> dict:
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp

    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


class FakeFactory:
    """Hands out one fresh instance per session, each with a unique fake pid."""

    def __init__(self, make=FakeOrigin):
        self.created: list = []
        self._make = make
        self._next_pid = 70000
        self._lock = threading.Lock()

    def __call__(self):
        with self._lock:
            inst = self._make()
            inst.fake_pid = self._next_pid
            self._next_pid += 1
            self.created.append(inst)
            return inst

    @staticmethod
    def get_pid(instance) -> int:
        return instance.fake_pid


class BlockingSaveOrigin(FakeOrigin):
    """A FakeOrigin whose graceful ``Save`` BLOCKS FOREVER (wedged worker)."""

    def __init__(self):
        super().__init__()
        self.release = threading.Event()
        self.save_entered = threading.Event()

    def Save(self, path):
        self.saved_paths.append(path)
        self.save_entered.set()
        self.release.wait()  # never returns until the test releases it
        return self.save_result


@pytest.fixture
def lifecycle():
    """Start/teardown daemons with injected clock + terminate hook."""
    daemons: list[Daemon] = []
    conns: list = []

    def start(factory, *, clock=None, reap_grace=5.0,
              heartbeat_reap_after=30.0, idle_exit_after=600.0,
              reconnect_grace=0.0,
              project_path_getter=None, recovery_dir=None, get_pid=None,
              terminate=None, is_alive=None, monitor_tick=0.005,
              lockfile_path=None, max_size=3):
        killed: list[int] = []
        term = terminate or (lambda pid: killed.append(pid))
        daemon = Daemon()
        ok = daemon.start(
            origin_factory=factory,
            registry=_real_registry(),
            max_size=max_size,
            host="127.0.0.1",
            port=0,
            get_pid=get_pid,
            terminate_process=term,
            clock=clock,
            reap_grace=reap_grace,
            heartbeat_reap_after=heartbeat_reap_after,
            idle_exit_after=idle_exit_after,
            reconnect_grace=reconnect_grace,
            recovery_dir=recovery_dir,
            project_path_getter=project_path_getter,
            is_alive=is_alive,
            monitor_tick=monitor_tick,
            lockfile_path=lockfile_path,
        )
        assert ok is True
        daemons.append(daemon)
        return daemon, killed

    def client(daemon, session_id):
        conn = transport.connect(daemon.host, daemon.port, daemon.token,
                                 session_id=session_id)
        conn.settimeout(5.0)
        conns.append(conn)
        return conn

    try:
        yield start, client
    finally:
        for conn in conns:
            conn.close()
        for daemon in daemons:
            daemon.stop()


def _request(conn, request_id, name, kwargs):
    conn.send_frame({"type": "request", "request_id": request_id,
                     "name": name, "kwargs": kwargs})
    return conn.recv_frame()


# --------------------------------------------------------------------------- #
# recovery_path scheme (collision-safe, session-id namespaced)                 #
# --------------------------------------------------------------------------- #


def test_recovery_path_named_project_is_session_namespaced(tmp_path):
    p = recovery_path(str(tmp_path), "abc123", str(tmp_path / "myproj.opju"))
    assert p == str(tmp_path / "myproj.abc123.recover.opju")


def test_recovery_path_no_project_uses_session_id(tmp_path):
    p = recovery_path(str(tmp_path), "abc123", None)
    assert p == str(tmp_path / "abc123.recover.opju")


def test_recovery_path_never_overwrites_increments_counter(tmp_path):
    stem = str(tmp_path / "myproj.opju")
    first = recovery_path(str(tmp_path), "s", stem)
    open(first, "w").write("FIRST")               # the file now exists on disk
    second = recovery_path(str(tmp_path), "s", stem)
    assert second == str(tmp_path / "myproj.s.recover.1.opju")
    open(second, "w").write("SECOND")
    third = recovery_path(str(tmp_path), "s", stem)
    assert third == str(tmp_path / "myproj.s.recover.2.opju")
    # The original recovery file was never touched.
    assert open(first).read() == "FIRST"


# --------------------------------------------------------------------------- #
# GRACEFUL REAP — connection close saves a recovery file, frees the slot       #
# --------------------------------------------------------------------------- #


def test_graceful_reap_saves_recovery_file_and_frees_slot(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory()
    rec_dir = str(tmp_path / "recovery")
    proj = str(tmp_path / "myproj.opju")
    daemon, killed = start(
        factory, clock=clock, reap_grace=5.0, get_pid=FakeFactory.get_pid,
        recovery_dir=rec_dir, project_path_getter=lambda inst: proj,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "graceful")
    resp = _request(conn, "r0", "run_labtalk", {"script": "g = 1;"})
    assert resp["ok"] is True
    assert daemon.pool.session_ids() == ["graceful"]

    # Closing the connection (clean EOF) schedules the reap.
    conn.close()

    expected = os.path.join(rec_dir, "myproj.graceful.recover.opju")
    assert _wait_until(lambda: expected in factory.created[0].saved_paths), \
        factory.created[0].saved_paths
    # Slot freed, and the graceful path never invoked the force-kill.
    assert _wait_until(lambda: daemon.pool.session_ids() == [])
    assert killed == []


# --------------------------------------------------------------------------- #
# WATCHDOG REAP (C1, the headline) — wedged worker, slot reclaimed out-of-band #
# --------------------------------------------------------------------------- #


def test_watchdog_force_kills_wedged_session_without_waiting(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory(make=BlockingSaveOrigin)
    daemon, killed = start(
        factory, clock=clock, reap_grace=10.0, get_pid=FakeFactory.get_pid,
        recovery_dir=str(tmp_path / "recovery"),
        project_path_getter=lambda inst: None,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "wedged")
    resp = _request(conn, "r0", "run_labtalk", {"script": "w = 1;"})
    assert resp["ok"] is True
    wedged = factory.created[0]
    fake_pid = wedged.fake_pid
    assert fake_pid in daemon.pool.child_pids()

    try:
        # Trigger the reap; stage 1 enters Save and BLOCKS forever.
        daemon.reap_session("wedged")
        assert wedged.save_entered.wait(timeout=2.0), "graceful save never ran"
        # Still wedged: nothing reclaimed yet, no kill yet.
        assert killed == []
        assert daemon.pool.session_ids() == ["wedged"]

        # Push the fake clock past reap_grace -> the watchdog must fire.
        clock.advance(11.0)

        assert _wait_until(lambda: killed == [fake_pid]), killed
        # Slot reclaimed WITHOUT waiting on the (still-blocked) worker.
        assert _wait_until(lambda: daemon.pool.session_ids() == [])
        assert fake_pid not in daemon.pool.child_pids()
        # The worker is provably still stuck inside Save (never released).
        assert not wedged.release.is_set()
    finally:
        wedged.release.set()  # let the abandoned worker thread unwind


# --------------------------------------------------------------------------- #
# COLLISION — a pre-existing recovery file is NEVER overwritten                #
# --------------------------------------------------------------------------- #


def test_reap_never_overwrites_existing_recovery_file(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory()
    rec_dir = str(tmp_path / "recovery")
    os.makedirs(rec_dir, exist_ok=True)
    proj = str(tmp_path / "proj.opju")

    # A recovery file from a prior (crashed) daemon already exists on disk.
    preexisting = os.path.join(rec_dir, "proj.dup.recover.opju")
    with open(preexisting, "w") as fh:
        fh.write("PRIOR-WORK")

    daemon, _killed = start(
        factory, clock=clock, get_pid=FakeFactory.get_pid,
        recovery_dir=rec_dir, project_path_getter=lambda inst: proj,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "dup")
    _request(conn, "r0", "run_labtalk", {"script": "d = 1;"})
    conn.close()

    target = os.path.join(rec_dir, "proj.dup.recover.1.opju")
    assert _wait_until(lambda: target in factory.created[0].saved_paths), \
        factory.created[0].saved_paths
    # The reap chose the counter-suffixed path; the prior file is untouched.
    assert preexisting not in factory.created[0].saved_paths
    assert open(preexisting).read() == "PRIOR-WORK"


# --------------------------------------------------------------------------- #
# HEARTBEAT-GAP REAP — a silent session is reaped (half-open detection)        #
# --------------------------------------------------------------------------- #


def test_heartbeat_gap_triggers_reap(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory()
    daemon, killed = start(
        factory, clock=clock, heartbeat_reap_after=30.0, reap_grace=5.0,
        get_pid=FakeFactory.get_pid, recovery_dir=str(tmp_path / "recovery"),
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "silent")
    _request(conn, "r0", "run_labtalk", {"script": "s = 1;"})
    assert daemon.pool.session_ids() == ["silent"]

    # No heartbeats arrive; push past heartbeat_reap_after. The connection stays
    # open, so ONLY the heartbeat backstop (not EOF) can drive this reap.
    clock.advance(31.0)

    assert _wait_until(lambda: daemon.pool.session_ids() == [])
    assert factory.created[0].saved_paths, "reap should have saved a recovery file"
    assert killed == []  # graceful path; no force-kill needed


# --------------------------------------------------------------------------- #
# IDLE SELF-EXIT — 0 sessions past idle_exit_after -> daemon shuts down        #
# --------------------------------------------------------------------------- #


def test_idle_self_exit_shuts_down_and_removes_lockfile(lifecycle, tmp_path):
    start, _client = lifecycle
    clock = FakeClock()
    lockfile = str(tmp_path / "daemon.json")
    daemon, _killed = start(
        FakeFactory(), clock=clock, idle_exit_after=100.0,
        recovery_dir=str(tmp_path / "recovery"), lockfile_path=lockfile,
    )
    assert os.path.exists(lockfile)
    assert daemon.running is True

    clock.advance(101.0)  # 0 sessions, idle past the limit -> self-exit

    assert _wait_until(lambda: daemon.running is False)
    assert _wait_until(lambda: not os.path.exists(lockfile))
    assert daemon._stopped_event.is_set()


# --------------------------------------------------------------------------- #
# STARTUP SWEEP — orphans from a crashed prior daemon are killed on start      #
# --------------------------------------------------------------------------- #


def test_startup_sweep_kills_surviving_child_pids(lifecycle, tmp_path):
    start, _client = lifecycle
    lockfile = str(tmp_path / "daemon.json")
    # A stale lockfile left by a crashed daemon, recording a still-alive child.
    write_lockfile(lockfile, port=1, token="stale", pid=4242,
                   child_pids=[55555, 66666])

    killed: list[int] = []

    def is_alive(pid):
        return pid == 55555  # 66666 already dead -> must NOT be killed

    daemon, _ = start(
        FakeFactory(), is_alive=is_alive,
        terminate=lambda pid: killed.append(pid),
        recovery_dir=str(tmp_path / "recovery"), lockfile_path=lockfile,
    )

    # The survivor was force-killed; the dead pid was skipped.
    assert killed == [55555]
    # The new lockfile replaced the stale one (fresh pid, no stale children).
    from origin_pro_mcp.daemon import read_lockfile

    data = read_lockfile(lockfile)
    assert data["pid"] == os.getpid()
    assert data["child_pids"] == []


# --------------------------------------------------------------------------- #
# COM-INVARIANT — watchdog + monitor threads NEVER touch a COM proxy           #
# --------------------------------------------------------------------------- #


def test_monitors_never_touch_com_proxy(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()

    guards: list[ThreadGuardedFake] = []

    def make_guarded():
        g = ThreadGuardedFake(threading.get_ident())  # owner == worker thread
        guards.append(g)  # storing the ref only; no attribute access
        return g

    factory = FakeFactory(make=make_guarded)
    daemon, killed = start(
        factory, clock=clock, heartbeat_reap_after=10.0, reap_grace=2.0,
        get_pid=lambda inst: 13579,  # never dereferences the proxy
        recovery_dir=str(tmp_path / "recovery"),
        project_path_getter=lambda inst: None,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "guarded")
    _request(conn, "r0", "run_labtalk", {"script": "x = 1;"})
    assert _wait_until(lambda: daemon.pool.session_ids() == ["guarded"])

    # (a) Fire the watchdog directly with a bare PID — proves it never needs the
    #     proxy to reclaim a slot.
    daemon.watchdog.arm("manual", pid=98765, deadline=clock() + 1.0)
    clock.advance(2.0)
    assert _wait_until(lambda: 98765 in killed), killed

    # (b) Drive a heartbeat-gap reap through the MONITOR thread.
    clock.advance(20.0)
    assert _wait_until(lambda: daemon.pool.session_ids() == [])

    # Neither the watchdog nor the monitor ever dereferenced the COM proxy.
    assert len(guards) == 1
    assert guards[0].touched is False


# --------------------------------------------------------------------------- #
# Determinism guard: a fresh daemon per test, no real long sleeps anywhere.    #
# --------------------------------------------------------------------------- #


# --------------------------------------------------------------------------- #
# CRITICAL — a reap whose pid is unknown (None) must NOT self-kill the daemon  #
# --------------------------------------------------------------------------- #


def test_reap_with_unknown_pid_skips_force_kill_but_frees_slot(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory(make=BlockingSaveOrigin)
    daemon, killed = start(
        factory, clock=clock, reap_grace=2.0,
        get_pid=lambda inst: None,  # production-like: unknown child pid
        recovery_dir=str(tmp_path / "recovery"),
        project_path_getter=lambda inst: None,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "unknown")
    assert _request(conn, "r0", "run_labtalk", {"script": "u = 1;"})["ok"] is True
    wedged = factory.created[0]
    # An unknown pid must NOT be recorded as os.getpid() in the lockfile.
    assert os.getpid() not in daemon.pool.child_pids()

    try:
        daemon.reap_session("unknown")
        assert wedged.save_entered.wait(timeout=2.0), "graceful save never ran"
        clock.advance(3.0)  # push past reap_grace -> watchdog fires

        # Slot reclaimed out-of-band, but the daemon's own pid was never killed.
        assert _wait_until(lambda: daemon.pool.session_ids() == [])
        assert killed == [], killed
        assert os.getpid() not in killed
    finally:
        wedged.release.set()


# --------------------------------------------------------------------------- #
# HIGH 1 — connection-reap-vs-reconnect race must reuse the SAME session       #
# --------------------------------------------------------------------------- #


def test_reconnect_within_grace_reuses_session(lifecycle, tmp_path):
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory()
    daemon, killed = start(
        factory, clock=clock, reconnect_grace=10.0, reap_grace=5.0,
        get_pid=FakeFactory.get_pid, recovery_dir=str(tmp_path / "recovery"),
        project_path_getter=lambda inst: None,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "S")
    assert _request(conn, "r0", "run_labtalk", {"script": "keep = 1;"})["ok"]
    assert _wait_until(lambda: daemon.pool.session_ids() == ["S"])
    inst = factory.created[0]

    # Drop the connection, then immediately reconnect with the SAME session id
    # (within the reconnect grace, clock not advanced past the deadline).
    conn.close()
    conn2 = client(daemon, "S")
    assert _request(conn2, "r1", "run_labtalk", {"script": "more = 2;"})["ok"]

    # The SAME instance was reused (state preserved); no reap completed, no kill.
    assert len(factory.created) == 1, "a reap destroyed and respawned the session"
    assert daemon.pool.session_ids() == ["S"]
    assert inst.executed == ["keep = 1;", "more = 2;"]
    assert inst.saved_paths == []  # no graceful save -> no reap ran
    assert killed == []


def test_reconnect_after_commit_still_gets_a_response(lifecycle, tmp_path):
    """If the reconnect lands AFTER the reap committed (stage-1 started), the
    request must still get an actionable response — never a lost reply/stall."""
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory()
    daemon, killed = start(
        factory, clock=clock, reconnect_grace=0.0,  # commit on close
        get_pid=FakeFactory.get_pid, recovery_dir=str(tmp_path / "recovery"),
        project_path_getter=lambda inst: None,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "S")
    assert _request(conn, "r0", "run_labtalk", {"script": "x = 1;"})["ok"]
    conn.close()
    # The reap commits and completes (graceful, FakeOrigin) -> slot freed.
    assert _wait_until(lambda: daemon.pool.session_ids() == [])

    # Reconnecting with the same id gets a FRESH session and a real response.
    conn2 = client(daemon, "S")
    resp = _request(conn2, "r1", "run_labtalk", {"script": "y = 2;"})
    assert resp["type"] == "response"
    assert resp["request_id"] == "r1"
    assert resp["ok"] is True


# --------------------------------------------------------------------------- #
# MEDIUM 1 — any handling error becomes an actionable response (never a hang)  #
# --------------------------------------------------------------------------- #


def test_factory_failure_returns_error_response_not_hang(lifecycle, tmp_path):
    def boom_factory():
        raise RuntimeError("DispatchEx failed: Origin not licensed")

    start, client = lifecycle
    daemon, _killed = start(
        boom_factory, recovery_dir=str(tmp_path / "recovery"),
        lockfile_path=str(tmp_path / "daemon.json"),
    )
    conn = client(daemon, "boom")
    resp = _request(conn, "r0", "run_labtalk", {"script": "x = 1;"})
    assert resp["type"] == "response"
    assert resp["request_id"] == "r0"
    assert resp["ok"] is False
    assert resp["error"]  # an actionable, non-empty error message
    assert "DispatchEx failed" in resp["error"]


def test_unknown_tool_returns_actionable_error(lifecycle, tmp_path):
    start, client = lifecycle
    daemon, _killed = start(
        FakeFactory(), get_pid=FakeFactory.get_pid,
        recovery_dir=str(tmp_path / "recovery"),
        lockfile_path=str(tmp_path / "daemon.json"),
    )
    conn = client(daemon, "u")
    resp = _request(conn, "r0", "no_such_tool", {})
    assert resp["ok"] is False
    assert "no_such_tool" in resp["error"]


# --------------------------------------------------------------------------- #
# LOW — reaping a None session must still pop _last_seen (no unbounded growth) #
# --------------------------------------------------------------------------- #


def test_reap_unknown_session_pops_last_seen(lifecycle, tmp_path):
    start, _client = lifecycle
    clock = FakeClock()
    daemon, _killed = start(
        FakeFactory(), clock=clock, recovery_dir=str(tmp_path / "recovery"),
        lockfile_path=str(tmp_path / "daemon.json"),
    )
    # A heartbeat-only session id that never created a pool session.
    daemon._last_seen["ghost"] = clock()
    daemon.reap_session("ghost")  # session is None -> early return
    assert "ghost" not in daemon._last_seen


# --------------------------------------------------------------------------- #
# HIGH-1 follow-up — production default reconnect_grace is ON (3 s)           #
# --------------------------------------------------------------------------- #


def test_production_default_grace_is_3s_and_prevents_premature_reap(lifecycle, tmp_path):
    """The production default (DEFAULT_RECONNECT_GRACE == 3.0) must:
    1. Be the actual default of Daemon.start().
    2. Cancel a pending reap when the same session reconnects within the window.
    3. NOT prevent reaping: once the grace elapses, the slot is freed.
    """
    import inspect

    # — structural check: Daemon.start uses the constant as its default —
    assert DEFAULT_RECONNECT_GRACE == 3.0
    sig = inspect.signature(Daemon.start)
    assert sig.parameters["reconnect_grace"].default == DEFAULT_RECONNECT_GRACE

    # — behavioural check with a fake clock at the production-default value —
    start, client = lifecycle
    clock = FakeClock()
    factory = FakeFactory()
    daemon, killed = start(
        factory, clock=clock,
        reconnect_grace=DEFAULT_RECONNECT_GRACE,  # ← same value as production
        reap_grace=5.0, get_pid=FakeFactory.get_pid,
        recovery_dir=str(tmp_path / "recovery"),
        project_path_getter=lambda inst: None,
        lockfile_path=str(tmp_path / "daemon.json"),
    )

    conn = client(daemon, "prd")
    assert _request(conn, "r0", "run_labtalk", {"script": "x = 1;"})["ok"]
    inst = factory.created[0]

    # Drop connection; reap is scheduled but NOT committed yet (clock at 0).
    conn.close()
    time.sleep(0.05)  # give monitor a tick

    # Reconnect within grace (clock still at 0 < 3.0 deadline) → same session.
    conn2 = client(daemon, "prd")
    assert _request(conn2, "r1", "run_labtalk", {"script": "y = 2;"})["ok"]
    assert len(factory.created) == 1, "reap fired prematurely; session was replaced"
    assert inst.executed == ["x = 1;", "y = 2;"]
    assert inst.saved_paths == []      # no graceful save → no reap ran
    assert killed == []

    # Drop the second connection; grace elapses → reap MUST now commit and free.
    conn2.close()
    time.sleep(0.05)
    clock.advance(DEFAULT_RECONNECT_GRACE + 0.1)   # past the deadline
    assert _wait_until(lambda: daemon.pool.session_ids() == []), \
        f"slot not freed after grace elapsed; sessions={daemon.pool.session_ids()}"


def test_two_independent_daemons_do_not_interfere(lifecycle, tmp_path):
    start, client = lifecycle
    c1, c2 = FakeClock(), FakeClock()
    f1, f2 = FakeFactory(), FakeFactory()
    d1, _ = start(f1, clock=c1, get_pid=FakeFactory.get_pid,
                  recovery_dir=str(tmp_path / "r1"),
                  lockfile_path=str(tmp_path / "d1.json"))
    d2, _ = start(f2, clock=c2, get_pid=FakeFactory.get_pid,
                  recovery_dir=str(tmp_path / "r2"),
                  lockfile_path=str(tmp_path / "d2.json"))
    a = client(d1, "a")
    b = client(d2, "b")
    assert _request(a, "ra", "run_labtalk", {"script": "a = 1;"})["ok"] is True
    assert _request(b, "rb", "run_labtalk", {"script": "b = 1;"})["ok"] is True
    assert d1.pool.session_ids() == ["a"]
    assert d2.pool.session_ids() == ["b"]
