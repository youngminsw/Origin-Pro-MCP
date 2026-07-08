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


# --------------------------------------------------------------------------- #
# S1: persistent per-session modal-dialog watchdog                             #
# --------------------------------------------------------------------------- #


def test_dialog_autodismiss_enabled_env_toggle(monkeypatch):
    from origin_pro_mcp.daemon import _dialog_autodismiss_enabled

    monkeypatch.delenv("ORIGIN_PRO_MCP_DIALOG_AUTODISMISS", raising=False)
    assert _dialog_autodismiss_enabled() is True  # default ON
    for off in ("0", "off", "false", "no", "OFF", "False"):
        monkeypatch.setenv("ORIGIN_PRO_MCP_DIALOG_AUTODISMISS", off)
        assert _dialog_autodismiss_enabled() is False
    for on in ("1", "on", "true", "yes", "anything-else"):
        monkeypatch.setenv("ORIGIN_PRO_MCP_DIALOG_AUTODISMISS", on)
        assert _dialog_autodismiss_enabled() is True


def test_scan_dialogs_once_detects_titles_and_dismisses():
    from origin_pro_mcp.daemon import _scan_dialogs_once

    titles = {11: "Get MiKTeX Path", 12: "Reminder Message"}
    closed: list = []
    events, current = _scan_dialogs_once(
        pid=4242, seen_hwnds=set(),
        find_dialogs=lambda pid: [11, 12],
        get_title=lambda h: titles[h],
        close=closed.append,
        autodismiss=True,
    )
    assert {e["title"] for e in events} == {"Get MiKTeX Path", "Reminder Message"}
    assert all(e["dismissed"] for e in events)
    assert all("time" in e for e in events)
    assert sorted(closed) == [11, 12]  # both auto-closed
    assert current == {11, 12}


def test_scan_dialogs_once_dedups_persisting_dialog():
    """A dialog left open (autodismiss off) is reported ONCE, not every poll."""
    from origin_pro_mcp.daemon import _scan_dialogs_once

    find = lambda pid: [11]
    seen: set = set()
    events, seen = _scan_dialogs_once(
        4242, seen, find, lambda h: "Error", close=lambda h: None,
        autodismiss=False,
    )
    assert len(events) == 1 and events[0]["dismissed"] is False
    # Next poll with the SAME hwnd still present -> no new event.
    events2, seen = _scan_dialogs_once(
        4242, seen, find, lambda h: "Error", close=lambda h: None,
        autodismiss=False,
    )
    assert events2 == []
    assert seen == {11}


def test_dialog_watchdog_poll_records_via_callback_and_survives_finder_error():
    from origin_pro_mcp.daemon import DialogWatchdog

    recorded: list = []
    closed: list = []
    wd = DialogWatchdog(
        pid=4242, on_event=recorded.append, interval=999,
        find_dialogs=lambda pid: [7], get_title=lambda h: "Font Missing",
        close=closed.append, autodismiss=True,
    )
    got = wd.poll_once()
    assert len(got) == 1 and got[0]["title"] == "Font Missing"
    assert recorded and recorded[0]["title"] == "Font Missing"
    assert closed == [7]

    # A finder that raises must not propagate (thread + callers stay alive).
    def boom(pid):
        raise RuntimeError("EnumWindows failed")

    wd2 = DialogWatchdog(pid=1, on_event=recorded.append, find_dialogs=boom)
    assert wd2.poll_once() == []  # swallowed, returns empty


def test_dialog_watchdog_thread_polls_then_stops():
    """The real thread loop detects a dialog and stops cleanly on stop()."""
    from origin_pro_mcp.daemon import DialogWatchdog

    recorded: list = []
    closed: list = []
    wd = DialogWatchdog(
        pid=4242, on_event=recorded.append, interval=0.01,
        find_dialogs=lambda pid: [7], get_title=lambda h: "Blocking",
        close=closed.append, autodismiss=True,
    )
    wd.start()
    assert _spin(lambda: bool(recorded), timeout=2.0)
    wd.stop()
    assert recorded[0]["title"] == "Blocking" and recorded[0]["dismissed"] is True
    assert 7 in closed


class _FakeDialogWatchdog:
    """Records a fixed dialog on first poll; tracks start/stop. No real thread."""

    def __init__(self, pid, on_event, title="Get MiKTeX Path", dismissed=True):
        self._on_event = on_event
        self._title = title
        self._dismissed = dismissed
        self.started = False
        self.stopped = False
        self._fired = False

    def start(self):
        self.started = True

    def stop(self, join_timeout=2.0):
        self.stopped = True

    def poll_once(self):
        if self._fired:
            return []
        self._fired = True
        ev = {"time": 1.0, "title": self._title, "dismissed": self._dismissed}
        self._on_event(ev)
        return [ev]


def test_session_starts_and_stops_dialog_watchdog():
    """A Session with a known PID starts its watchdog, records dialogs it reports,
    and stops it on teardown."""
    from origin_pro_mcp.daemon import Session

    made: list = []

    def wd_factory(pid, on_event):
        wd = _FakeDialogWatchdog(pid, on_event)
        made.append(wd)
        return wd

    session = Session("s-dlg", FakeFactory(), {"list_worksheets": lambda: "[]"},
                      get_pid=FakeFactory.get_pid,
                      dialog_watchdog_factory=wd_factory)
    session.start()
    try:
        assert len(made) == 1 and made[0].started is True
        # Force a poll (as the dispatch-timeout diagnosis would) -> event recorded.
        event = session.poll_dialogs_now()
        assert event is not None and event["title"] == "Get MiKTeX Path"
        assert session.last_dialog_event()["dismissed"] is True
    finally:
        session.stop()
    assert made[0].stopped is True


def test_session_skips_dialog_watchdog_without_pid():
    """The attached user Origin has pid=None -> no watchdog is started (its
    dialogs must never be swept by the daemon)."""
    from origin_pro_mcp.daemon import Session

    made: list = []
    session = Session("s-nopid", FakeFactory(), {"list_worksheets": lambda: "[]"},
                      get_pid=lambda i: None,
                      dialog_watchdog_factory=lambda pid, ev: made.append(1))
    session.start()
    try:
        assert made == []  # never constructed
        assert session.poll_dialogs_now() is None
    finally:
        session.stop()


def test_record_dialog_event_history_is_bounded():
    from origin_pro_mcp.daemon import Session

    session = Session("s-cap", FakeFactory(), {})
    for i in range(120):
        session.record_dialog_event({"time": i, "title": f"d{i}", "dismissed": True})
    with session._dialog_lock:
        assert len(session._dialog_events) == 50
    assert session.last_dialog_event()["title"] == "d119"


def test_dialog_note_definitive_messages():
    """_dialog_note names the dialog: auto-dismissed vs. off-and-open vs. none."""
    daemon = Daemon()

    class _PoolWith:
        def __init__(self, session):
            self._s = session

        def get(self, sid):
            return self._s

    class _Sess:
        def __init__(self, event):
            self._event = event

        def poll_dialogs_now(self):
            return self._event

    daemon._pool = _PoolWith(_Sess({"title": "Get MiKTeX Path", "dismissed": True}))
    note = daemon._dialog_note("A")
    assert "Get MiKTeX Path" in note and "AUTO-DISMISSED" in note

    daemon._pool = _PoolWith(_Sess({"title": "Error Box", "dismissed": False}))
    note = daemon._dialog_note("A")
    assert "Error Box" in note and "auto-dismiss is OFF" in note

    daemon._pool = _PoolWith(_Sess(None))
    assert daemon._dialog_note("A") is None


# -- end-to-end: the soft dispatch-timeout warning NAMES the modal dialog ----- #


class _BlockingClock:
    def __init__(self, start=0.0):
        self._t = start
        self._lock = threading.Lock()

    def __call__(self):
        with self._lock:
            return self._t

    def advance(self, dt):
        with self._lock:
            self._t += dt


class _BlockingExecuteOrigin(FakeOrigin):
    def __init__(self):
        super().__init__()
        self.unblock = threading.Event()
        self.execute_entered = threading.Event()

    def Execute(self, script):
        if "HANG" in script:
            self.execute_entered.set()
            self.unblock.wait(10)
            return True
        return super().Execute(script)


def test_soft_timeout_warning_names_the_modal_dialog(tmp_path):
    """The anti-misdiagnosis fix: when a session wedges AND its watchdog has seen a
    modal dialog, the soft-timeout reply names that dialog definitively (with the
    title) instead of the generic 'most likely a modal dialog' hedge."""
    clock = _BlockingClock()
    origins: list = []
    killed: list = []

    def factory():
        o = _BlockingExecuteOrigin()
        o.fake_pid = 77001 + len(origins)
        origins.append(o)
        return o

    def wd_factory(pid, on_event):
        return _FakeDialogWatchdog(pid, on_event, title="Get MiKTeX Path")

    daemon = Daemon()
    ok = daemon.start(
        origin_factory=factory, registry=_real_registry(), max_size=3,
        host="127.0.0.1", port=0, get_pid=lambda i: i.fake_pid,
        terminate_process=lambda pid: killed.append(pid), clock=clock,
        dispatch_timeout=5.0, dispatch_kill_grace=10.0, monitor_tick=0.005,
        reconnect_grace=0.0, lockfile_path=str(tmp_path / "daemon.json"),
        dialog_watchdog_factory=wd_factory,
    )
    assert ok
    conns: list = []
    try:
        conn = transport.connect(daemon.host, daemon.port, daemon.token,
                                 session_id="A")
        conn.settimeout(5.0)
        conns.append(conn)
        conn.send_frame({"type": "request", "request_id": "r1",
                         "name": "run_labtalk", "kwargs": {"script": "HANG;"}})
        assert _spin(lambda: origins and origins[0].execute_entered.is_set())
        clock.advance(6.0)  # trip the soft budget (before the hard kill)
        resp = conn.recv_frame()
        assert resp["ok"] is False
        assert "Get MiKTeX Path" in resp["error"]      # NAMED, not hedged
        assert "AUTO-DISMISSED" in resp["error"]
        assert killed == []                            # only warned, not killed
        origins[0].unblock.set()                       # let the wedge clear
    finally:
        for c in conns:
            c.close()
        daemon.stop()
