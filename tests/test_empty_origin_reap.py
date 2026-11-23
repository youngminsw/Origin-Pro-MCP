"""Detach reap closes an EMPTY Origin (and only an empty one).

Origin 2018 launched over COM (-Embedding) ignores Exit, and the default detach
reap used to leave it running with no controller — the user then cannot close
it ("controlled by another program"). ``Session._close_if_empty`` closes it
only when it provably holds nothing; everything uncertain keeps it. No real
PID is ever signalled: terminate/liveness hooks are injected.
"""
import threading
import time

import pytest

from fakes import FakeOrigin
from origin_pro_mcp.daemon import Pool, Session


class FakeProcess:
    """A fake OS process for one Origin: Exit ends it unless ``ignores_exit``."""

    def __init__(self, pid=4242, ignores_exit=False, survives_kill=False):
        self.pid = pid
        self.running = True
        self.ignores_exit = ignores_exit
        self.survives_kill = survives_kill
        self.killed = []

    def is_running(self, pid):
        assert pid == self.pid
        return self.running

    def terminate(self, pid):
        self.killed.append(pid)
        if not self.survives_kill:
            self.running = False


class EmptyOrigin(FakeOrigin):
    """A FakeOrigin with no windows that records Exit and ends its process."""

    def __init__(self, process=None):
        super().__init__()
        self.books, self.graphs, self.matrices = [], [], []
        self.exited = False
        self.process = process

    def Exit(self):
        self.exited = True
        if self.process is not None and not self.process.ignores_exit:
            self.process.running = False


class _NoFactory:
    def __call__(self):  # never used: tests set .instance directly
        raise AssertionError("factory should not be called")


def _session(instance, process, *, pid="from-process", attached=False):
    s = Session("s-empty", _NoFactory(), {},
                terminate_process=process.terminate,
                process_running=process.is_running,
                empty_close_wait=0.05)
    s.instance = instance
    s.pid = process.pid if pid == "from-process" else pid
    s.attached = attached
    return s


@pytest.fixture(autouse=True)
def _detach_default(monkeypatch):
    monkeypatch.delenv("ORIGIN_PRO_MCP_REAP_CLOSE", raising=False)


def test_empty_origin_is_closed_on_detach_reap(tmp_path):
    proc = FakeProcess()
    inst = EmptyOrigin(proc)
    s = _session(inst, proc)
    s._graceful_reap(str(tmp_path), None)
    assert inst.exited is True
    assert inst.executed[0] == "doc -s;"      # marked clean BEFORE Exit
    assert proc.killed == []                    # Exit took; no force-kill
    assert inst.saved_paths == []               # never saved
    assert s.empty_close_result == "closed"


def test_force_kills_when_exit_is_ignored(tmp_path):
    proc = FakeProcess(ignores_exit=True)       # Origin 2018 -Embedding
    inst = EmptyOrigin(proc)
    s = _session(inst, proc)
    s._graceful_reap(str(tmp_path), None)
    assert inst.exited is True
    assert proc.killed == [proc.pid]
    assert proc.running is False
    assert s.empty_close_result == "forced"


def test_reports_still_running_when_kill_does_not_take():
    proc = FakeProcess(ignores_exit=True, survives_kill=True)
    s = _session(EmptyOrigin(proc), proc)
    assert s._close_if_empty() == "still-running"
    assert proc.killed == [proc.pid]


def test_origin_with_content_is_kept(tmp_path):
    proc = FakeProcess()
    inst = FakeOrigin()                          # Book1 + Graph1
    inst.exited = False
    inst.Exit = lambda: setattr(inst, "exited", True)
    s = _session(inst, proc)
    s._graceful_reap(str(tmp_path), None)
    assert inst.exited is False
    assert inst.executed == []                  # not even `doc -s;`
    assert proc.killed == []
    assert s.empty_close_result == "kept"


@pytest.mark.parametrize("which", ["books", "graphs", "matrices"])
def test_any_single_window_keeps_origin(which):
    from fakes import FakeBook, FakeGraph, FakeMatrix

    proc = FakeProcess()
    inst = EmptyOrigin(proc)
    setattr(inst, which, [{"books": FakeBook("B"), "graphs": FakeGraph("G"),
                           "matrices": FakeMatrix("M")}[which]])
    s = _session(inst, proc)
    assert s._close_if_empty() == "kept"
    assert inst.exited is False


@pytest.mark.parametrize("broken", ["WorksheetPages", "GraphPages", "MatrixPages"])
def test_count_read_failure_keeps_origin(broken):
    proc = FakeProcess()

    class Unreadable(EmptyOrigin):
        pass

    def boom(self):
        raise RuntimeError("COM error")

    setattr(Unreadable, broken, property(boom))
    inst = Unreadable(proc)
    s = _session(inst, proc)
    assert s._close_if_empty() == "kept"
    assert inst.exited is False
    assert proc.killed == []


def test_non_integer_count_keeps_origin():
    proc = FakeProcess()

    class Weird(EmptyOrigin):
        @property
        def MatrixPages(self):
            class P:
                Count = "n/a"
            return P()

    inst = Weird(proc)
    assert _session(inst, proc)._close_if_empty() == "kept"
    assert inst.exited is False


def test_other_pages_keep_origin_when_exposed():
    # e.g. a project holding only a Notes window: the three counts are 0, but
    # a readable all-pages count says otherwise -> keep.
    proc = FakeProcess()

    class WithNotes(EmptyOrigin):
        @property
        def Pages(self):
            class P:
                Count = 1
            return P()

    inst = WithNotes(proc)
    assert _session(inst, proc)._close_if_empty() == "kept"
    assert inst.exited is False


def test_attach_mode_keeps_origin(tmp_path):
    proc = FakeProcess()
    inst = EmptyOrigin(proc)
    s = _session(inst, proc, attached=True)
    s._graceful_reap(str(tmp_path), None)
    assert inst.exited is False
    assert proc.killed == []
    assert s.empty_close_result == "kept"


def test_unknown_pid_sends_exit_but_never_force_kills():
    proc = FakeProcess(ignores_exit=True)
    inst = EmptyOrigin(proc)
    s = _session(inst, proc, pid=None)
    assert s._close_if_empty() == "exit-sent"
    assert inst.exited is True
    assert proc.killed == []


def test_exception_in_close_never_breaks_the_reap(tmp_path):
    proc = FakeProcess()

    class ExitRaises(EmptyOrigin):
        def Exit(self):
            raise RuntimeError("RPC server unavailable")

    s = _session(ExitRaises(proc), proc)
    s._graceful_reap(str(tmp_path), None)      # must not raise
    # Exit failed -> process still running -> fallback kill by known PID
    assert proc.killed == [proc.pid]


def test_reap_close_opt_in_path_is_unchanged(tmp_path, monkeypatch):
    """ORIGIN_PRO_MCP_REAP_CLOSE=1 keeps the legacy save-and-close: Save a
    recovery copy, then Exit — no emptiness check, no PID force-kill here."""
    monkeypatch.setenv("ORIGIN_PRO_MCP_REAP_CLOSE", "1")
    proc = FakeProcess(ignores_exit=True)
    inst = EmptyOrigin(proc)
    s = _session(inst, proc)
    s._graceful_reap(str(tmp_path / "rec"), None)
    assert inst.exited is True
    assert len(inst.saved_paths) == 1
    assert proc.killed == []
    assert s.empty_close_result is None         # _close_if_empty not involved


# --------------------------------------------------------------------------- #
# Wiring: Pool passes the injected hooks to its sessions                       #
# --------------------------------------------------------------------------- #


def test_pool_wires_hooks_and_closes_empty_session_origin():
    proc = FakeProcess(ignores_exit=True)
    made = []

    def factory():
        inst = EmptyOrigin(proc)
        made.append(inst)
        return inst

    pool = Pool(factory, {}, max_size=1, get_pid=lambda inst: proc.pid,
                terminate_process=proc.terminate,
                process_running=proc.is_running)
    session = pool.acquire("p1")
    session._empty_close_wait = 0.05
    done = threading.Event()
    session.submit_reap("unused-dir", None, on_done=done.set)
    assert done.wait(5.0)
    assert made[0].exited is True
    assert proc.killed == [proc.pid]
    pool.discard("p1")


def test_pool_attached_session_is_kept():
    proc = FakeProcess()
    user_origin = EmptyOrigin(proc)
    pool = Pool(lambda: EmptyOrigin(FakeProcess(pid=1)), {}, max_size=2,
                attach_factory=lambda: user_origin,
                attach_get_pid=lambda inst: None,
                terminate_process=proc.terminate,
                process_running=proc.is_running)
    session = pool.acquire("att", attach=True)
    assert session.attached is True
    done = threading.Event()
    session.submit_reap("unused-dir", None, on_done=done.set)
    assert done.wait(5.0)
    assert user_origin.exited is False
    assert proc.killed == []
    pool.discard("att")


def test_wait_is_bounded():
    # A process that never dies: the whole close (wait + post-kill check) must
    # stay bounded, well inside the default 5 s reap grace.
    proc = FakeProcess(ignores_exit=True, survives_kill=True)
    s = _session(EmptyOrigin(proc), proc)
    s._empty_close_wait = 0.2
    t0 = time.monotonic()
    s._close_if_empty()
    assert time.monotonic() - t0 < 3.0


def test_default_wait_fits_inside_default_reap_grace():
    from origin_pro_mcp.daemon import EMPTY_CLOSE_WAIT_DEFAULT

    # wait + 1 s post-kill check < the 5 s watchdog grace
    assert EMPTY_CLOSE_WAIT_DEFAULT + 1.0 < 5.0


# --------------------------------------------------------------------------- #
# End to end through the Daemon (its terminate/is_alive reach the session)     #
# --------------------------------------------------------------------------- #

from test_lifecycle import (  # noqa: E402
    FakeClock,
    FakeFactory,
    _request,
    _wait_until,
    lifecycle,  # noqa: F401 — pytest fixture
)


def test_daemon_detach_reap_closes_empty_origin(lifecycle, tmp_path):  # noqa: F811
    start, client = lifecycle
    dead: set = set()

    class DaemonEmptyOrigin(EmptyOrigin):
        def Exit(self):
            self.exited = True
            dead.add(self.fake_pid)

    factory = FakeFactory(make=DaemonEmptyOrigin)
    daemon, killed = start(
        factory, clock=FakeClock(), reap_grace=5.0, get_pid=FakeFactory.get_pid,
        recovery_dir=str(tmp_path / "recovery"),
        is_alive=lambda pid: pid not in dead,
        lockfile_path=str(tmp_path / "daemon.json"),
    )
    conn = client(daemon, "emptysess")
    assert _request(conn, "r0", "run_labtalk", {"script": "g=1;"})["ok"] is True
    inst = factory.created[0]
    conn.close()
    assert _wait_until(lambda: daemon.pool.session_ids() == [])
    assert inst.exited is True       # empty -> closed even in detach mode
    assert killed == []              # Exit took; nothing force-killed
    assert inst.saved_paths == []    # detach never saves
