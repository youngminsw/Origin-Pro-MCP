"""A session keeps tracking its Origin after a relaunch.

When the COM proxy dies (close_origin, an Origin crash), get_origin relaunches
the session's own Origin through the factory the session registered. The
session must then point at the NEW instance and PID, or the detach reap closes
(or keeps) the wrong Origin and the relaunched one is left running when the
session ends. No real process is touched: PID, liveness and kill are faked.
"""
import threading

import pytest

from fakes import FakeOrigin
from origin_pro_mcp import origin_connection as oc
from origin_pro_mcp.daemon import Session


class EmptyClosable(FakeOrigin):
    """No windows; Exit kills the COM proxy, like a closed Origin."""

    def __init__(self, pid):
        super().__init__()
        self.books, self.graphs, self.matrices = [], [], []
        self.pid = pid
        self.alive = True
        self.exited = False

    @property
    def Visible(self):
        if not self.alive:
            raise OSError("The RPC server is unavailable.")
        return True

    def Exit(self):
        self.exited = True
        self.alive = False


class FakeDialogWatchdog:
    def __init__(self, pid, _record):
        self.pid = pid
        self.running = False

    def start(self):
        self.running = True

    def stop(self):
        self.running = False


@pytest.fixture(autouse=True)
def _detach_default(monkeypatch):
    monkeypatch.delenv("ORIGIN_PRO_MCP_REAP_CLOSE", raising=False)


@pytest.fixture
def relaunching_session():
    made = []
    watchdogs = []

    def factory():
        inst = EmptyClosable(pid=5000 + len(made))
        made.append(inst)
        return inst

    def watchdog_factory(pid, record):
        wd = FakeDialogWatchdog(pid, record)
        watchdogs.append(wd)
        return wd

    def close_then_touch():
        # What close_origin followed by the next tool call does on the worker.
        oc.get_origin().Exit()
        return oc.get_origin().pid

    s = Session("s-relaunch", factory, {"close_then_touch": close_then_touch},
                get_pid=lambda inst: inst.pid,
                dialog_watchdog_factory=watchdog_factory,
                terminate_process=lambda pid: None,
                process_running=lambda pid: not next(
                    m for m in made if m.pid == pid).exited,
                empty_close_wait=0.05)
    s.start(ready_timeout=5)
    yield s, made, watchdogs
    s.stop()


def _call(session, name):
    done = threading.Event()
    box = {}

    def reply(msg):
        box.update(msg)
        done.set()

    session.submit("r1", name, {}, reply)
    assert done.wait(5), "tool call did not finish"
    return box


def test_relaunch_updates_instance_and_pid(relaunching_session):
    s, made, _ = relaunching_session
    assert (s.instance, s.pid) == (made[0], 5000)
    _call(s, "close_then_touch")
    assert len(made) == 2
    assert s.instance is made[1]
    assert s.pid == 5001


def test_reap_closes_the_relaunched_empty_origin(relaunching_session):
    s, made, _ = relaunching_session
    _call(s, "close_then_touch")
    assert s._close_if_empty() == "closed"
    assert made[1].exited is True            # the Origin running now is closed


def test_dialog_watchdog_follows_the_new_pid(relaunching_session):
    s, _, watchdogs = relaunching_session
    assert [w.pid for w in watchdogs] == [5000]
    _call(s, "close_then_touch")
    assert [w.pid for w in watchdogs] == [5000, 5001]
    assert watchdogs[0].running is False
    assert watchdogs[1].running is True
