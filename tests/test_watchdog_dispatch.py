"""Typed, routable Watchdog deadlines (D1): a dispatch deadline and a reap
deadline coexist for one session, route to distinct callbacks, and disarming
one reason never cancels the other. Legacy arm/disarm behavior is preserved."""
import threading
import time

from origin_pro_mcp.daemon import Watchdog


def _spin(pred, timeout=2.0, interval=0.01):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if pred():
            return True
        time.sleep(interval)
    return pred()


def _wd(**kw):
    kw.setdefault("tick", 0.005)
    wd = Watchdog(**kw)
    wd.start()
    return wd


def test_legacy_arm_disarm_reap(monkeypatch):
    killed, reaped = [], []
    wd = _wd(terminate_process=lambda p: killed.append(p),
             on_reap=lambda s, p: reaped.append((s, p)))
    try:
        wd.arm("s", pid=7, deadline=time.monotonic() + 0.02)  # legacy signature
        assert _spin(lambda: reaped == [("s", 7)]), reaped
        assert killed == [7]
    finally:
        wd.stop()


def test_legacy_disarm_cancels_reap():
    reaped = []
    wd = _wd(terminate_process=lambda p: None, on_reap=lambda s, p: reaped.append(s))
    try:
        wd.arm("s", pid=1, deadline=time.monotonic() + 0.3)
        wd.disarm("s")  # legacy disarm(session_id) -> cancels the reap
        time.sleep(0.5)
        assert reaped == []
    finally:
        wd.stop()


def test_dispatch_and_reap_coexist_route_to_distinct_callbacks():
    reaped, dispatched = [], []
    wd = _wd(terminate_process=lambda p: None, on_reap=lambda s, p: reaped.append((s, p)))
    try:
        now = time.monotonic()
        wd.arm("s", pid=10, deadline=now + 0.02)  # reap (default reason, on_reap)
        wd.arm("s", pid=11, deadline=now + 0.02, reason="dispatch",
               callback=lambda s, p: dispatched.append((s, p)))
        assert _spin(lambda: reaped == [("s", 10)] and dispatched == [("s", 11)]), (reaped, dispatched)
    finally:
        wd.stop()


def test_disarm_dispatch_leaves_reap_armed():
    reaped, dispatched = [], []
    wd = _wd(terminate_process=lambda p: None, on_reap=lambda s, p: reaped.append(s))
    try:
        now = time.monotonic()
        wd.arm("s", pid=1, deadline=now + 0.15)  # reap
        wd.arm("s", pid=2, deadline=now + 0.15, reason="dispatch",
               callback=lambda s, p: dispatched.append(s))
        wd.disarm("s", reason="dispatch")  # cancel ONLY the dispatch deadline
        assert _spin(lambda: reaped == ["s"]), reaped
        assert dispatched == []  # dispatch was cancelled; reap still fired
    finally:
        wd.stop()


def test_disarm_reap_leaves_dispatch_armed():
    reaped, dispatched = [], []
    wd = _wd(terminate_process=lambda p: None, on_reap=lambda s, p: reaped.append(s))
    try:
        now = time.monotonic()
        wd.arm("s", pid=1, deadline=now + 0.15)  # reap
        wd.arm("s", pid=2, deadline=now + 0.15, reason="dispatch",
               callback=lambda s, p: dispatched.append(s))
        wd.disarm("s")  # legacy disarm -> reason 'reap' only
        assert _spin(lambda: dispatched == ["s"]), dispatched
        assert reaped == []  # reap cancelled; dispatch still fired
    finally:
        wd.stop()


def test_dispatch_deadline_kills_pid_then_routes_callback():
    killed, dispatched = [], []
    wd = _wd(terminate_process=lambda p: killed.append(p), on_reap=lambda s, p: None)
    try:
        wd.arm("s", pid=4242, deadline=time.monotonic() + 0.02, reason="dispatch",
               callback=lambda s, p: dispatched.append((s, p)))
        assert _spin(lambda: dispatched == [("s", 4242)]), dispatched
        assert killed == [4242]  # wedged session's Origin PID force-killed
    finally:
        wd.stop()


def test_dispatch_pid_none_skips_kill_still_routes():
    killed, dispatched = [], []
    wd = _wd(terminate_process=lambda p: killed.append(p), on_reap=lambda s, p: None)
    try:
        wd.arm("s", pid=None, deadline=time.monotonic() + 0.02, reason="dispatch",
               callback=lambda s, p: dispatched.append((s, p)))
        assert _spin(lambda: dispatched == [("s", None)]), dispatched
        assert killed == []  # unknown PID -> no kill, but callback still frees the slot
    finally:
        wd.stop()


def test_disarm_all_clears_both_reasons():
    fired = []
    wd = _wd(terminate_process=lambda p: None, on_reap=lambda s, p: fired.append(("reap", s)))
    try:
        now = time.monotonic()
        wd.arm("s", pid=1, deadline=now + 0.2)
        wd.arm("s", pid=2, deadline=now + 0.2, reason="dispatch",
               callback=lambda s, p: fired.append(("dispatch", s)))
        wd.disarm_all("s")
        time.sleep(0.4)
        assert fired == []
    finally:
        wd.stop()
