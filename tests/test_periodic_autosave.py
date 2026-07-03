"""Proactive periodic autosave scheduling (COM-free).

The daemon monitor enqueues a snapshot onto each HEALTHY, agent-isolated session
whose interval has elapsed. It must skip: sessions being reaped, sessions with a
dispatch in flight (busy), and any session with an unknown/None PID (the attached
USER Origin is pid=None and must NEVER be auto-saved). A busy/attached session is
retried next tick (its last-snapshot clock is not advanced)."""
from __future__ import annotations

from origin_pro_mcp.autosave import AutosavePolicy
from origin_pro_mcp.daemon import Daemon


class _StubSession:
    def __init__(self, pid):
        self.pid = pid
        self.snapshots = 0

    def submit_snapshot(self):
        self.snapshots += 1


class _StubPool:
    def __init__(self, sessions):
        self._sessions = sessions

    def session_ids(self):
        return list(self._sessions.keys())

    def get(self, sid):
        return self._sessions.get(sid)


def _daemon_with(sessions, interval=300.0, enabled=True, busy=(), reaping=()):
    d = Daemon()
    d._pool = _StubPool(sessions)
    d._autosave_interval = interval
    d._autosave_policy = AutosavePolicy(enabled=enabled) if enabled else None
    d._dispatch_tickets = {sid: {} for sid in busy}
    d._reaping = {sid: object() for sid in reaping}
    d._last_snapshot = {}
    return d


def test_schedule_snapshots_enqueues_healthy_isolated_session():
    s = _StubSession(pid=4242)
    d = _daemon_with({"A": s})
    d._schedule_snapshots(now=1000.0)
    assert s.snapshots == 1
    assert d._last_snapshot["A"] == 1000.0


def test_schedule_snapshots_skips_attached_or_unknown_pid():
    # pid=None models the attached user Origin (never auto-save it).
    s = _StubSession(pid=None)
    d = _daemon_with({"A": s})
    d._schedule_snapshots(now=1000.0)
    assert s.snapshots == 0
    assert "A" not in d._last_snapshot  # not advanced -> retried next tick


def test_schedule_snapshots_skips_busy_session():
    s = _StubSession(pid=4242)
    d = _daemon_with({"A": s}, busy=("A",))
    d._schedule_snapshots(now=1000.0)
    assert s.snapshots == 0
    assert "A" not in d._last_snapshot


def test_schedule_snapshots_skips_reaping_session():
    s = _StubSession(pid=4242)
    d = _daemon_with({"A": s}, reaping=("A",))
    d._schedule_snapshots(now=1000.0)
    assert s.snapshots == 0


def test_schedule_snapshots_respects_interval():
    s = _StubSession(pid=4242)
    d = _daemon_with({"A": s}, interval=300.0)
    d._schedule_snapshots(now=1000.0)
    assert s.snapshots == 1
    # too soon -> no second snapshot
    d._schedule_snapshots(now=1000.0 + 299.0)
    assert s.snapshots == 1
    # interval elapsed -> snapshot again
    d._schedule_snapshots(now=1000.0 + 300.0)
    assert s.snapshots == 2


def test_schedule_snapshots_multiple_sessions_mixed():
    good = _StubSession(pid=1)
    attached = _StubSession(pid=None)
    busy = _StubSession(pid=2)
    d = _daemon_with({"good": good, "att": attached, "busy": busy}, busy=("busy",))
    d._schedule_snapshots(now=500.0)
    assert good.snapshots == 1
    assert attached.snapshots == 0
    assert busy.snapshots == 0
