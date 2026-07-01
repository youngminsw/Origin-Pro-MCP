"""N4 (P0) stability: orphan-Origin reaping via the persistent spawn-log, and
load_project's empty-session verification."""

import pytest

from fakes import FakeOrigin
from origin_pro_mcp import origin_connection as oc
from origin_pro_mcp.daemon import (
    Daemon,
    record_spawned_pid,
    read_spawned_pids,
    clear_spawn_log,
)


# --- spawn-log roundtrip ------------------------------------------------------

def test_spawn_log_roundtrip(tmp_path):
    log = str(tmp_path / "spawned-pids.log")
    record_spawned_pid(111, log)
    record_spawned_pid(222, log)
    record_spawned_pid(None, log)  # ignored
    assert read_spawned_pids(log) == [111, 222]
    clear_spawn_log(log)
    assert read_spawned_pids(log) == []


# --- startup sweep reclaims spawn-logged orphans ------------------------------

def test_startup_sweep_kills_spawnlog_orphans(tmp_path):
    log = str(tmp_path / "spawned-pids.log")
    record_spawned_pid(4242, log)
    record_spawned_pid(4243, log)

    d = Daemon()
    d._spawn_log_path = log
    killed = []
    d._terminate = lambda pid: killed.append(pid)

    # No lockfile; _origin_process_pids() is empty on non-Windows so the sweep
    # falls back to is_alive (report both alive).
    d._startup_sweep(str(tmp_path / "no_such_lockfile.json"), is_alive=lambda p: True)

    assert set(killed) == {4242, 4243}
    assert read_spawned_pids(log) == []  # log cleared after sweep


def test_startup_sweep_skips_dead_pids(tmp_path):
    log = str(tmp_path / "spawned-pids.log")
    record_spawned_pid(500, log)
    record_spawned_pid(501, log)

    d = Daemon()
    d._spawn_log_path = log
    killed = []
    d._terminate = lambda pid: killed.append(pid)
    d._startup_sweep(str(tmp_path / "none.json"), is_alive=lambda p: p == 500)

    assert killed == [500]  # only the alive one


# --- get_origin closes a stale instance before relaunch (no orphan) ----------

class _DeadProxy:
    def __init__(self):
        self.dead = False
        self.closed = False

    @property
    def Visible(self):
        if self.dead:
            raise RuntimeError("RPC server unavailable")
        return 1

    def Exit(self):
        self.closed = True


def test_get_origin_closes_stale_before_relaunch(tmp_path):
    oc.clear_session_origin()
    try:
        dead = _DeadProxy()
        fresh = FakeOrigin()
        oc.set_session_origin(dead, lambda: fresh)
        dead.dead = True
        got = oc.get_origin()
        assert got is fresh          # relaunched
        assert dead.closed is True   # old instance was closed, not orphaned
    finally:
        oc.clear_session_origin()


def test_connection_alive_retries_before_giving_up():
    calls = {"n": 0}

    class Flaky:
        @property
        def Visible(self):
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("call rejected by callee (busy)")
            return 1

    assert oc._connection_alive(Flaky()) is True
    assert calls["n"] == 2  # retried once


# --- load_project verifies a non-empty load ----------------------------------

def test_load_project_rejects_empty_load(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import load_project

    proj = tmp_path / "empty.opju"
    proj.write_text("stub")
    fake_origin.books = []
    fake_origin.graphs = []
    fake_origin.matrices = []
    with pytest.raises(ValueError, match="empty"):
        load_project(str(proj))


def test_load_project_accepts_nonempty_load(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import load_project

    proj = tmp_path / "real.opju"
    proj.write_text("stub")
    # FakeOrigin defaults to one workbook -> non-empty.
    msg = load_project(str(proj))
    assert "Loaded project" in msg
