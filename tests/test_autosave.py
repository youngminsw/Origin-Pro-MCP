"""COM-free unit tests for the autosave core (policy, classification, has-work,
backup pathing/retention, save-copy primitive)."""
import os

import pytest

from fakes import FakeOrigin
from origin_pro_mcp.autosave import (
    AutosavePolicy, HasWorkTracker, backup_path, classify_autosave_labtalk,
    prune_backups, save_copy, should_snapshot,
)


# --- policy env parsing ---------------------------------------------------- #

def test_policy_defaults_enabled_and_required():
    p = AutosavePolicy.from_env({})
    assert p.enabled is True
    assert p.required is True
    assert p.retention == 3


def test_policy_opt_out():
    for off in ("off", "false", "no", "0"):
        p = AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE": off})
        assert p.enabled is False


def test_policy_required_can_be_cleared():
    p = AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED": "0"})
    assert p.required is False


def test_policy_retention_parse_and_fallback():
    assert AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE_RETENTION": "5"}).retention == 5
    assert AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE_RETENTION": "x"}).retention == 3
    # retention floored at 1
    assert AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE_RETENTION": "0"}).retention == 1


# --- labtalk classification ------------------------------------------------ #

@pytest.mark.parametrize("script", [
    "delete col(1);", "del data1;", "win -cd Graph1;", "win -c Book1;",
    "doc -n;", "for(i=1;i<=3;i++) delete wks.col$(i);",
])
def test_classify_labtalk_destructive_true(script):
    assert classify_autosave_labtalk(script) is True


@pytest.mark.parametrize("script", [
    "col(3) = col(1) + col(2);", "plotxy iy:=(1,2);", "save;",
    'print("delete this");',            # token inside a string literal
    "// delete col(1)",                  # token inside a comment
    "", "win -a Book1;", "doc -s;",      # save is not a destroyer
])
def test_classify_labtalk_destructive_false(script):
    assert classify_autosave_labtalk(script) is False


# --- should_snapshot ------------------------------------------------------- #

def test_should_snapshot_typed_destructive():
    assert should_snapshot("delete_graph", {"graph_name": "G"}) is True
    assert should_snapshot("remove_plot", {"graph_name": "G"}) is True
    assert should_snapshot("manage_columns", {}) is True
    assert should_snapshot("new_project", {}) is True
    assert should_snapshot("load_project", {"file_path": "p.opju"}) is True


def test_should_snapshot_excludes_worksheet_to_matrix():
    assert should_snapshot("worksheet_to_matrix", {"book_name": "Book1"}) is False


def test_should_snapshot_readonly_false():
    assert should_snapshot("get_worksheet_data", {"book_name": "Book1"}) is False
    assert should_snapshot("list_worksheets", {}) is False


def test_should_snapshot_overwrite_only_when_target_has_data():
    fake = FakeOrigin()
    # default worksheet_data is non-empty => overwrite => snapshot
    assert should_snapshot("set_worksheet_data",
                           {"book_name": "Book1", "sheet_name": "Sheet1"}, fake) is True
    # unknown/empty target => GetWorksheet returns HRESULT int => no snapshot
    assert should_snapshot("set_worksheet_data",
                           {"book_name": "Nope", "sheet_name": "Sheet1"}, fake) is False


def test_should_snapshot_run_labtalk_requires_confirm_and_destructive():
    # destructive + confirmed => snapshot
    assert should_snapshot("run_labtalk",
                           {"script": "delete col(1);", "confirm": True}) is True
    # destructive but NOT confirmed (won't execute) => no snapshot
    assert should_snapshot("run_labtalk",
                           {"script": "delete col(1);", "confirm": False}) is False
    # confirmed but non-destructive => no snapshot
    assert should_snapshot("run_labtalk",
                           {"script": "col(1)=1;", "confirm": True}) is False


# --- has-work tracker ------------------------------------------------------ #

def test_has_work_tracker():
    t = HasWorkTracker()
    assert t.has_work is False
    t.record_success("list_worksheets")   # readonly => still no work
    assert t.has_work is False
    t.record_success("set_worksheet_data")  # mutating => work
    assert t.has_work is True
    t.record_success("new_project")        # resets
    assert t.has_work is False


def test_has_work_build_then_destroy():
    """A build (create) followed by a destroy: has-work is set before the
    destroy so a snapshot would fire."""
    t = HasWorkTracker()
    t.record_success("create_worksheet")
    assert t.has_work is True
    assert should_snapshot("delete_graph", {"graph_name": "G"}) and t.has_work


def test_has_work_load_then_destroy():
    t = HasWorkTracker()
    t.record_success("load_project")  # load is non-readonly => work
    assert t.has_work is True


# --- backup pathing + retention -------------------------------------------- #

def test_backup_path_named_project(tmp_path):
    proj = str(tmp_path / "study.opju")
    p = backup_path(AutosavePolicy(), proj, now=0)
    assert os.path.dirname(p) == str(tmp_path)
    assert os.path.basename(p).startswith("study.autosave-")
    assert p.endswith(".opju")


def test_backup_path_unsaved_project_uses_backup_dir(tmp_path):
    policy = AutosavePolicy(backup_dir=str(tmp_path))
    p = backup_path(policy, None, now=0)
    assert os.path.dirname(p) == str(tmp_path)
    assert os.path.basename(p).startswith("untitled.autosave-")


def test_prune_backups_keeps_newest(tmp_path):
    proj = str(tmp_path / "study.opju")
    names = [f"study.autosave-2024010{i}-000000.opju" for i in range(1, 6)]
    for n in names:
        (tmp_path / n).write_text("x")
    (tmp_path / "study.opju").write_text("real")  # not an autosave copy
    removed = prune_backups(AutosavePolicy(retention=2), proj)
    # 5 copies, keep 2 => remove 3 oldest
    assert len(removed) == 3
    survivors = sorted(f.name for f in tmp_path.iterdir()
                       if f.name.startswith("study.autosave-"))
    assert survivors == names[3:]  # newest 2
    assert (tmp_path / "study.opju").exists()  # real project untouched


# --- save-copy primitive --------------------------------------------------- #

def test_save_copy_writes_and_reports(tmp_path):
    fake = FakeOrigin()
    dest = str(tmp_path / "study.autosave-0.opju")
    assert save_copy(fake, dest, None) is True
    assert fake.saved_paths == [dest]

def test_save_copy_restores_original_binding(tmp_path):
    """Save(path) rebinds project identity (spike-verified), so save_copy must
    Save(dest) then Save(remembered) to restore the user's project binding."""
    fake = FakeOrigin()
    dest = str(tmp_path / "study.autosave-0.opju")
    original = str(tmp_path / "study.opju")
    assert save_copy(fake, dest, original) is True
    assert fake.saved_paths == [dest, original]  # backup, then restore binding


def test_save_copy_reports_failure(tmp_path):
    fake = FakeOrigin()
    fake.save_result = False
    assert save_copy(fake, str(tmp_path / "x.opju"), None) is False


# --- daemon dispatch integration (COM-free) -------------------------------- #

from origin_pro_mcp import transport
from origin_pro_mcp.daemon import Daemon


def _registry():
    from origin_pro_mcp import server  # noqa: F401 registers tools
    from origin_pro_mcp.app import mcp
    return {name: t.fn for name, t in mcp._tool_manager._tools.items()}


def _start_daemon(tmp_path, policy):
    origins = []

    def factory():
        o = FakeOrigin()
        origins.append(o)
        return o

    d = Daemon()
    assert d.start(origin_factory=factory, registry=_registry(), max_size=3,
                   host="127.0.0.1", port=0, get_pid=lambda i: 4242,
                   autosave_policy=policy, monitor_tick=0.01, reconnect_grace=0.0,
                   lockfile_path=str(tmp_path / "daemon.json"))
    return d, origins


def _call(conn, rid, name, kwargs):
    conn.send_frame({"type": "request", "request_id": rid, "name": name, "kwargs": kwargs})
    return conn.recv_frame()


def test_daemon_autosave_snapshots_before_destructive(tmp_path):
    policy = AutosavePolicy(enabled=True, required=True, backup_dir=str(tmp_path))
    d, origins = _start_daemon(tmp_path, policy)
    conn = transport.connect(d.host, d.port, d.token, session_id="A")
    conn.settimeout(5.0)
    try:
        # a mutating op first => has-work True
        r = _call(conn, "r1", "run_labtalk", {"script": "col(1)=1;"})
        assert r["ok"] is True
        origin = origins[0]
        assert origin.saved_paths == []          # no snapshot for a non-destructive op
        # destructive op => snapshot fires BEFORE it
        r = _call(conn, "r2", "delete_graph", {"graph_name": "Graph1"})
        assert r["ok"] is True
        assert len(origin.saved_paths) == 1
        assert "untitled.autosave-" in origin.saved_paths[0]
    finally:
        conn.close()
        d.stop()


def test_daemon_autosave_skips_when_no_work(tmp_path):
    policy = AutosavePolicy(enabled=True, required=True, backup_dir=str(tmp_path))
    d, origins = _start_daemon(tmp_path, policy)
    conn = transport.connect(d.host, d.port, d.token, session_id="A")
    conn.settimeout(5.0)
    try:
        # first op is destructive but there is NO prior work => skip snapshot
        r = _call(conn, "r1", "delete_graph", {"graph_name": "Graph1"})
        assert r["ok"] is True
        assert origins[0].saved_paths == []
    finally:
        conn.close()
        d.stop()


def test_daemon_autosave_required_failure_blocks_destructive(tmp_path):
    policy = AutosavePolicy(enabled=True, required=True, backup_dir=str(tmp_path))
    d, origins = _start_daemon(tmp_path, policy)
    conn = transport.connect(d.host, d.port, d.token, session_id="A")
    conn.settimeout(5.0)
    try:
        _call(conn, "r1", "run_labtalk", {"script": "col(1)=1;"})  # has-work
        origins[0].save_result = False  # autosave will fail
        r = _call(conn, "r2", "delete_graph", {"graph_name": "Graph1"})
        assert r["ok"] is False
        assert "Autosave before 'delete_graph' failed" in r["error"]
        # the destructive win -cd must NOT have run
        assert not any("win -cd" in s for s in origins[0].executed)
    finally:
        conn.close()
        d.stop()


def test_daemon_autosave_off_by_default(tmp_path):
    d, origins = _start_daemon(tmp_path, None)  # no policy => off
    conn = transport.connect(d.host, d.port, d.token, session_id="A")
    conn.settimeout(5.0)
    try:
        _call(conn, "r1", "run_labtalk", {"script": "col(1)=1;"})
        r = _call(conn, "r2", "delete_graph", {"graph_name": "Graph1"})
        assert r["ok"] is True
        assert origins[0].saved_paths == []  # nothing saved when autosave is off
    finally:
        conn.close()
        d.stop()
