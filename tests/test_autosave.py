"""COM-free unit tests for the autosave core (policy, classification, has-work,
in-place save primitive)."""
import os

import pytest

from fakes import FakeOrigin
from origin_pro_mcp.autosave import (
    AutosavePolicy, HasWorkTracker, classify_autosave_labtalk,
    save_in_place, should_snapshot,
)


# --- policy env parsing ---------------------------------------------------- #

def test_policy_defaults_enabled_and_required():
    p = AutosavePolicy.from_env({})
    assert p.enabled is True
    assert p.required is True


def test_policy_opt_out():
    for off in ("off", "false", "no", "0"):
        p = AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE": off})
        assert p.enabled is False


def test_policy_required_can_be_cleared():
    p = AutosavePolicy.from_env({"ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED": "0"})
    assert p.required is False


def test_policy_no_retention_or_backup_dir_attrs():
    # in-place autosave has no retention / backup-dir concept anymore
    p = AutosavePolicy.from_env({})
    assert not hasattr(p, "retention")
    assert not hasattr(p, "backup_dir")


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


# --- in-place save primitive ----------------------------------------------- #

def test_save_in_place_saves_to_own_file(tmp_path):
    fake = FakeOrigin()
    proj = tmp_path / "study.opju"
    proj.write_text("real")  # in-place only overwrites an EXISTING file
    assert save_in_place(fake, str(proj)) is True
    assert fake.saved_paths == [str(proj)]          # saved IN PLACE, same name


def test_save_in_place_never_creates_a_differently_named_file(tmp_path):
    fake = FakeOrigin()
    proj = tmp_path / "study.opju"
    proj.write_text("real")
    save_in_place(fake, str(proj))
    assert fake.saved_paths == [str(proj)]          # no ".autosave-<timestamp>" copy
    assert not any(".autosave-" in p for p in fake.saved_paths)


def test_save_in_place_skips_when_no_on_disk_file(tmp_path):
    # never-saved project: no remembered path, LTStr("%X"/"%G") empty -> None
    # (issue #12 tri-state fix: "nothing to protect", NOT a save failure).
    fake = FakeOrigin()
    assert save_in_place(fake, None) is None
    assert fake.saved_paths == []


class _EmptyOrigin(FakeOrigin):
    def __init__(self):
        super().__init__()
        self.books = []
        self.graphs = []
        self.matrices = []


def test_save_in_place_skips_empty_project(tmp_path):
    """N5: an EMPTY project (0 windows, e.g. after a flaky empty-load) must NEVER
    overwrite a real file. Returns None ("nothing to protect"), not False —
    an empty project is not a save FAILURE."""
    fake = _EmptyOrigin()
    proj = tmp_path / "study.opju"
    proj.write_text("real 579 KB stand-in")
    assert save_in_place(fake, str(proj)) is None
    assert fake.saved_paths == []                    # real file untouched


def test_save_in_place_reports_failure(tmp_path):
    """A REAL save attempt (on-disk file exists) that fails must return
    False, distinctly from the None ("nothing to protect") cases above."""
    fake = FakeOrigin()
    fake.save_result = False
    proj = tmp_path / "x.opju"
    proj.write_text("real")
    assert save_in_place(fake, str(proj)) is False


# --- daemon dispatch integration (COM-free) -------------------------------- #

from origin_pro_mcp import transport
from origin_pro_mcp.daemon import Daemon


def _registry():
    from origin_pro_mcp import server  # noqa: F401 registers tools
    from origin_pro_mcp.app import mcp
    return {name: t.fn for name, t in mcp._tool_manager._tools.items()}


class _PathedOrigin(FakeOrigin):
    """A FakeOrigin that reports a real on-disk project path via %X/%G, so the
    in-place autosave has a file to save to."""
    def __init__(self, folder, name):
        super().__init__()
        self._folder = folder
        self._name = name

    def LTStr(self, key):
        if key == "%X":
            return self._folder
        if key == "%G":
            return self._name
        return ""


def _start_daemon(tmp_path, policy):
    origins = []
    # the project's on-disk file must exist (in-place save only overwrites it)
    proj = tmp_path / "proj.opju"
    proj.write_text("real project")

    def factory():
        o = _PathedOrigin(str(tmp_path) + os.sep, "proj")
        origins.append(o)
        return o

    d = Daemon()
    assert d.start(origin_factory=factory, registry=_registry(), max_size=3,
                   host="127.0.0.1", port=0, get_pid=lambda i: 4242,
                   autosave_policy=policy, monitor_tick=0.01, reconnect_grace=0.0,
                   lockfile_path=str(tmp_path / "daemon.json"))
    return d, origins, str(proj)


def _call(conn, rid, name, kwargs):
    conn.send_frame({"type": "request", "request_id": rid, "name": name, "kwargs": kwargs})
    return conn.recv_frame()


def test_daemon_autosave_snapshots_before_destructive(tmp_path):
    policy = AutosavePolicy(enabled=True, required=True)
    d, origins, proj = _start_daemon(tmp_path, policy)
    conn = transport.connect(d.host, d.port, d.token, session_id="A")
    conn.settimeout(5.0)
    try:
        # a mutating op first => has-work True
        r = _call(conn, "r1", "run_labtalk", {"script": "col(1)=1;"})
        assert r["ok"] is True
        origin = origins[0]
        assert origin.saved_paths == []          # no save for a non-destructive op
        # destructive op => project saved IN PLACE (same name) BEFORE it
        r = _call(conn, "r2", "delete_graph", {"graph_name": "Graph1"})
        assert r["ok"] is True
        assert origin.saved_paths == [proj]          # saved to its own file
        assert not any(".autosave-" in p for p in origin.saved_paths)
    finally:
        conn.close()
        d.stop()


def test_daemon_autosave_skips_when_no_work(tmp_path):
    policy = AutosavePolicy(enabled=True, required=True)
    d, origins, proj = _start_daemon(tmp_path, policy)
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
    policy = AutosavePolicy(enabled=True, required=True)
    d, origins, proj = _start_daemon(tmp_path, policy)
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


def _start_daemon_never_saved(tmp_path, policy):
    """A daemon whose Origin project has never been saved to disk (plain
    FakeOrigin: LTStr("%X"/"%G") is "" and no remembered path) — the
    never-saved-project half of the issue #12 fix."""
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


def test_daemon_autosave_never_saved_project_does_not_block_delete(tmp_path):
    """Issue #12: a never-saved project (no on-disk file, e.g. a fresh session
    that only just created data) must NOT block a destructive op under a
    REQUIRED autosave policy — there is nothing on disk to protect, so
    save_in_place returns None and the preflight proceeds."""
    policy = AutosavePolicy(enabled=True, required=True)
    d, origins = _start_daemon_never_saved(tmp_path, policy)
    conn = transport.connect(d.host, d.port, d.token, session_id="A")
    conn.settimeout(5.0)
    try:
        _call(conn, "r1", "run_labtalk", {"script": "col(1)=1;"})  # has-work
        r = _call(conn, "r2", "delete_graph", {"graph_name": "Graph1"})
        assert r["ok"] is True
        assert "Autosave" not in (r["error"] or "")
        assert origins[0].saved_paths == []  # nothing to save to; not attempted
    finally:
        conn.close()
        d.stop()


def test_daemon_autosave_off_by_default(tmp_path):
    d, origins, proj = _start_daemon(tmp_path, None)  # no policy => off
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
