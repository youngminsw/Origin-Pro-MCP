"""close_origin (origin2018-improvements).

COM-free: a FakeOrigin subclass models Exit (optionally ignored, like Origin
2018 with a -Embedding instance), the PID helpers are monkeypatched, and a
fake clock keeps the 10 s / 3 s waits instant.
"""
import sys
import threading
import types

import pytest
from fakes import FakeOrigin

from origin_pro_mcp import origin_connection as oc
from origin_pro_mcp.tools import project
from origin_pro_mcp.tools.project import close_origin

# The real helpers, captured before the autouse fixture stubs them out.
_REAL_SESSION_PID = project._session_origin_pid
_REAL_FORGET_PID = project._forget_session_origin_pid
_REAL_TERMINATE = project._terminate_pid
_REAL_PID_ALIVE = project._pid_alive


class ClosableOrigin(FakeOrigin):
    """FakeOrigin with a lifecycle. ``exit_works=False`` models an Exit that
    is ignored (Origin 2018 -Embedding); ``Visible`` raises once closed, so
    ``_connection_alive`` reports a dead proxy."""

    def __init__(self, log=None, exit_works=True):
        super().__init__()
        self.alive = True
        self.exit_works = exit_works
        self.exit_calls = 0
        self.load_calls = []
        self.log = log if log is not None else []

    @property
    def Visible(self):
        if not self.alive:
            raise OSError("The RPC server is unavailable.")
        return True

    def Execute(self, script):
        self.log.append(("execute", script))
        return super().Execute(script)

    def Exit(self):
        self.exit_calls += 1
        self.log.append(("exit",))
        if self.exit_works:
            self.alive = False

    def Save(self, path):
        self.log.append(("save", path))
        return super().Save(path)

    def Load(self, path):
        self.load_calls.append(path)
        return True


class FakeClock:
    """Instant stand-in for the ``time`` module used by project._wait_closed."""

    def __init__(self):
        self.now = 0.0

    def monotonic(self):
        return self.now

    def sleep(self, s):
        self.now += s


@pytest.fixture(autouse=True)
def _isolate(monkeypatch):
    oc.clear_session_origin()
    monkeypatch.setattr(project, "time", FakeClock())
    # Never touch real processes / real daemon PID stores from these tests.
    monkeypatch.setattr(project, "_session_origin_pid", lambda: None)
    monkeypatch.setattr(project, "_forget_session_origin_pid", lambda: None)

    def no_terminate(pid):  # pragma: no cover - only hit if a test forgets
        raise AssertionError(f"unexpected terminate of PID {pid}")

    monkeypatch.setattr(project, "_terminate_pid", no_terminate)
    monkeypatch.setattr(project, "_pid_alive",
                        lambda pid: (_ for _ in ()).throw(AssertionError("pid probe")))
    yield
    oc.clear_session_origin()


def _bind_own(o, factory=None):
    """Bind ``o`` as this thread's OWN isolated instance (daemon session)."""
    oc.set_session_origin(o, factory or (lambda: ClosableOrigin()), "sess-1")
    oc.mark_session_attached(False)


# --- nothing to close ------------------------------------------------------ #


def test_no_origin_bound_reports_nothing_to_close():
    assert "nothing to close" in close_origin()


def test_dead_proxy_reports_nothing_to_close_and_does_not_relaunch():
    dead = ClosableOrigin()
    dead.alive = False
    launched = []
    oc.set_session_origin(dead, lambda: launched.append(1) or ClosableOrigin())
    assert "nothing to close" in close_origin()
    assert launched == []          # peek_session_origin never relaunches
    assert dead.exit_calls == 0


# --- shared Origin guard ---------------------------------------------------- #


def test_in_process_mode_requires_confirm():
    o = ClosableOrigin()
    oc.set_session_origin(o)       # no factory -> shared Origin.ApplicationSI
    assert oc.session_origin_is_shared() is True
    with pytest.raises(ValueError, match="confirm=True"):
        close_origin()
    assert o.exit_calls == 0
    assert o.log == []             # not even "doc -s;"
    assert oc.peek_session_origin() is o   # still bound


def test_attach_mode_requires_confirm():
    o = ClosableOrigin()
    oc.set_session_origin(o, lambda: ClosableOrigin(), "sess-1")
    oc.mark_session_attached(True)
    assert oc.session_origin_is_shared() is True
    with pytest.raises(ValueError, match="Refusing to close"):
        close_origin(save=False)
    assert o.exit_calls == 0


def test_attach_mode_with_confirm_closes_but_never_terminates(monkeypatch):
    o = ClosableOrigin(exit_works=False)   # the user's Origin ignores Exit
    oc.set_session_origin(o, lambda: ClosableOrigin(), "sess-1")
    oc.mark_session_attached(True)
    # Even if a PID were known it must not be used for an attached window.
    monkeypatch.setattr(project, "_session_origin_pid", lambda: 4242)
    out = close_origin(confirm=True)
    assert o.exit_calls == 1
    assert "has not exited yet" in out     # no terminate fallback
    assert oc.get_remembered_project_path() is None


def test_in_process_with_confirm_closes_via_proxy_death():
    o = ClosableOrigin()
    oc.set_session_origin(o)
    out = close_origin(confirm=True)
    assert out == "Closed this session's Origin (unsaved changes discarded)."
    assert o.exit_calls == 1


def test_own_instance_is_not_shared():
    _bind_own(ClosableOrigin())
    assert oc.session_origin_is_shared() is False


def test_clear_session_origin_resets_attached_flag():
    oc.set_session_origin(ClosableOrigin(), lambda: ClosableOrigin(), "s")
    oc.mark_session_attached(True)
    oc.clear_session_origin()
    oc.set_session_origin(ClosableOrigin(), lambda: ClosableOrigin(), "s")
    assert oc.session_origin_is_shared() is False


# --- save / discard ---------------------------------------------------------- #


def test_save_false_marks_clean_before_exit():
    log = []
    o = ClosableOrigin(log=log)
    _bind_own(o)
    out = close_origin(save=False)
    assert log == [("execute", "doc -s;"), ("exit",)]
    assert "unsaved changes discarded" in out
    assert o.saved_paths == []


def test_save_false_survives_doc_s_raising():
    o = ClosableOrigin()

    def boom(script):
        raise OSError("busy")

    o.Execute = boom
    _bind_own(o)
    assert "discarded" in close_origin()
    assert o.exit_calls == 1


def test_save_true_saves_in_place_then_exits(tmp_path):
    proj = tmp_path / "exp.opju"
    proj.write_bytes(b"x" * 4096)
    log = []
    o = ClosableOrigin(log=log)
    _bind_own(o)
    oc.remember_project_path(str(proj))
    out = close_origin(save=True)
    assert log == [("save", str(proj)), ("exit",)]   # full path, never Save("")
    assert "changes saved" in out


def test_save_true_refuses_empty_project():
    o = ClosableOrigin()
    o.books, o.graphs, o.matrices = [], [], []
    _bind_own(o)
    with pytest.raises(ValueError, match="project is empty"):
        close_origin(save=True)
    assert o.exit_calls == 0
    assert o.saved_paths == []
    assert oc.peek_session_origin() is o


def test_save_true_reports_save_failure_and_keeps_origin_open():
    o = ClosableOrigin()
    o.save_result = False          # untitled project: no file to save over
    _bind_own(o)
    with pytest.raises(ValueError, match="save_project"):
        close_origin(save=True)
    assert o.saved_paths == []     # never calls Save("") on an untitled project
    assert o.exit_calls == 0
    assert oc.peek_session_origin() is o


# --- PID-based confirmation / terminate backstop ------------------------------ #


def test_pid_confirms_exit_without_terminate(monkeypatch):
    o = ClosableOrigin()
    _bind_own(o)
    monkeypatch.setattr(project, "_session_origin_pid", lambda: 777)
    probes = []

    def alive(pid):
        probes.append(pid)
        return o.alive

    monkeypatch.setattr(project, "_pid_alive", alive)
    forgot = []
    monkeypatch.setattr(project, "_forget_session_origin_pid", lambda: forgot.append(1))
    out = close_origin()
    assert out == "Closed this session's Origin (unsaved changes discarded)."
    assert probes and set(probes) == {777}
    assert forgot == [1]


def test_ignored_exit_is_terminated_by_pid(monkeypatch):
    o = ClosableOrigin(exit_works=False)   # Origin 2018 -Embedding behaviour
    _bind_own(o)
    monkeypatch.setattr(project, "_session_origin_pid", lambda: 4242)
    state = {"alive": True}
    monkeypatch.setattr(project, "_pid_alive", lambda pid: state["alive"])
    killed = []

    def terminate(pid):
        killed.append(pid)
        state["alive"] = False

    monkeypatch.setattr(project, "_terminate_pid", terminate)
    out = close_origin()
    assert o.exit_calls == 1
    assert killed == [4242]
    assert "PID 4242" in out and "terminated" in out
    assert oc.get_remembered_project_path() is None


def test_terminate_that_fails_reports_not_exited(monkeypatch):
    o = ClosableOrigin(exit_works=False)
    _bind_own(o)
    monkeypatch.setattr(project, "_session_origin_pid", lambda: 4242)
    monkeypatch.setattr(project, "_pid_alive", lambda pid: True)
    killed = []

    def terminate(pid):
        killed.append(pid)
        raise OSError("access denied")

    monkeypatch.setattr(project, "_terminate_pid", terminate)
    out = close_origin()
    assert killed == [4242]
    assert "has not exited yet" in out


def test_unknown_pid_never_terminates():
    o = ClosableOrigin(exit_works=False)
    _bind_own(o)                   # _session_origin_pid -> None (fixture)
    out = close_origin()
    assert "has not exited yet" in out   # fixture's _terminate_pid would raise


# --- after close: fresh, EMPTY instance --------------------------------------- #


def test_close_drops_project_and_next_get_origin_is_fresh_and_empty(tmp_path):
    proj = tmp_path / "exp.opju"
    proj.write_bytes(b"x" * 64)    # exists on disk -> would be auto-reopened
    fresh = ClosableOrigin()
    launches = []

    def factory():
        launches.append(1)
        return fresh

    old = ClosableOrigin()
    _bind_own(old, factory)
    oc.remember_project_path(str(proj))
    assert oc.get_remembered_project_path() == str(proj)

    close_origin()
    assert oc.get_remembered_project_path() is None
    assert oc.peek_session_origin() is None

    got = oc.get_origin()
    assert got is fresh
    assert launches == [1]
    assert fresh.load_calls == []  # the closed project is NOT reopened


def test_drop_session_origin_keeps_factory_and_notifies_ledger(monkeypatch):
    events = []
    monkeypatch.setattr(oc, "_session_context_writer", lambda sid, p: events.append((sid, p)))
    fresh = ClosableOrigin()
    oc.set_session_origin(ClosableOrigin(), lambda: fresh, "sess-9")
    oc.remember_project_path("C:/x.opju")
    oc.drop_session_origin()
    assert events[-1] == ("sess-9", None)
    assert oc.get_remembered_project_path() is None
    assert oc.get_origin() is fresh


def test_contrast_crash_relaunch_still_reopens_project(tmp_path):
    # Guard that drop_session_origin is what suppresses the reopen: an
    # UNINTENTIONAL death keeps the existing crash-recovery behaviour.
    proj = tmp_path / "exp.opju"
    proj.write_bytes(b"x" * 64)
    fresh = ClosableOrigin()
    old = ClosableOrigin()
    oc.set_session_origin(old, lambda: fresh, "sess-1")
    oc.remember_project_path(str(proj))
    old.alive = False              # crashed, not closed via close_origin
    assert oc.get_origin() is fresh
    assert fresh.load_calls == [str(proj)]


# --- daemon PID store lookup (sys.modules, never a fresh import) -------------- #


def _restore_real_helpers(monkeypatch):
    monkeypatch.setattr(project, "_session_origin_pid", _REAL_SESSION_PID)
    monkeypatch.setattr(project, "_forget_session_origin_pid", _REAL_FORGET_PID)



def test_pid_read_from_main_module_when_daemon_runs_as_main(monkeypatch):
    _restore_real_helpers(monkeypatch)
    tls = threading.local()
    tls.pid = 1234
    fake_main = types.ModuleType("__main__")
    fake_main._real_pid_tls = tls
    monkeypatch.setitem(sys.modules, "__main__", fake_main)
    # Simulate `python -m origin_pro_mcp.daemon`: the package-qualified name
    # is NOT imported, and reading the PID must not import a second copy.
    monkeypatch.delitem(sys.modules, "origin_pro_mcp.daemon", raising=False)
    assert project._session_origin_pid() == 1234
    assert "origin_pro_mcp.daemon" not in sys.modules
    project._forget_session_origin_pid()
    assert tls.pid is None
    assert project._session_origin_pid() is None


def test_pid_read_from_imported_daemon_module(monkeypatch):
    _restore_real_helpers(monkeypatch)
    from origin_pro_mcp import daemon

    tls = threading.local()
    tls.pid = 5678
    monkeypatch.setattr(daemon, "_real_pid_tls", tls)
    monkeypatch.setitem(sys.modules, "__main__", types.ModuleType("__main__"))
    assert project._session_origin_pid() == 5678


def test_pid_store_shared_by_both_names_is_listed_once(monkeypatch):
    from origin_pro_mcp import daemon

    tls = threading.local()
    monkeypatch.setattr(daemon, "_real_pid_tls", tls)
    monkeypatch.setitem(sys.modules, "__main__", daemon)
    stores = project._daemon_pid_stores()
    assert len(stores) == 1 and stores[0] is tls


@pytest.mark.parametrize("bad", [None, 0, -3, "123", 1.5])
def test_invalid_stored_pid_is_ignored(monkeypatch, bad):
    _restore_real_helpers(monkeypatch)
    tls = threading.local()
    tls.pid = bad
    fake_main = types.ModuleType("__main__")
    fake_main._real_pid_tls = tls
    monkeypatch.setitem(sys.modules, "__main__", fake_main)
    monkeypatch.delitem(sys.modules, "origin_pro_mcp.daemon", raising=False)
    assert project._session_origin_pid() is None


def test_pid_is_thread_local(monkeypatch):
    # Another session's worker thread must never see this thread's PID.
    _restore_real_helpers(monkeypatch)
    tls = threading.local()
    tls.pid = 1111
    fake_main = types.ModuleType("__main__")
    fake_main._real_pid_tls = tls
    monkeypatch.setitem(sys.modules, "__main__", fake_main)
    monkeypatch.delitem(sys.modules, "origin_pro_mcp.daemon", raising=False)
    seen = {}
    t = threading.Thread(target=lambda: seen.setdefault("pid", project._session_origin_pid()))
    t.start()
    t.join()
    assert seen["pid"] is None


def test_terminate_pid_uses_daemon_terminator_from_sys_modules(monkeypatch):
    killed = []
    fake_main = types.ModuleType("__main__")
    fake_main.default_terminate_process = killed.append
    monkeypatch.setitem(sys.modules, "__main__", fake_main)
    monkeypatch.setattr(project, "_terminate_pid", _REAL_TERMINATE)
    project._terminate_pid(99)
    assert killed == [99]



def test_real_pid_alive_on_posix_self():
    import os

    if os.name == "nt":
        pytest.skip("POSIX branch")
    assert _REAL_PID_ALIVE(os.getpid()) is True
    assert _REAL_PID_ALIVE(2 ** 22 + 12345) is False



# --- daemon integration: Session.attached set by Pool.acquire ----------------- #


def _registry():
    from origin_pro_mcp import server  # noqa: F401 - registers tools
    from origin_pro_mcp.app import mcp
    return {n: t.fn for n, t in mcp._tool_manager._tools.items()}


def _call(session, name, **kwargs):
    done = threading.Event()
    box = {}

    def reply(resp):
        box["resp"] = resp
        done.set()

    session.submit("r1", name, kwargs, reply)
    assert done.wait(15), "worker did not reply"
    return box["resp"]


def test_pool_marks_attached_session_and_close_origin_refuses_it():
    from origin_pro_mcp.daemon import Pool

    iso_instances, att_instances = [], []

    def iso():
        o = ClosableOrigin()
        iso_instances.append(o)
        return o

    def att():
        o = ClosableOrigin()
        att_instances.append(o)
        return o

    pool = Pool(iso, _registry(), max_size=3, attach_factory=att)
    try:
        s_att = pool.acquire("A", attach=True)
        s_iso = pool.acquire("B", attach=False)
        assert s_att.attached is True
        assert s_iso.attached is False

        resp = _call(s_att, "close_origin")
        assert resp["ok"] is False
        assert "Refusing to close" in resp["error"]
        assert att_instances[0].exit_calls == 0

        resp = _call(s_iso, "close_origin")
        assert resp["ok"] is True, resp
        assert "Closed this session's Origin" in resp["result"]
        assert iso_instances[0].exit_calls == 1
    finally:
        pool.stop_all()


def test_second_attach_request_falls_back_to_isolated_and_is_not_attached():
    from origin_pro_mcp.daemon import Pool

    pool = Pool(lambda: ClosableOrigin(), _registry(), max_size=3,
                attach_factory=lambda: ClosableOrigin())
    try:
        assert pool.acquire("A", attach=True).attached is True
        assert pool.acquire("B", attach=True).attached is False
    finally:
        pool.stop_all()
