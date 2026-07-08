"""Session-ledger + lifecycle-notice tests (COM-free, real loopback sockets).

Exercises the sidecar ledger (sessions.json), the one-shot mint notice attached
to the first successful response, the load_project collision warning, and the
shim's notice-append — all through injected fakes, never COM.
"""
from __future__ import annotations

import json
import threading
import time
from contextlib import contextmanager

import pytest

from fakes import FakeOrigin

from origin_pro_mcp import transport
from origin_pro_mcp.daemon import Daemon, read_sessions, write_sessions


class _Factory:
    """Hands out a fresh FakeOrigin with a unique fake pid on each call."""

    def __init__(self, start_pid: int = 90000):
        self.created: list[FakeOrigin] = []
        self._next = start_pid
        self._lock = threading.Lock()

    def __call__(self) -> FakeOrigin:
        with self._lock:
            fake = FakeOrigin()
            fake.fake_pid = self._next
            self._next += 1
            self.created.append(fake)
            return fake

    @staticmethod
    def get_pid(instance) -> int:
        return instance.fake_pid


def _attach_get_pid(_instance) -> None:
    return None


def _real_registry() -> dict:
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp

    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


@contextmanager
def run_daemon(tmp_path, *, generation="gen-A", live=(), attach=False,
               max_size=3, reconnect_grace=0.0):
    """Start a Daemon with the ledger knobs wired, yield a small test handle,
    and tear everything down."""
    factory = _Factory()
    attach_factory = _Factory(start_pid=70000) if attach else None
    daemon = Daemon()
    conns: list = []
    live_set = set(live)
    ok = daemon.start(
        origin_factory=factory,
        registry=_real_registry(),
        max_size=max_size,
        host="127.0.0.1",
        port=0,
        get_pid=_Factory.get_pid,
        lockfile_path=str(tmp_path / "daemon.json"),
        reconnect_grace=reconnect_grace,
        generation=generation,
        live_origin_pids=lambda: set(live_set),
        attach_factory=attach_factory,
        attach_get_pid=_attach_get_pid if attach else None,
    )
    assert ok is True

    def client(session_id):
        conn = transport.connect(daemon.host, daemon.port, daemon.token,
                                 session_id=session_id)
        conn.settimeout(5.0)
        conns.append(conn)
        return conn

    def call(conn, request_id, name, kwargs, attach=False):
        conn.send_frame({"type": "request", "request_id": request_id,
                         "name": name, "kwargs": kwargs, "attach": attach})
        return conn.recv_frame()

    try:
        yield {"daemon": daemon, "factory": factory, "client": client,
               "call": call, "sessions_path": str(tmp_path / "sessions.json")}
    finally:
        for conn in conns:
            conn.close()
        daemon.stop()


def _seed(tmp_path, generation, sessions):
    write_sessions(str(tmp_path / "sessions.json"),
                   {"generation": generation, "sessions": sessions})


# --------------------------------------------------------------------------- #
# Ledger write points                                                         #
# --------------------------------------------------------------------------- #


def test_ledger_written_on_spawn(tmp_path):
    with run_daemon(tmp_path) as h:
        conn = h["client"]("sess-1")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
        assert resp["ok"] is True, resp
        ledger = read_sessions(h["sessions_path"])
    assert ledger["generation"] == "gen-A"
    entry = ledger["sessions"]["sess-1"]
    assert entry["pid"] == 90000
    assert entry["attach"] is False
    assert entry["ended"] is False
    assert entry["project"] is None


def test_ledger_written_on_project_change(tmp_path):
    target = str(tmp_path / "study.opju")
    with run_daemon(tmp_path) as h:
        conn = h["client"]("sess-1")
        # A fresh FakeOrigin has open windows, so save_project proceeds and calls
        # remember_project_path -> the ledger writer records the path.
        resp = h["call"](conn, "r1", "save_project", {"file_path": target})
        assert resp["ok"] is True, resp
        ledger = read_sessions(h["sessions_path"])
    assert ledger["sessions"]["sess-1"]["project"] == target


def test_ledger_marks_ended_on_reap(tmp_path):
    with run_daemon(tmp_path, reconnect_grace=0.0) as h:
        conn = h["client"]("sess-1")
        assert h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})["ok"]
        conn.close()  # last connection closing schedules an immediate reap
        deadline = time.monotonic() + 5.0
        entry = None
        while time.monotonic() < deadline:
            entry = read_sessions(h["sessions_path"])["sessions"].get("sess-1")
            if entry and entry.get("ended"):
                break
            time.sleep(0.02)
    assert entry is not None and entry["ended"] is True
    # The entry is KEPT (a detached Origin may still be alive), pid preserved.
    assert entry["pid"] == 90000


# --------------------------------------------------------------------------- #
# Mint notices                                                                #
# --------------------------------------------------------------------------- #


def test_notice_prev_generation_live(tmp_path):
    _seed(tmp_path, "gen-OLD", {
        "sess-1": {"pid": 4242, "project": r"C:\proj.opju",
                   "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live={4242}) as h:
        conn = h["client"]("sess-1")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
    assert resp["ok"] is True
    notice = resp["notice"]
    assert "pid 4242" in notice
    assert "still open with project C:\\proj.opju" in notice
    assert "work in this fresh instance" in notice


def test_notice_prev_generation_dead(tmp_path):
    _seed(tmp_path, "gen-OLD", {
        "sess-1": {"pid": 4242, "project": r"C:\proj.opju",
                   "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live=set()) as h:
        conn = h["client"]("sess-1")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
    notice = resp["notice"]
    assert "daemon restarted" in notice
    assert r"load_project(r'C:\proj.opju')" in notice


def test_notice_ghosts(tmp_path):
    _seed(tmp_path, "gen-OLD", {
        "ghost-1": {"pid": 5001, "project": "/data/a.opju",
                    "attach": False, "ts": 1.0},
        "ghost-2": {"pid": 5002, "project": "/data/b.opju",
                    "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live={5001, 5002}) as h:
        conn = h["client"]("sess-new")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
    notice = resp["notice"]
    assert "2 leftover Origin window(s)" in notice
    assert "a.opju" in notice and "b.opju" in notice
    assert "ORIGIN_PRO_MCP_SWEEP_ORPHANS=1" in notice


def test_notice_attach_granted(tmp_path):
    with run_daemon(tmp_path, attach=True) as h:
        conn = h["client"]("sess-1")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"},
                         attach=True)
    assert resp["ok"] is True
    assert "attached to the user's open Origin" in resp["notice"]
    assert "Autosave and force-recovery are disabled" in resp["notice"]


def test_notice_attach_fallback(tmp_path):
    with run_daemon(tmp_path, attach=True) as h:
        c1 = h["client"]("sess-1")
        assert h["call"](c1, "r1", "run_labtalk", {"script": "x=1;"},
                         attach=True)["ok"]  # grabs the single attach slot
        c2 = h["client"]("sess-2")
        resp = h["call"](c2, "r2", "run_labtalk", {"script": "y=1;"},
                         attach=True)
    assert "another session already holds the user's Origin" in resp["notice"]


def test_notice_one_shot(tmp_path):
    _seed(tmp_path, "gen-OLD", {
        "sess-1": {"pid": 4242, "project": r"C:\proj.opju",
                   "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live={4242}) as h:
        conn = h["client"]("sess-1")
        first = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
        second = h["call"](conn, "r2", "run_labtalk", {"script": "x=2;"})
    assert "notice" in first
    assert "notice" not in second  # delivered at most once per session


def test_error_response_carries_no_notice_but_keeps_it_pending(tmp_path):
    _seed(tmp_path, "gen-OLD", {
        "sess-1": {"pid": 4242, "project": r"C:\proj.opju",
                   "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live={4242}) as h:
        conn = h["client"]("sess-1")
        # First call fails (unknown tool) -> no notice on the error frame.
        err = h["call"](conn, "r1", "does_not_exist", {})
        assert err["ok"] is False
        assert "notice" not in err
        # The notice rides the next SUCCESSFUL response instead.
        ok = h["call"](conn, "r2", "run_labtalk", {"script": "x=1;"})
    assert ok["ok"] is True
    assert "notice" in ok


# --------------------------------------------------------------------------- #
# load_project collision warning                                              #
# --------------------------------------------------------------------------- #


def test_load_project_collision_warning(tmp_path):
    proj = tmp_path / "shared.opju"
    proj.write_bytes(b"x" * 4096)  # a real, non-empty project file on disk
    _seed(tmp_path, "gen-OLD", {
        "other": {"pid": 6001, "project": str(proj),
                  "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live={6001}) as h:
        conn = h["client"]("loader")
        resp = h["call"](conn, "r1", "load_project", {"file_path": str(proj)})
    assert resp["ok"] is True, resp
    assert "WARNING: another Origin instance (pid 6001)" in resp["result"]


def test_load_project_no_warning_when_other_dead(tmp_path):
    proj = tmp_path / "shared.opju"
    proj.write_bytes(b"x" * 4096)
    _seed(tmp_path, "gen-OLD", {
        "other": {"pid": 6001, "project": str(proj),
                  "attach": False, "ts": 1.0},
    })
    with run_daemon(tmp_path, generation="gen-NEW", live=set()) as h:
        conn = h["client"]("loader")
        resp = h["call"](conn, "r1", "load_project", {"file_path": str(proj)})
    assert resp["ok"] is True
    assert "WARNING" not in resp["result"]


# --------------------------------------------------------------------------- #
# Corrupt ledger tolerance                                                    #
# --------------------------------------------------------------------------- #


def test_corrupt_ledger_tolerated(tmp_path):
    (tmp_path / "sessions.json").write_text("{ this is not valid json ][")
    with run_daemon(tmp_path) as h:
        conn = h["client"]("sess-1")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
        assert resp["ok"] is True  # corrupt sidecar never breaks a call
        ledger = read_sessions(h["sessions_path"])
    # It was overwritten with a valid ledger.
    assert ledger["sessions"]["sess-1"]["pid"] == 90000


def test_read_sessions_missing_returns_empty(tmp_path):
    data = read_sessions(str(tmp_path / "nope.json"))
    assert data == {"generation": None, "sessions": {}}


def test_malformed_entry_ledger_never_breaks_first_call(tmp_path):
    # Valid JSON but a session ENTRY that is not an object: the mint path must
    # not raise, and the first tool call must succeed (regression: M1).
    (tmp_path / "sessions.json").write_text(
        '{"generation": "old-gen", "sessions": {"sess-1": "notadict", '
        '"sess-2": 42}}'
    )
    with run_daemon(tmp_path) as h:
        conn = h["client"]("sess-1")
        resp = h["call"](conn, "r1", "run_labtalk", {"script": "x=1;"})
        assert resp["ok"] is True
        ledger = read_sessions(h["sessions_path"])
    assert ledger["sessions"]["sess-1"]["pid"] == 90000
    # The malformed entries were dropped, not propagated.
    assert "sess-2" not in ledger["sessions"] or isinstance(
        ledger["sessions"].get("sess-2"), dict)


# --------------------------------------------------------------------------- #
# Shim notice append                                                          #
# --------------------------------------------------------------------------- #


class _EchoConn:
    """A fake connection that echoes one canned response (with a notice) for the
    request_id the shim generates."""

    def __init__(self, ok=True, result="Done", notice="Note: hello"):
        self._ok = ok
        self._result = result
        self._notice = notice
        self._req_id = None

    def send_frame(self, frame):
        self._req_id = frame["request_id"]

    def recv_frame(self):
        return {"type": "response", "request_id": self._req_id,
                "ok": self._ok, "result": self._result,
                "error": None, "notice": self._notice}


def test_shim_appends_notice():
    from origin_pro_mcp.shim import ShimClient

    client = ShimClient.__new__(ShimClient)  # skip __init__ (no daemon needed)
    client._session_id = "sid"
    client._call_lock = threading.Lock()
    client._attach = False
    conn = _EchoConn(result="Loaded project: X", notice="Note: hello world")
    out = client._exchange(conn, "load_project", {"file_path": "X"})
    assert out == "Loaded project: X\n\n[origin-mcp] Note: hello world"


def test_shim_no_notice_returns_bare_result():
    from origin_pro_mcp.shim import ShimClient

    client = ShimClient.__new__(ShimClient)
    client._session_id = "sid"
    client._call_lock = threading.Lock()
    client._attach = False
    conn = _EchoConn(result="plain", notice=None)
    out = client._exchange(conn, "run_labtalk", {"script": "x=1;"})
    assert out == "plain"
