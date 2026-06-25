"""WSL-safe regression tests for the Windows real-Origin PID capture.

The real factory cannot run here (no COM), but the two pieces that caused real
bugs on Origin 2020 are testable: (1) parsing Origin64.exe PIDs from tasklist
output, and (2) reading the snapshot-diff PID back from the thread-local that
the factory stashes it on. Validated live against Origin 2020; these lock it in.
"""
import subprocess
import threading

from origin_pro_mcp import daemon


def test_origin_process_pids_parses_tasklist_csv(monkeypatch):
    csv = (
        '"Origin64.exe","1034064","Console","1","250,000 K"\r\n'
        '"Origin64.exe","1140532","Console","1","250,000 K"\r\n'
    )

    def fake_run(*a, **k):
        return subprocess.CompletedProcess(a, 0, stdout=csv, stderr="")

    monkeypatch.setattr(subprocess, "run", fake_run)
    assert daemon._origin_process_pids() == {1034064, 1140532}


def test_origin_process_pids_empty(monkeypatch):
    msg = "INFO: No tasks are running which match the specified criteria.\r\n"
    monkeypatch.setattr(
        subprocess, "run",
        lambda *a, **k: subprocess.CompletedProcess(a, 0, stdout=msg, stderr=""),
    )
    assert daemon._origin_process_pids() == set()


def test_real_get_pid_reads_thread_local_stash():
    # The factory stashes the captured PID here; get_pid reads it back on the
    # same worker thread.
    daemon._real_pid_tls.pid = 4242
    try:
        assert daemon._real_origin_get_pid(object()) == 4242
    finally:
        del daemon._real_pid_tls.pid


def test_real_get_pid_defaults_to_none_when_unset():
    # A fresh thread has no stashed PID -> None ("do not force-kill"), never self.
    result = {}

    def worker():
        result["pid"] = daemon._real_origin_get_pid(object())

    t = threading.Thread(target=worker)
    t.start()
    t.join()
    assert result["pid"] is None


def test_origin_visible_mode(monkeypatch):
    # Default: visible (watch the agent work).
    monkeypatch.delenv("ORIGIN_PRO_MCP_VISIBLE", raising=False)
    assert daemon._origin_visible() == 1
    # Truthy values -> visible.
    for v in ("1", "true", "yes", "on", "visible"):
        monkeypatch.setenv("ORIGIN_PRO_MCP_VISIBLE", v)
        assert daemon._origin_visible() == 1
    # Falsey/hidden values -> invisible.
    for v in ("0", "false", "no", "off", "hidden", "invisible", "FALSE", " 0 "):
        monkeypatch.setenv("ORIGIN_PRO_MCP_VISIBLE", v)
        assert daemon._origin_visible() == 0


def test_session_has_com_init_helpers():
    # The worker thread must initialize a COM apartment (Origin 2020 needs it);
    # the helpers are guarded no-ops where pythoncom is unavailable (WSL).
    assert hasattr(daemon.Session, "_com_initialize")
    assert hasattr(daemon.Session, "_com_uninitialize")
    # On WSL (no pythoncom) initialize returns False without raising.
    assert daemon.Session._com_initialize() is False
    daemon.Session._com_uninitialize()  # must not raise
