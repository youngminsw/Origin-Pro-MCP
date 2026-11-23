"""_replace_with_retry: Windows refuses to replace a file another process has
open (a shim reading daemon.json, an antivirus scan). The lockfile and session
ledger writes retry briefly instead of failing silently."""
import pytest

from origin_pro_mcp import daemon


class FlakyReplace:
    def __init__(self, failures):
        self.failures = failures
        self.calls = 0

    def __call__(self, src, dst):
        self.calls += 1
        if self.calls <= self.failures:
            raise PermissionError(13, "Access is denied")


def test_retries_sharing_violation_then_succeeds(monkeypatch):
    monkeypatch.setattr(daemon.sys, "platform", "win32")
    flaky = FlakyReplace(failures=2)
    monkeypatch.setattr(daemon.os, "replace", flaky)
    slept = []
    daemon._replace_with_retry("a.tmp", "a", sleep=slept.append)
    assert flaky.calls == 3
    assert slept == list(daemon._REPLACE_RETRY_DELAYS[:2])


def test_gives_up_after_the_last_retry(monkeypatch):
    monkeypatch.setattr(daemon.sys, "platform", "win32")
    flaky = FlakyReplace(failures=99)
    monkeypatch.setattr(daemon.os, "replace", flaky)
    with pytest.raises(PermissionError):
        daemon._replace_with_retry("a.tmp", "a", sleep=lambda _s: None)
    assert flaky.calls == len(daemon._REPLACE_RETRY_DELAYS) + 1


def test_other_errors_are_not_retried(monkeypatch):
    monkeypatch.setattr(daemon.sys, "platform", "win32")
    calls = []

    def missing(src, dst):
        calls.append(src)
        raise FileNotFoundError(src)

    monkeypatch.setattr(daemon.os, "replace", missing)
    with pytest.raises(FileNotFoundError):
        daemon._replace_with_retry("a.tmp", "a", sleep=lambda _s: None)
    assert len(calls) == 1


def test_posix_replaces_once_without_retry(monkeypatch):
    monkeypatch.setattr(daemon.sys, "platform", "linux")
    flaky = FlakyReplace(failures=1)
    monkeypatch.setattr(daemon.os, "replace", flaky)
    with pytest.raises(PermissionError):
        daemon._replace_with_retry("a.tmp", "a", sleep=lambda _s: None)
    assert flaky.calls == 1


def test_lockfile_write_survives_a_brief_lock(tmp_path, monkeypatch):
    monkeypatch.setattr(daemon.sys, "platform", "win32")
    real = daemon.os.replace
    state = {"n": 0}

    def once_locked(src, dst):
        state["n"] += 1
        if state["n"] == 1:
            raise PermissionError(13, "Access is denied")
        real(src, dst)

    monkeypatch.setattr(daemon.os, "replace", once_locked)  # one 10 ms retry
    path = str(tmp_path / "daemon.json")
    daemon.write_lockfile(path, port=1, token="t", pid=2, child_pids=[3])
    assert daemon.read_lockfile(path)["child_pids"] == [3]
