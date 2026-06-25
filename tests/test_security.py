"""POSIX lockfile/dir hardening tests (SECURITY 4 + 5).

The daemon discovery lockfile holds the auth token. It must never be
world-readable at any instant, must live in a user-private 0700 dir, and the
singleton lock must refuse a symlinked path (defeating a symlink attack).
"""
from __future__ import annotations

import os
import stat
import sys

import pytest

from origin_pro_mcp import daemon as d
from origin_pro_mcp.daemon import (
    SingletonGuard,
    default_lockfile_path,
    write_lockfile,
)

posix_only = pytest.mark.skipif(
    sys.platform == "win32", reason="POSIX file-permission semantics"
)


@posix_only
def test_written_lockfile_is_0600(tmp_path):
    path = str(tmp_path / "daemon.json")
    write_lockfile(path, port=1, token="secret-token", pid=1, child_pids=[])
    mode = stat.S_IMODE(os.stat(path).st_mode)
    assert mode == 0o600, oct(mode)


@posix_only
def test_lockfile_tempfile_never_world_readable(tmp_path, monkeypatch):
    """The temp file must already be 0600 at the instant just before the
    atomic os.replace — never created world-readable."""
    captured: dict = {}
    real_replace = os.replace

    def spy(src, dst):
        captured["mode"] = stat.S_IMODE(os.stat(src).st_mode)
        return real_replace(src, dst)

    monkeypatch.setattr(d.os, "replace", spy)
    write_lockfile(str(tmp_path / "daemon.json"), port=1, token="t", pid=1,
                   child_pids=[])
    assert captured["mode"] == 0o600, oct(captured.get("mode", -1))


@posix_only
def test_default_lockfile_dir_is_private_0700(monkeypatch, tmp_path):
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(tmp_path / "run"))
    path = default_lockfile_path()
    directory = os.path.dirname(path)
    mode = stat.S_IMODE(os.stat(directory).st_mode)
    assert mode == 0o700, oct(mode)
    assert os.stat(directory).st_uid == os.getuid()


@posix_only
def test_default_lockfile_dir_without_xdg_is_private(monkeypatch, tmp_path):
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.delenv("XDG_RUNTIME_DIR", raising=False)
    monkeypatch.setattr(d.tempfile, "gettempdir", lambda: str(tmp_path))
    path = default_lockfile_path()
    directory = os.path.dirname(path)
    mode = stat.S_IMODE(os.stat(directory).st_mode)
    assert mode == 0o700, oct(mode)
    # uid-namespaced so two users never share the dir
    assert str(os.getuid()) in os.path.basename(directory)


@posix_only
def test_singleton_refuses_symlinked_lock(tmp_path):
    real = tmp_path / "real.lock"
    real.write_text("")
    link = tmp_path / "link.lock"
    os.symlink(str(real), str(link))

    guard = SingletonGuard(str(link))
    try:
        # O_NOFOLLOW must make acquiring a symlinked lock path fail (refused).
        assert guard.acquire() is False
    finally:
        guard.release()


@posix_only
def test_singleton_still_works_for_regular_path(tmp_path):
    path = str(tmp_path / "ok.lock")
    first = SingletonGuard(path)
    second = SingletonGuard(path)
    try:
        assert first.acquire() is True
        assert second.acquire() is False
        first.release()
        assert second.acquire() is True
    finally:
        first.release()
        second.release()
