"""Reconnect recovery: after the COM proxy dies, get_origin relaunches this
session's own instance via the factory AND reopens the last on-disk project
(P0 — 'reconnect loses the whole project')."""

import pytest

from fakes import FakeOrigin
from origin_pro_mcp import origin_connection as oc


class _DeadProxy:
    """A COM proxy whose liveness ping raises — i.e. Origin crashed/closed."""

    def __init__(self):
        self._dead = False

    @property
    def Visible(self):
        if self._dead:
            raise RuntimeError("RPC server is unavailable (Origin died)")
        return 1


class _RecordingOrigin(FakeOrigin):
    def __init__(self):
        super().__init__()
        self.loaded = []

    def Load(self, path):
        self.loaded.append(path)
        return True


@pytest.fixture
def clean_state():
    oc.clear_session_origin()
    yield
    oc.clear_session_origin()


def _factory_returning(instances):
    it = iter(instances)

    def factory():
        return next(it)

    return factory


def test_reconnect_reopens_remembered_project(tmp_path, clean_state):
    proj = tmp_path / "work.opju"
    proj.write_text("stub")
    fresh = _RecordingOrigin()
    dead = _DeadProxy()

    oc.set_session_origin(dead, _factory_returning([fresh]))
    oc.remember_project_path(str(proj))

    dead._dead = True  # simulate the crash
    got = oc.get_origin()

    assert got is fresh                       # relaunched this session's instance
    assert fresh.loaded == [str(proj)]        # and reopened the last project


def test_reconnect_without_remembered_project_does_not_load(tmp_path, clean_state):
    fresh = _RecordingOrigin()
    dead = _DeadProxy()
    oc.set_session_origin(dead, _factory_returning([fresh]))

    dead._dead = True
    got = oc.get_origin()

    assert got is fresh
    assert fresh.loaded == []                 # nothing to recover on first-ever launch


def test_reconnect_skips_missing_project_file(tmp_path, clean_state):
    fresh = _RecordingOrigin()
    dead = _DeadProxy()
    oc.set_session_origin(dead, _factory_returning([fresh]))
    oc.remember_project_path(str(tmp_path / "gone.opju"))  # never created

    dead._dead = True
    got = oc.get_origin()

    assert got is fresh
    assert fresh.loaded == []                 # missing file => no reopen attempt


def test_remember_none_forgets_path(clean_state):
    oc.remember_project_path("/mnt/c/x.opju")
    assert oc.get_remembered_project_path() == "/mnt/c/x.opju"
    oc.remember_project_path(None)
    assert oc.get_remembered_project_path() is None


def test_clear_session_forgets_project_path(clean_state):
    oc.set_session_origin(FakeOrigin(), lambda: FakeOrigin())
    oc.remember_project_path("/mnt/c/x.opju")
    oc.clear_session_origin()
    assert oc.get_remembered_project_path() is None
