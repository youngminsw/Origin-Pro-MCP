import threading

import pytest

from origin_pro_mcp import origin_connection


WSL_MSG = (
    "Origin Pro COM automation requires Windows Python with pywin32. "
    "Run this MCP server on Windows, not WSL/Linux."
)


class StubOrigin:
    """Minimal stand-in: no Visible property, so _connection_alive() is True."""


class DeadProxy:
    """Mimics a stale COM proxy: accessing Visible raises like a dead RPC."""

    @property
    def Visible(self):
        raise OSError("The object invoked has disconnected from its clients.")


def test_dead_proxy_with_factory_relaunches_via_factory_not_applicationsi():
    # Daemon path: a session binds its instance WITH a factory. When the proxy
    # dies, get_origin must re-run the session's own factory (relaunch an
    # isolated instance), NOT fall back to ApplicationSI (which on WSL would
    # raise the win32com import error, and on Windows could hijack the user's
    # open Origin).
    fresh = StubOrigin()
    calls = []

    def factory():
        calls.append(1)
        return fresh

    origin_connection.set_session_origin(DeadProxy(), factory)
    got = origin_connection.get_origin()
    assert got is fresh          # relaunched via the factory
    assert calls == [1]          # factory was used (no ApplicationSI Dispatch)


@pytest.fixture(autouse=True)
def _isolate_thread_local():
    """Guarantee no proxy leaks between tests on the main thread."""
    origin_connection.clear_session_origin()
    try:
        yield
    finally:
        origin_connection.clear_session_origin()


def test_set_then_get_returns_injected_instance():
    # With a live proxy set, get_origin() returns it without importing win32com
    # or attempting a Dispatch — on WSL any connect attempt raises RuntimeError,
    # so a clean identity return is itself proof that no connect was tried.
    stub = StubOrigin()
    origin_connection.set_session_origin(stub)
    assert origin_connection.get_origin() is stub


def test_clear_drops_proxy_and_forces_reconnect(monkeypatch):
    # Make win32com import fail so a reconnect surfaces the WSL RuntimeError.
    _force_no_win32com(monkeypatch)

    origin_connection.set_session_origin(StubOrigin())
    origin_connection.clear_session_origin()

    with pytest.raises(RuntimeError) as exc:
        origin_connection.get_origin()
    assert str(exc.value) == WSL_MSG


def test_dead_proxy_triggers_reconnect(monkeypatch):
    # A stale proxy must be discarded; the reconnect path then surfaces here.
    _force_no_win32com(monkeypatch)

    origin_connection.set_session_origin(DeadProxy())

    with pytest.raises(RuntimeError) as exc:
        origin_connection.get_origin()
    assert str(exc.value) == WSL_MSG
    # The dead proxy must have been cleared from the thread-local store.
    assert not hasattr(origin_connection._state, "origin")


def test_thread_isolation():
    """Each thread sees only its own proxy — the whole point of the seam."""
    main_proxy = StubOrigin()
    origin_connection.set_session_origin(main_proxy)

    worker_proxy = StubOrigin()
    result: dict[str, object] = {}
    started = threading.Event()

    def worker():
        # Worker starts with no proxy of its own (separate thread-local slot).
        result["initial_get_raises"] = not hasattr(
            origin_connection._state, "origin"
        )
        origin_connection.set_session_origin(worker_proxy)
        seen = origin_connection.get_origin()
        result["worker_sees_own"] = seen is worker_proxy
        result["worker_sees_main"] = seen is main_proxy
        started.set()

    t = threading.Thread(target=worker)
    t.start()
    t.join()
    assert started.is_set()

    # Worker had its own empty slot, saw its own proxy, never the main thread's.
    assert result["initial_get_raises"] is True
    assert result["worker_sees_own"] is True
    assert result["worker_sees_main"] is False

    # Main thread is untouched by the worker's set/get.
    assert origin_connection.get_origin() is main_proxy


def _force_no_win32com(monkeypatch):
    """Make `import win32com.client` raise ModuleNotFoundError inside get_origin."""
    import builtins

    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "win32com.client" or name.startswith("win32com"):
            raise ModuleNotFoundError(f"No module named {name!r}")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
