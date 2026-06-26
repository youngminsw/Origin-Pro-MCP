"""Daemon core: per-session worker pool, watchdog, and singleton guard.

This phase (3b) is entirely COM-free. The daemon owns:

* :class:`Session`  - a dedicated worker thread + request queue. The worker
  creates its Origin instance via an injected ``origin_factory`` and binds it
  to *its own* thread via ``set_session_origin`` — so the real tool bodies hit
  that session's instance (model B1: one isolated Origin per session).
* :class:`Pool`     - up to ``max_size`` sessions; rejects the overflow session
  with an actionable :class:`PoolFull` error (never hangs).
* :class:`Watchdog` - a daemon-level thread that arms reap deadlines and, on
  expiry, performs an out-of-band kill via an injectable ``terminate_process``
  hook. It only ever touches PIDs — never a COM proxy.
* :class:`SingletonGuard` - an exclusive lockfile lock so a second daemon exits.
* :class:`Daemon`   - ties it together over the loopback-TCP transport.
"""
from __future__ import annotations

import json
import os
import queue
import secrets
import stat
import sys
import tempfile
import threading
import time
from typing import Callable, Optional

from .origin_connection import clear_session_origin, set_session_origin
from .transport import Connection, FrameError, TcpServer

OriginFactory = Callable[[], object]
GetPid = Callable[[object], Optional[int]]
TerminateProcess = Callable[[int], None]
ReplyFn = Callable[[dict], None]

POOL_MAX_DEFAULT = 3
DEFAULT_RECONNECT_GRACE: float = 3.0  # seconds; env: ORIGIN_PRO_MCP_RECONNECT_GRACE


# --------------------------------------------------------------------------- #
# OS-level process termination (default watchdog hook)                         #
# --------------------------------------------------------------------------- #


def default_terminate_process(pid: int) -> None:
    """Hard-kill a process by PID — out-of-band, no COM involved.

    POSIX uses ``SIGKILL``; Windows opens the process and calls
    ``TerminateProcess``. Tests inject a recording fake instead.
    """
    if sys.platform == "win32":
        import ctypes  # local import: Windows-only path

        PROCESS_TERMINATE = 0x0001
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
        if handle:
            try:
                ctypes.windll.kernel32.TerminateProcess(handle, 1)
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
    else:
        import signal

        os.kill(pid, signal.SIGKILL)


def _default_get_pid(_instance: object) -> Optional[int]:
    """Default child-PID resolver: UNKNOWN.

    Returning ``None`` (not ``os.getpid()``) is a safety-critical contract: an
    unknown PID means "do not force-kill", never "kill the daemon itself". A
    real PID is wired in only when the production factory can resolve the
    spawned ``Origin.exe`` process id (see ``_real_origin_get_pid``).
    """
    return None


# --------------------------------------------------------------------------- #
# Session                                                                      #
# --------------------------------------------------------------------------- #


class Session:
    """A single Origin session: one worker thread, one instance, one queue.

    The worker thread creates the Origin instance and calls
    ``set_session_origin`` on itself, so every tool dispatched here resolves to
    this session's instance via the thread-local seam.
    """

    _STOP = object()
    _REAP = object()

    def __init__(self, session_id: str, origin_factory: OriginFactory,
                 registry: dict, get_pid: Optional[GetPid] = None):
        self.session_id = session_id
        self._factory = origin_factory
        self._registry = registry
        self._get_pid = get_pid or _default_get_pid
        self._queue: "queue.Queue" = queue.Queue()
        self._ready = threading.Event()
        self.pid: Optional[int] = None
        self.instance: object = None
        self.saved_recovery_path: Optional[str] = None
        self.reaping: bool = False  # set once a reap has COMMITTED (stage 1)
        self._start_error: Optional[BaseException] = None
        self._thread = threading.Thread(
            target=self._run, name=f"session-{session_id}", daemon=True
        )

    def start(self, ready_timeout: float = 10.0) -> None:
        self._thread.start()
        if not self._ready.wait(timeout=ready_timeout):
            raise TimeoutError(
                f"session {self.session_id!r} worker did not start in time"
            )
        if self._start_error is not None:
            raise self._start_error

    @staticmethod
    def _com_initialize() -> bool:
        """Initialize this thread's COM apartment (STA). Windows-only, guarded.

        DispatchEx and every later COM call on this worker thread require the
        thread to have called ``CoInitialize`` first; on a non-main thread
        without it the first COM call raises. No-op (returns False) where
        ``pythoncom`` is unavailable (WSL/tests with a fake factory).
        """
        try:
            import pythoncom

            pythoncom.CoInitialize()  # STA apartment (what Origin expects)
            return True
        except Exception:
            return False

    @staticmethod
    def _com_uninitialize() -> None:
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _run(self) -> None:
        com_inited = self._com_initialize()
        try:
            try:
                self.instance = self._factory()
            except BaseException as exc:  # factory/COM failure: surface to start()
                self._start_error = exc
                self._ready.set()
                return
            # Pass the factory so a dead proxy relaunches THIS session's own
            # isolated instance, never the shared ApplicationSI (which could
            # hijack the user's open Origin).
            set_session_origin(self.instance, self._factory)
            try:
                pid = self._get_pid(self.instance)
                self.pid = int(pid) if pid is not None else None
            except Exception:
                # Unknown PID -> None ("do not force-kill"), NEVER os.getpid().
                self.pid = None
            self._ready.set()
            try:
                while True:
                    item = self._queue.get()
                    if item is self._STOP:
                        break
                    if item[0] is self._REAP:
                        _, recovery_dir, getter, on_done = item
                        try:
                            self._graceful_reap(recovery_dir, getter)
                        finally:
                            if on_done is not None:
                                on_done()
                        break  # the session is reaped; the worker exits
                    request_id, name, kwargs, reply_fn = item
                    reply_fn(self._dispatch(request_id, name, kwargs))
            finally:
                clear_session_origin()
        finally:
            if com_inited:
                self._com_uninitialize()

    def _graceful_reap(self, recovery_dir: str, getter) -> None:
        """Stage 1 of a reap, ON THIS WORKER THREAD (so COM affinity holds).

        Best-effort: resolve the session's open-project path (via the injected
        ``getter``, on this thread), pick a collision-safe recovery path, save
        the project to it, then close the instance. Every step is wrapped — a
        failure here just means the watchdog's force-kill reclaims the slot.
        """
        project_path = None
        if getter is not None:
            try:
                project_path = getter(self.instance)
            except Exception:
                project_path = None
        path = recovery_path(recovery_dir, self.session_id, project_path)
        try:
            os.makedirs(recovery_dir, exist_ok=True)
        except OSError:
            pass
        try:
            self.instance.Save(path)
            self.saved_recovery_path = path
        except Exception:
            pass
        for closer in ("Exit", "Close"):
            fn = getattr(self.instance, closer, None)
            if callable(fn):
                try:
                    fn()
                except Exception:
                    pass
                break

    def submit_reap(self, recovery_dir: str, getter,
                    on_done: Optional[Callable[[], None]] = None) -> None:
        """Enqueue the graceful reap task onto this session's worker thread."""
        self._queue.put((self._REAP, recovery_dir, getter, on_done))

    def _dispatch(self, request_id: str, name: str, kwargs: dict) -> dict:
        try:
            fn = self._registry.get(name)
            if fn is None:
                raise KeyError(f"unknown tool: {name!r}")
            result = fn(**(kwargs or {}))
            if result is not None and not isinstance(result, str):
                result = json.dumps(result)
            return {"type": "response", "request_id": request_id, "ok": True,
                    "result": result, "error": None}
        except Exception as exc:
            return {"type": "response", "request_id": request_id, "ok": False,
                    "result": None, "error": f"{type(exc).__name__}: {exc}"}

    def submit(self, request_id: str, name: str, kwargs: dict,
               reply_fn: ReplyFn) -> None:
        self._queue.put((request_id, name, kwargs, reply_fn))

    def stop(self, join_timeout: float = 5.0) -> None:
        self._queue.put(self._STOP)
        self._thread.join(timeout=join_timeout)

    def force_close(self) -> None:
        """Best-effort teardown of a half-started session (rollback path).

        Signals the worker to stop and closes the instance if one exists, so a
        start-timeout/start-failure never leaks an untracked Origin process. It
        never joins (the worker may be wedged in the factory).
        """
        self._queue.put(self._STOP)
        inst = self.instance
        if inst is None:
            return
        for closer in ("Exit", "Close"):
            fn = getattr(inst, closer, None)
            if callable(fn):
                try:
                    fn()
                except Exception:
                    pass
                break


# --------------------------------------------------------------------------- #
# Pool                                                                         #
# --------------------------------------------------------------------------- #


class PoolFull(RuntimeError):
    """Raised when a new session is requested but the pool is at capacity."""


class Pool:
    """A bounded set of :class:`Session` workers (default 3)."""

    def __init__(self, origin_factory: OriginFactory, registry: dict,
                 max_size: int = POOL_MAX_DEFAULT,
                 get_pid: Optional[GetPid] = None,
                 start_timeout: float = 10.0):
        self._factory = origin_factory
        self._registry = registry
        self._max_size = max_size
        self._get_pid = get_pid
        self._start_timeout = start_timeout
        self._sessions: dict[str, Session] = {}
        self._lock = threading.Lock()

    @property
    def max_size(self) -> int:
        return self._max_size

    def _full_message(self) -> str:
        n = self._max_size
        return (
            f"Origin pool full ({n}/{n}). "
            "Close another Origin MCP session and retry."
        )

    def acquire(self, session_id: str) -> Session:
        # Phase 1 (under lock): reuse a live session, refuse a reaping one, and
        # RESERVE a slot for a new session — but do NOT call Session.start()
        # here (it can block up to ready_timeout; holding the lock would stall
        # every other pool op and, on timeout, orphan the worker).
        with self._lock:
            existing = self._sessions.get(session_id)
            if existing is not None and not existing.reaping:
                return existing
            if existing is not None:  # a committed reap is in progress
                self._sessions.pop(session_id, None)
            if len(self._sessions) >= self._max_size:
                raise PoolFull(self._full_message())
            session = Session(
                session_id, self._factory, self._registry, self._get_pid
            )
            self._sessions[session_id] = session  # reserve the slot
        # Phase 2 (OUTSIDE the lock): start the worker; commit on ready, roll
        # back on timeout/failure (and tear down the half-started instance).
        try:
            session.start(ready_timeout=self._start_timeout)
        except BaseException:
            with self._lock:
                if self._sessions.get(session_id) is session:
                    self._sessions.pop(session_id, None)
            session.force_close()
            raise
        return session

    def release(self, session_id: str) -> None:
        with self._lock:
            session = self._sessions.pop(session_id, None)
        if session is not None:
            session.stop()

    def get(self, session_id: str) -> Optional[Session]:
        with self._lock:
            return self._sessions.get(session_id)

    def discard(self, session_id: str,
                expected: Optional[Session] = None) -> Optional[Session]:
        """Drop a session from the pool WITHOUT joining its worker thread.

        Used by the reaper so slot reclamation never blocks on a wedged worker
        (the graceful path's worker exits on its own; the watchdog path's real
        process has already been killed). When ``expected`` is given, only drop
        the slot if it still holds that exact session — so a reconnect that
        replaced a reaping session with a FRESH one is never clobbered.
        """
        with self._lock:
            current = self._sessions.get(session_id)
            if current is None:
                return None
            if expected is not None and current is not expected:
                return None  # a fresh session took the slot; leave it alone
            return self._sessions.pop(session_id, None)

    def session_ids(self) -> list[str]:
        with self._lock:
            return list(self._sessions.keys())

    def child_pids(self) -> list[int]:
        with self._lock:
            return [s.pid for s in self._sessions.values() if s.pid is not None]

    def stop_all(self) -> None:
        with self._lock:
            sessions = list(self._sessions.values())
            self._sessions.clear()
        for session in sessions:
            session.stop()


# --------------------------------------------------------------------------- #
# Watchdog                                                                     #
# --------------------------------------------------------------------------- #


class Watchdog:
    """Out-of-band reaper. Arms per-session deadlines; on expiry it kills the
    recorded PID and frees the slot — independent of any (possibly wedged)
    worker thread. It NEVER dereferences a COM proxy; it deals only in PIDs.
    """

    def __init__(self, terminate_process: Optional[TerminateProcess] = None,
                 on_reap: Optional[Callable[[str, int], None]] = None,
                 tick: float = 0.01,
                 clock: Optional[Callable[[], float]] = None):
        self._terminate = terminate_process or default_terminate_process
        self._on_reap = on_reap
        self._tick = tick
        self._clock = clock or time.monotonic
        self._deadlines: dict[str, tuple[float, int]] = {}
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = threading.Thread(
            target=self._run, name="watchdog", daemon=True
        )

    def start(self) -> None:
        self._thread.start()

    def arm(self, session_id: str, pid: Optional[int], deadline: float) -> None:
        """Arm a reap for ``session_id`` at monotonic time ``deadline``.

        ``pid`` may be ``None`` (unknown child PID); on expiry the kill is then
        skipped but the slot is still freed.
        """
        with self._lock:
            self._deadlines[session_id] = (deadline, pid)

    def disarm(self, session_id: str) -> None:
        with self._lock:
            self._deadlines.pop(session_id, None)

    def _run(self) -> None:
        while not self._stop.is_set():
            now = self._clock()
            fired: list[tuple[str, Optional[int]]] = []
            with self._lock:
                for sid, (deadline, pid) in list(self._deadlines.items()):
                    if now >= deadline:
                        fired.append((sid, pid))
                        del self._deadlines[sid]
            for sid, pid in fired:
                # SAFE-FAIL: never force-kill an unknown (None) pid or the
                # daemon's OWN pid. In that case skip the kill but still free
                # the slot (idle-exit reclaims the process). Guard the kill so a
                # ProcessLookupError (PID already exited) can't kill this loop.
                if pid is not None and pid != os.getpid():
                    try:
                        self._terminate(pid)
                    except Exception:
                        pass  # already-dead / unkillable PID: log-and-continue
                if self._on_reap is not None:
                    try:
                        self._on_reap(sid, pid)
                    except Exception:
                        pass
            self._stop.wait(self._tick)

    def stop(self, join_timeout: float = 2.0) -> None:
        self._stop.set()
        self._thread.join(timeout=join_timeout)


# --------------------------------------------------------------------------- #
# Singleton guard + lockfile                                                   #
# --------------------------------------------------------------------------- #


if sys.platform == "win32":
    import msvcrt

    def _lock_exclusive_nb(fh) -> None:
        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)

    def _unlock(fh) -> None:
        try:
            fh.seek(0)
            msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
else:
    import fcntl

    def _lock_exclusive_nb(fh) -> None:
        fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)

    def _unlock(fh) -> None:
        try:
            fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass


class SingletonGuard:
    """Exclusive lockfile lock. The first holder wins; a second
    :meth:`acquire` on the same path fails fast so that daemon can exit.
    """

    def __init__(self, lock_path: str):
        self._lock_path = lock_path
        self._fh = None

    def acquire(self) -> bool:
        try:
            if sys.platform == "win32":
                fh = open(self._lock_path, "a+")
            else:
                # O_NOFOLLOW defeats a symlink attack on the lock path; 0600
                # keeps it private. A symlinked path raises ELOOP -> refused.
                fd = os.open(
                    self._lock_path,
                    os.O_RDWR | os.O_CREAT | os.O_NOFOLLOW, 0o600,
                )
                fh = os.fdopen(fd, "a+")
        except OSError:
            return False
        try:
            _lock_exclusive_nb(fh)
        except OSError:
            fh.close()
            return False
        self._fh = fh
        return True

    def release(self) -> None:
        if self._fh is not None:
            _unlock(self._fh)
            self._fh.close()
            self._fh = None


def write_lockfile(path: str, port: int, token: str, pid: int,
                   child_pids: list[int]) -> None:
    """Atomically write the daemon discovery lockfile (user-only on POSIX).

    The token is a credential, so on POSIX the temp file is created 0600 via
    ``os.open`` BEFORE any bytes are written — it is never world-readable at
    any instant — then atomically renamed into place.
    """
    data = {"port": port, "token": token, "pid": pid,
            "child_pids": list(child_pids)}
    tmp = f"{path}.{os.getpid()}.tmp"
    if sys.platform == "win32":
        with open(tmp, "w") as fh:
            json.dump(data, fh)
    else:
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as fh:
            json.dump(data, fh)
    os.replace(tmp, path)
    if sys.platform != "win32":
        os.chmod(path, 0o600)


def read_lockfile(path: str) -> dict:
    with open(path) as fh:
        return json.load(fh)


def _ensure_private_dir(directory: str) -> None:
    """Create ``directory`` mode 0700 and verify it is ours and not group/other
    writable — raise on an insecure pre-existing dir (POSIX only)."""
    os.makedirs(directory, mode=0o700, exist_ok=True)
    try:
        os.chmod(directory, 0o700)
    except OSError:
        pass
    st = os.stat(directory)
    if st.st_uid != os.getuid():
        raise RuntimeError(
            f"refusing insecure lockfile dir (not owned by us): {directory}"
        )
    if st.st_mode & (stat.S_IWGRP | stat.S_IWOTH):
        raise RuntimeError(
            f"refusing insecure lockfile dir (group/other-writable): {directory}"
        )


def default_lockfile_path() -> str:
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
        directory = os.path.join(base, "origin-pro-mcp")
        os.makedirs(directory, exist_ok=True)
        return os.path.join(directory, "daemon.json")
    # POSIX: keep the token-bearing lockfile in a user-private 0700 dir.
    runtime = os.environ.get("XDG_RUNTIME_DIR")
    if runtime:
        directory = os.path.join(runtime, "origin-pro-mcp")
    else:
        directory = os.path.join(
            tempfile.gettempdir(), f"origin-pro-mcp-{os.getuid()}"
        )
    _ensure_private_dir(directory)
    return os.path.join(directory, "daemon.json")


def default_recovery_dir() -> str:
    """Where reap-time recovery sidecars are written (configurable)."""
    override = os.environ.get("ORIGIN_PRO_MCP_RECOVERY_DIR")
    if override:
        return override
    base = os.environ.get("LOCALAPPDATA")
    directory = (
        os.path.join(base, "origin-pro-mcp", "recovery") if base
        else os.path.join(tempfile.gettempdir(), "origin-pro-mcp", "recovery")
    )
    return directory


def recovery_path(recovery_dir: str, session_id: str,
                  project_path: Optional[str]) -> str:
    """Collision-safe recovery sidecar path for a reaped session.

    Scheme: ``<project_stem>.<session_id>.recover.opju`` — the session id
    namespaces the file UNCONDITIONALLY (even when a project is named) so two
    agents never collide. When no project is open, the stem is dropped and the
    session id alone names the file under ``recovery_dir``. If the chosen path
    already exists it is NEVER overwritten — an incrementing counter is suffixed
    (``.recover.1.opju``, ``.recover.2.opju`` …).
    """
    if project_path:
        stem = os.path.splitext(os.path.basename(project_path))[0]
    else:
        stem = ""
    base = f"{stem}.{session_id}.recover" if stem else f"{session_id}.recover"
    candidate = os.path.join(recovery_dir, f"{base}.opju")
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(recovery_dir, f"{base}.{counter}.opju")
        counter += 1
    return candidate


def default_is_alive(pid: int) -> bool:
    """Best-effort liveness check for a PID — out-of-band, no COM."""
    if not pid:
        return False
    if sys.platform == "win32":
        import ctypes  # local import: Windows-only path

        SYNCHRONIZE = 0x00100000
        handle = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True  # exists but not ours to signal
    except OSError:
        return False
    return True


def _default_registry() -> dict:
    from . import server  # noqa: F401 — importing registers every tool
    from .app import mcp

    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


# --------------------------------------------------------------------------- #
# Daemon                                                                       #
# --------------------------------------------------------------------------- #


class Daemon:
    """The singleton daemon: transport server + pool + watchdog.

    One TCP connection == one session (the session id arrives in the hello
    frame, or per-request). Request frames route to ``pool.acquire(...).submit``
    and responses go back over the same connection with the matching
    ``request_id``.
    """

    def __init__(self):
        self._guard: Optional[SingletonGuard] = None
        self._server: Optional[TcpServer] = None
        self._pool: Optional[Pool] = None
        self._watchdog: Optional[Watchdog] = None
        self._lockfile_path: Optional[str] = None
        self.token: Optional[str] = None
        self.port: Optional[int] = None
        self.host: Optional[str] = None
        self._accept_thread: Optional[threading.Thread] = None
        self._conns: list[Connection] = []
        self._conns_lock = threading.Lock()
        self._running = False

        self._conns_by_session: dict[str, set] = {}

        # -- lifecycle (3d) -------------------------------------------------- #
        self._clock: Callable[[], float] = time.monotonic
        self._reap_grace = 5.0
        self._heartbeat_reap_after = 30.0
        self._idle_exit_after = 600.0
        self._reconnect_grace = 0.0
        self._recovery_dir: Optional[str] = None
        self._project_path_getter = None
        self._terminate: TerminateProcess = default_terminate_process
        self._reap_lock = threading.Lock()
        # session_id -> the Session whose reap has COMMITTED (stage 1 started).
        self._reaping: dict[str, Session] = {}
        # session_id -> monotonic deadline at which a pending reap COMMITS. A
        # reconnect within the grace cancels it (see _mark_live).
        self._reap_pending: dict[str, float] = {}
        self._last_seen: dict[str, float] = {}
        self._seen_lock = threading.Lock()
        self._idle_since: Optional[float] = None
        self._lockfile_lock = threading.Lock()
        self._monitor_tick = 0.01
        self._monitor_stop: Optional[threading.Event] = None
        self._monitor_thread: Optional[threading.Thread] = None
        self._stop_lock = threading.Lock()
        self._stopped = False
        self._stopped_event = threading.Event()

    @property
    def pool(self) -> Optional[Pool]:
        return self._pool

    @property
    def watchdog(self) -> Optional[Watchdog]:
        return self._watchdog

    def start(self, origin_factory: OriginFactory, registry: Optional[dict] = None,
              max_size: int = POOL_MAX_DEFAULT, host: str = "127.0.0.1",
              port: int = 0, terminate_process: Optional[TerminateProcess] = None,
              lockfile_path: Optional[str] = None,
              get_pid: Optional[GetPid] = None,
              clock: Optional[Callable[[], float]] = None,
              reap_grace: float = 5.0,
              heartbeat_reap_after: float = 30.0,
              idle_exit_after: float = 600.0,
              reconnect_grace: float = DEFAULT_RECONNECT_GRACE,
              recovery_dir: Optional[str] = None,
              project_path_getter=None,
              is_alive: Optional[Callable[[int], bool]] = None,
              start_timeout: float = 10.0,
              monitor_tick: float = 0.01) -> bool:
        """Acquire the singleton, sweep any orphans left by a crashed prior
        daemon, start the server/pool/watchdog/monitor, write the lockfile, and
        begin accepting connections. Returns ``False`` (and starts nothing) if
        another daemon already holds the singleton.
        """
        if registry is None:
            registry = _default_registry()
        if lockfile_path is None:
            lockfile_path = default_lockfile_path()
        self._lockfile_path = lockfile_path
        self._clock = clock or time.monotonic
        self._reap_grace = reap_grace
        self._heartbeat_reap_after = heartbeat_reap_after
        self._idle_exit_after = idle_exit_after
        self._reconnect_grace = reconnect_grace
        self._recovery_dir = recovery_dir or default_recovery_dir()
        self._project_path_getter = project_path_getter
        self._monitor_tick = monitor_tick
        self._terminate = terminate_process or default_terminate_process

        self._guard = SingletonGuard(lockfile_path + ".lock")
        if not self._guard.acquire():
            self._guard = None
            return False

        # Startup sweep: reclaim orphans recorded by a crashed prior daemon
        # BEFORE we overwrite its lockfile. PID-authoritative, COM-free.
        self._startup_sweep(lockfile_path, is_alive or default_is_alive)

        self.token = secrets.token_hex(16)
        self._server = TcpServer(self.token, host=host, port=port)
        self.host, self.port = self._server.host, self._server.port
        self._pool = Pool(origin_factory, registry, max_size=max_size,
                          get_pid=get_pid, start_timeout=start_timeout)
        self._watchdog = Watchdog(terminate_process=self._terminate,
                                  on_reap=self._on_watchdog_reap,
                                  tick=monitor_tick, clock=self._clock)
        self._watchdog.start()
        write_lockfile(lockfile_path, self.port, self.token, os.getpid(),
                       self._pool.child_pids())

        self._running = True
        self._idle_since = self._clock()
        self._accept_thread = threading.Thread(
            target=self._accept_loop, name="daemon-accept", daemon=True
        )
        self._accept_thread.start()
        self._monitor_stop = threading.Event()
        self._monitor_thread = threading.Thread(
            target=self._monitor_loop, name="daemon-monitor", daemon=True
        )
        self._monitor_thread.start()
        return True

    @property
    def running(self) -> bool:
        return self._running

    def _startup_sweep(self, lockfile_path: str,
                       is_alive: Callable[[int], bool]) -> None:
        try:
            data = read_lockfile(lockfile_path)
        except (OSError, ValueError):
            return  # no prior lockfile / unreadable -> nothing to sweep
        for pid in data.get("child_pids", []) or []:
            try:
                pid = int(pid)
                # SAFE-FAIL: never sweep an unknown/zero pid or our own pid.
                if not pid or pid == os.getpid():
                    continue
                if is_alive(pid):
                    self._terminate(pid)
            except Exception:
                pass

    # -- reaping ------------------------------------------------------------- #

    def reap_session(self, session_id: str, reason: str = "") -> None:
        """COMMIT a two-stage reap of ``session_id`` (idempotent, non-blocking).

        Arms the watchdog (stage 2) at ``reap_grace`` from now, then enqueues the
        graceful save/close (stage 1) onto the session's OWN worker thread. This
        method never waits on that worker — if it wedges, the watchdog force-kills
        the recorded PID and frees the slot regardless.

        Once committed the session is marked ``reaping`` so a reconnect can no
        longer reuse it (``Pool.acquire`` then mints a fresh one). The cancelable
        grace window lives in the connection-close scheduler (see
        :meth:`_schedule_reap`), not here.
        """
        if not session_id or self._pool is None:
            return
        with self._reap_lock:
            self._reap_pending.pop(session_id, None)
            if session_id in self._reaping:
                return
            session = self._pool.get(session_id)
            if session is None:
                # Heartbeat-only id with no pool session: still drop liveness
                # state so _last_seen can't grow unbounded.
                with self._seen_lock:
                    self._last_seen.pop(session_id, None)
                return
            session.reaping = True
            self._reaping[session_id] = session
        # An unknown PID stays None here -> the watchdog skips the kill but
        # still frees the slot. NEVER substitute os.getpid().
        pid = session.pid
        if self._watchdog is not None:
            self._watchdog.arm(session_id, pid,
                               self._clock() + self._reap_grace)
        session.submit_reap(
            self._recovery_dir, self._project_path_getter,
            on_done=lambda sid=session_id: self._on_graceful_done(sid),
        )

    def _on_graceful_done(self, session_id: str) -> None:
        """Stage 1 finished in time: cancel the watchdog kill, free the slot."""
        if self._watchdog is not None:
            self._watchdog.disarm(session_id)
        self._finish_reap(session_id)

    def _on_watchdog_reap(self, session_id: str, _pid: int) -> None:
        """Stage 2 fired (worker wedged): the PID was killed; free the slot."""
        self._finish_reap(session_id)

    def _finish_reap(self, session_id: str) -> None:
        with self._reap_lock:
            session = self._reaping.pop(session_id, None)
        if self._pool is not None:
            # Only drop THIS reaping session — a reconnect may have already
            # replaced the slot with a fresh session under the same id.
            self._pool.discard(session_id, expected=session)
        with self._seen_lock:
            self._last_seen.pop(session_id, None)
        self._rewrite_lockfile()

    def _rewrite_lockfile(self) -> None:
        if not self._lockfile_path or self._stopped or self._pool is None:
            return
        try:
            with self._lockfile_lock:
                write_lockfile(self._lockfile_path, self.port, self.token,
                               os.getpid(), self._pool.child_pids())
        except OSError:
            pass

    # -- liveness tracking + monitor ----------------------------------------- #

    def _touch(self, session_id: str) -> None:
        if not session_id:
            return
        with self._seen_lock:
            self._last_seen[session_id] = self._clock()

    # -- connection refcount + cancelable reap scheduling -------------------- #

    def _mark_live(self, session_id: str, conn: Connection) -> None:
        """Register a live connection for a session and CANCEL any pending
        (not-yet-committed) reap — this is the reconnect-vs-reap fix: a new
        connection arriving within the grace reuses the session intact."""
        if not session_id:
            return
        with self._conns_lock:
            self._conns_by_session.setdefault(session_id, set()).add(conn)
        with self._reap_lock:
            self._reap_pending.pop(session_id, None)

    def _mark_dead(self, conn: Connection) -> None:
        """Deregister a closed connection; only when the LAST connection for a
        session_id closes do we SCHEDULE a (cancelable) reap."""
        session_id = conn.session_id
        if not session_id:
            return
        with self._conns_lock:
            conns = self._conns_by_session.get(session_id)
            if conns is not None:
                conns.discard(conn)
                empty = not conns
                if empty:
                    self._conns_by_session.pop(session_id, None)
            else:
                empty = True  # never registered -> treat as the last one
        if empty:
            self._schedule_reap(session_id)

    def _schedule_reap(self, session_id: str) -> None:
        """Arm a cancelable reap for ``session_id``. With a positive reconnect
        grace the commit is deferred (a reconnect can cancel it); with zero
        grace it commits immediately (preserving the original behavior)."""
        if self._reconnect_grace <= 0:
            self.reap_session(session_id, reason="connection-closed")
            return
        with self._reap_lock:
            if session_id in self._reaping:
                return  # already committed
            self._reap_pending[session_id] = (
                self._clock() + self._reconnect_grace
            )

    def _commit_due_reaps(self, now: float) -> None:
        """Commit pending reaps whose grace has elapsed and that still have no
        live connection (a reconnect would have cleared the pending entry)."""
        with self._reap_lock:
            due = [sid for sid, dl in self._reap_pending.items() if now >= dl]
        for sid in due:
            with self._conns_lock:
                live = bool(self._conns_by_session.get(sid))
            with self._reap_lock:
                still_pending = self._reap_pending.pop(sid, None) is not None
            if still_pending and not live:
                self.reap_session(sid, reason="connection-closed")

    def _monitor_loop(self) -> None:
        stop = self._monitor_stop
        while stop is not None and not stop.is_set():
            try:
                self._monitor_tick_once()
            except Exception:
                pass
            stop.wait(self._monitor_tick)

    def _monitor_tick_once(self) -> None:
        if self._pool is None:
            return
        now = self._clock()
        # Commit any connection-close reaps whose reconnect grace has elapsed.
        self._commit_due_reaps(now)
        # Heartbeat backstop: half-open detection only. The shim pings every
        # ~10s; a session silent past ``heartbeat_reap_after`` is reaped.
        for sid in self._pool.session_ids():
            if sid in self._reaping:
                continue
            with self._seen_lock:
                last = self._last_seen.get(sid)
            if last is not None and now - last > self._heartbeat_reap_after:
                self.reap_session(sid, reason="heartbeat-gap")
        # Idle self-exit: 0 active sessions for ``idle_exit_after`` -> shut down.
        if not self._pool.session_ids():
            if self._idle_since is None:
                self._idle_since = now
            elif now - self._idle_since > self._idle_exit_after:
                self._idle_shutdown()
        else:
            self._idle_since = None

    def _idle_shutdown(self) -> None:
        # Runs ON the monitor thread; stop() must not self-join that thread.
        self.stop()

    def _accept_loop(self) -> None:
        while self._running:
            try:
                conn = self._server.accept()
            except OSError:
                break  # server socket closed during shutdown
            if conn is None:
                continue  # bad token / failed handshake
            with self._conns_lock:
                self._conns.append(conn)
            threading.Thread(
                target=self._serve_connection, args=(conn,),
                name="daemon-conn", daemon=True
            ).start()

    def _serve_connection(self, conn: Connection) -> None:
        # Register the connection up front so a transient reconnect with the
        # same session_id cancels any pending reap before it commits.
        if conn.session_id:
            self._mark_live(conn.session_id, conn)
        try:
            while self._running:
                try:
                    frame = conn.recv_frame()
                except (OSError, FrameError):
                    break
                if frame is None:
                    break  # client closed
                if not self._handle_frame(conn, frame):
                    break
        finally:
            conn.close()
            # Connection-as-liveness: only when the LAST connection for this
            # session closes is a (cancelable) reap scheduled. Suppressed during
            # shutdown (the daemon closed the socket itself).
            if self._running and conn.session_id:
                self._mark_dead(conn)

    def _handle_frame(self, conn: Connection, frame: dict) -> bool:
        ftype = frame.get("type")
        session_id = frame.get("session_id") or conn.session_id
        if session_id and conn.session_id is None:
            conn.session_id = session_id
        if ftype in ("heartbeat", "hello"):
            self._mark_live(session_id, conn)  # cancel a pending reap
            self._touch(session_id)            # liveness backstop
            return True
        if ftype != "request":
            return True
        request_id = frame.get("request_id")
        name = frame.get("name")
        # MEDIUM 1: wrap the WHOLE request path so ANY failure (PoolFull,
        # start/factory/COM error, unknown tool, etc.) returns an actionable
        # response — the shim must never wait out its call_timeout on a hang.
        try:
            self._mark_live(session_id, conn)  # a request implies liveness
            session = self._pool.acquire(session_id)
            self._touch(session_id)
            session.submit(
                request_id, name, frame.get("kwargs") or {},
                lambda response: self._safe_send(conn, response),
            )
        except PoolFull as exc:
            self._safe_send(conn, {"type": "response", "request_id": request_id,
                                   "ok": False, "result": None,
                                   "error": str(exc)})
        except Exception as exc:  # noqa: BLE001 — surface, never hang the client
            self._safe_send(conn, {"type": "response", "request_id": request_id,
                                   "ok": False, "result": None,
                                   "error": f"{type(exc).__name__}: {exc}"})
        return True

    @staticmethod
    def _safe_send(conn: Connection, frame: dict) -> None:
        try:
            conn.send_frame(frame)
        except (OSError, FrameError):
            pass

    def stop(self) -> None:
        with self._stop_lock:
            if self._stopped:
                return
            self._stopped = True
        self._running = False
        if self._monitor_stop is not None:
            self._monitor_stop.set()
        if self._server is not None:
            self._server.close()
        with self._conns_lock:
            conns = list(self._conns)
            self._conns.clear()
        for conn in conns:
            conn.close()
        if self._pool is not None:
            self._pool.stop_all()
        if self._watchdog is not None:
            self._watchdog.stop()
        # Skip a self-join when idle-exit calls stop() from the monitor thread.
        mon = self._monitor_thread
        if mon is not None and mon is not threading.current_thread():
            mon.join(timeout=2.0)
        if self._lockfile_path and os.path.exists(self._lockfile_path):
            try:
                os.remove(self._lockfile_path)
            except OSError:
                pass
        if self._guard is not None:
            self._guard.release()
            self._guard = None
        self._stopped_event.set()


# --------------------------------------------------------------------------- #
# Origin factory resolution + daemon entry point                              #
# --------------------------------------------------------------------------- #


class _EmptyPages:
    """An empty COM page collection (no open books/graphs/matrices)."""

    Count = 0

    def Item(self, _i):  # pragma: no cover - never reached (Count == 0)
        raise IndexError(_i)


class _InPackageFakeOrigin:
    """Minimal in-package COM double for the WSL auto-spawn lane.

    Selected ONLY when ``ORIGIN_PRO_MCP_FAKE_ORIGIN=1`` so the shim's
    auto-spawn path (spawn a detached daemon, connect, forward) is exercisable
    on a machine without COM. It is never used in production. It implements just
    enough of the Origin COM surface for ``run_labtalk`` / list tools to run.
    """

    def __init__(self):
        self.executed: list[str] = []
        self._lt_vars: dict = {}

    def Execute(self, script):
        self.executed.append(script)
        return True

    def LTVar(self, name):
        return self._lt_vars.get(name, 0.0)

    def LTStr(self, _name):
        return ""

    @property
    def WorksheetPages(self):
        return _EmptyPages()

    @property
    def GraphPages(self):
        return _EmptyPages()

    @property
    def MatrixPages(self):
        return _EmptyPages()

    def Save(self, _path):
        return True

    def Load(self, _path):
        return True


# Origin's COM Application object exposes no usable window handle on the tested
# build (Origin 2020: ``instance.Hwnd`` does not exist), so the spawned process
# id is captured by diffing the ``Origin64.exe`` process list around the
# ``DispatchEx`` launch. The launch + snapshot is serialized by ``_LAUNCH_LOCK``
# so two concurrent session launches can't mis-attribute each other's new PID,
# and the captured PID is stashed on a thread-local that ``_real_origin_get_pid``
# reads back on the SAME worker thread (factory + get_pid run sequentially there).
# Candidate Origin executable image names, across versions/bitness. Override
# with ORIGIN_PRO_MCP_ORIGIN_IMAGE (comma-separated) for a non-standard install.
# If none match, PID capture returns nothing and the watchdog's safe-fail guard
# applies — graceful close + idle-exit still reclaim sessions, only the hard
# force-kill backstop is unavailable.
_DEFAULT_ORIGIN_IMAGES = ("Origin64.exe", "Origin.exe", "OriginPro.exe")
_LAUNCH_LOCK = threading.Lock()
_real_pid_tls = threading.local()


def _origin_image_names() -> tuple:
    override = os.environ.get("ORIGIN_PRO_MCP_ORIGIN_IMAGE")
    if override:
        return tuple(n.strip() for n in override.split(",") if n.strip())
    return _DEFAULT_ORIGIN_IMAGES


def _origin_process_pids() -> set:
    """PIDs of all running Origin processes (Windows, best-effort).

    ONE ``tasklist`` call (all processes, filtered in Python) — matching any
    candidate image name so it works across Origin versions and 32/64-bit
    installs. The call uses ``CREATE_NO_WINDOW`` so it never flashes a console
    window: the daemon is a windowless background process, so a console child
    would otherwise pop a window every call (and this runs in a tight poll loop).
    """
    import subprocess

    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    out = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, **kwargs,
    ).stdout
    names = {n.lower() for n in _origin_image_names()}
    pids = set()
    for line in out.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].lower() in names and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _origin_visible() -> int:
    """Whether a launched Origin shows its window. Env: ORIGIN_PRO_MCP_VISIBLE.

    Default 1 (visible — watch the agent work). Set to 0/false/hidden/invisible
    for headless/batch runs (e.g. many concurrent sessions with no windows).
    """
    val = os.environ.get("ORIGIN_PRO_MCP_VISIBLE")
    if val is None:
        return 1
    return 0 if val.strip().lower() in ("0", "false", "no", "off", "hidden", "invisible") else 1


def _real_origin_factory():
    """Default factory: a fresh, isolated ``Origin.exe`` per session (Windows).

    Uses ``DispatchEx`` (out-of-process) so each session gets its own process —
    model B1. Captures the spawned ``Origin64.exe`` PID by snapshot-diff (under
    ``_LAUNCH_LOCK`` to keep attribution race-free) and stashes it for
    ``_real_origin_get_pid``. On non-Windows this raises when invoked.
    """
    import win32com.client  # Windows-only; imported lazily on the worker thread

    pid: Optional[int] = None
    with _LAUNCH_LOCK:
        before = _origin_process_pids()
        instance = win32com.client.DispatchEx("Origin.Application")
        # DispatchEx returns once the server is connectable, but the process may
        # take a moment to appear in the task list — poll briefly.
        deadline = time.monotonic() + 10.0
        while time.monotonic() < deadline:
            new = _origin_process_pids() - before
            if new:
                pid = sorted(new)[0]
                break
            time.sleep(0.2)
    # Show (or hide, per ORIGIN_PRO_MCP_VISIBLE) the instance's window —
    # DispatchEx instances start hidden, unlike the old ApplicationSI path.
    # Best-effort: a headless/odd build may reject it.
    try:
        instance.Visible = _origin_visible()
    except Exception:
        pass
    _real_pid_tls.pid = pid
    return instance


def _real_origin_get_pid(instance: object) -> Optional[int]:
    """Return the ``Origin64.exe`` PID captured by ``_real_origin_factory``.

    The factory stashes the snapshot-diff PID on a thread-local that this reads
    back on the same worker thread. Returns ``None`` if capture failed, so the
    watchdog's safe-fail guard applies — an unknown PID means "do not
    force-kill", never "kill the daemon itself".
    """
    return getattr(_real_pid_tls, "pid", None)


def resolve_origin_factory() -> OriginFactory:
    """Pick the daemon's Origin factory from the environment.

    * ``ORIGIN_PRO_MCP_ORIGIN_FACTORY`` — a dotted path to a zero-arg callable
      (escape hatch for tests / custom integrations), takes precedence.
    * ``ORIGIN_PRO_MCP_FAKE_ORIGIN=1`` — the in-package fake (WSL test seam).
    * otherwise — the real ``DispatchEx`` factory.
    """
    dotted = os.environ.get("ORIGIN_PRO_MCP_ORIGIN_FACTORY")
    if dotted:
        import importlib

        module_name, _, attr = dotted.rpartition(".")
        module = importlib.import_module(module_name)
        return getattr(module, attr)
    if os.environ.get("ORIGIN_PRO_MCP_FAKE_ORIGIN") == "1":
        return _InPackageFakeOrigin
    return _real_origin_factory


def resolve_get_pid(factory: OriginFactory) -> Optional[GetPid]:
    """Pick the child-PID resolver for ``factory``.

    Only the real ``DispatchEx`` factory spawns a separate ``Origin.exe`` whose
    PID we can resolve (and force-kill). For any other factory (fakes / custom
    integrations) return ``None`` so the pool falls back to the safe default
    (``_default_get_pid`` -> ``None`` -> "do not force-kill").
    """
    if factory is _real_origin_factory:
        return _real_origin_get_pid
    return None


def main(argv: Optional[list] = None) -> int:
    """Daemon entry point (``python -m origin_pro_mcp.daemon``).

    Acquires the singleton, starts serving, and blocks until terminated. A
    losing daemon (another already holds the singleton) exits immediately with
    status 0 so the shim's auto-spawn race resolves cleanly. The lockfile path
    may be overridden via ``ORIGIN_PRO_MCP_LOCKFILE`` (used by the shim).
    """
    lockfile_path = os.environ.get("ORIGIN_PRO_MCP_LOCKFILE") or None
    _grace_env = os.environ.get("ORIGIN_PRO_MCP_RECONNECT_GRACE")
    reconnect_grace = float(_grace_env) if _grace_env is not None else DEFAULT_RECONNECT_GRACE
    # A cold Origin launch via DispatchEx takes ~8-15s (longer on first launch /
    # slow disks), so the production session start timeout must be generous —
    # well above the test default of 10s. Overridable via env.
    _start_env = os.environ.get("ORIGIN_PRO_MCP_START_TIMEOUT")
    start_timeout = float(_start_env) if _start_env is not None else 45.0
    factory = resolve_origin_factory()
    # Wire the real child-PID resolver so production force-kills the spawned
    # Origin.exe (never the daemon's own pid); falls back to the safe default
    # when the PID can't be resolved.
    get_pid = resolve_get_pid(factory)
    daemon = Daemon()
    if not daemon.start(origin_factory=factory, get_pid=get_pid,
                        lockfile_path=lockfile_path,
                        reconnect_grace=reconnect_grace,
                        start_timeout=start_timeout):
        return 0  # another daemon owns the singleton; the loser exits cleanly

    stop = threading.Event()

    def _on_signal(_signum, _frame):
        stop.set()

    try:
        import signal

        signal.signal(signal.SIGTERM, _on_signal)
    except (ValueError, OSError, AttributeError):
        pass  # not on the main thread / platform without SIGTERM

    try:
        while not stop.wait(0.5):
            if daemon._stopped_event.is_set():
                break  # idle self-exit shut the daemon down
    except KeyboardInterrupt:
        pass
    finally:
        daemon.stop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
