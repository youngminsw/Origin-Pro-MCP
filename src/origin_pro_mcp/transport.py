"""Frame transport for the shim<->daemon boundary.

The :class:`Transport` protocol captures the seam — ``send_frame`` /
``recv_frame`` / ``close`` — so a Windows named-pipe implementation (Phase 3e)
can be dropped in later without touching the daemon, pool, or lifecycle code.

The v1 implementation is loopback TCP (127.0.0.1, ephemeral port) with
length-prefixed JSON frames, a per-connection send lock, and a token
handshake. Frames are ``4-byte big-endian length`` + ``UTF-8 JSON body``; the
declared length is bounded by :data:`MAX_FRAME` (8 MiB) and rejected before any
allocation, so a hostile/garbled peer cannot trigger a pre-auth memory DoS.
Every frame always carries a ``type`` field:

* ``{"type": "hello", "token": <str>, "session_id": <str|None>}``
* ``{"type": "request", "request_id": <str>, "name": <tool>, "kwargs": {...}}``
* ``{"type": "response", "request_id": <str>, "ok": <bool>,
     "result": <str|None>, "error": <str|None>}``
* ``{"type": "heartbeat", "ts": <num>}``  (never carries a ``request_id``)
"""
from __future__ import annotations

import json
import secrets
import socket
import struct
import threading
from typing import Optional, Protocol, runtime_checkable

# 4-byte big-endian unsigned length prefix. The wire format could carry up to
# ~4 GiB, but recv_frame caps the declared length at MAX_FRAME (see below).
_LEN = struct.Struct(">I")

# Hard cap on a single frame's body. A larger declared length is rejected
# BEFORE allocating, closing a pre-auth memory-DoS vector.
MAX_FRAME = 8 * 1024 * 1024

# How long accept() waits for the hello/token handshake before giving up, so a
# silent peer can never wedge the single-threaded accept loop pre-auth.
HANDSHAKE_TIMEOUT = 5.0


@runtime_checkable
class Transport(Protocol):
    """The minimal seam every transport (TCP, named pipe, ...) must implement."""

    def send_frame(self, frame: dict) -> None:
        """Serialize and send one frame atomically."""

    def recv_frame(self) -> Optional[dict]:
        """Read one frame, or ``None`` on a clean EOF at a frame boundary."""

    def close(self) -> None:
        """Close the underlying connection."""


class FrameError(Exception):
    """A frame was truncated, malformed, or missing its required ``type``."""


def _require_type(frame: dict) -> dict:
    if not isinstance(frame, dict) or "type" not in frame:
        raise FrameError("frame must be a dict with a 'type' field")
    return frame


class Connection:
    """A connected socket speaking length-prefixed JSON frames.

    ``send_frame`` holds a per-connection lock (SEND-LOCK) so concurrent senders
    — e.g. FastMCP forwarder threads plus a heartbeat timer — never interleave
    their bytes on the wire. ``recv_frame`` is meant to be driven by a single
    reader thread; it reads exactly the prefixed number of bytes.
    """

    def __init__(self, sock: socket.socket, session_id: Optional[str] = None):
        self._sock = sock
        self._send_lock = threading.Lock()
        self.session_id = session_id

    def settimeout(self, timeout: Optional[float]) -> None:
        self._sock.settimeout(timeout)

    def send_frame(self, frame: dict) -> None:
        _require_type(frame)
        body = json.dumps(frame).encode("utf-8")
        header = _LEN.pack(len(body))
        with self._send_lock:
            self._sock.sendall(header + body)

    def recv_frame(self) -> Optional[dict]:
        header = self._recv_exact(_LEN.size, allow_eof=True)
        if header is None:
            return None  # clean EOF at a frame boundary
        (length,) = _LEN.unpack(header)
        if length > MAX_FRAME:
            # Reject BEFORE allocating/reading the body (pre-auth DoS guard).
            raise FrameError(
                f"frame length {length} exceeds MAX_FRAME ({MAX_FRAME})"
            )
        body = self._recv_exact(length, allow_eof=False)
        try:
            frame = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise FrameError(f"garbage frame body: {exc}") from exc
        return _require_type(frame)

    def _recv_exact(self, n: int, allow_eof: bool) -> Optional[bytes]:
        chunks: list[bytes] = []
        remaining = n
        while remaining > 0:
            chunk = self._sock.recv(remaining)
            if not chunk:
                if remaining == n and allow_eof:
                    return None  # nothing read yet -> clean EOF
                raise FrameError(
                    f"truncated frame: connection closed with {remaining} "
                    f"of {n} bytes unread"
                )
            chunks.append(chunk)
            remaining -= len(chunk)
        return b"".join(chunks)

    def close(self) -> None:
        try:
            self._sock.close()
        except OSError:
            pass


def _valid_hello(frame: Optional[dict], token: str) -> bool:
    if not isinstance(frame, dict) or frame.get("type") != "hello":
        return False
    # Constant-time compare so a peer can't probe the token byte-by-byte.
    return secrets.compare_digest(str(frame.get("token") or ""), token)


class TcpServer:
    """A loopback-TCP listener that token-authenticates each connection.

    Binds ``127.0.0.1`` only (never ``0.0.0.0``). Pass ``port=0`` to let the OS
    assign an ephemeral port, then read :attr:`port`.
    """

    def __init__(self, token: str, host: str = "127.0.0.1", port: int = 0,
                 backlog: int = 8, handshake_timeout: float = HANDSHAKE_TIMEOUT):
        self._token = token
        self._handshake_timeout = handshake_timeout
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._sock.bind((host, port))
        self._sock.listen(backlog)
        self.host, self.port = self._sock.getsockname()

    def accept(self) -> Optional[Connection]:
        """Accept and authenticate one connection.

        Returns an authenticated :class:`Connection` (with ``session_id`` taken
        from the hello frame), or ``None`` when the token was wrong/the
        handshake failed/timed out — in which case the connection has been
        closed. A handshake timeout bounds the hello/token read so a silent peer
        cannot wedge the single-threaded accept loop; the timeout is cleared
        (``settimeout(None)``) before the connection is returned for serving.
        Raises ``OSError`` when the listening socket is closed.
        """
        raw, _ = self._sock.accept()
        conn = Connection(raw)
        conn.settimeout(self._handshake_timeout)
        try:
            hello = conn.recv_frame()
        except (OSError, FrameError):
            conn.close()
            return None
        if not _valid_hello(hello, self._token):
            conn.close()
            return None
        conn.settimeout(None)  # restore blocking mode for the serving phase
        conn.session_id = hello.get("session_id")
        return conn

    def close(self) -> None:
        try:
            self._sock.close()
        except OSError:
            pass


def connect(host: str, port: int, token: str,
            session_id: Optional[str] = None) -> Connection:
    """Open a client connection and perform the hello/token handshake."""
    raw = socket.create_connection((host, port))
    conn = Connection(raw, session_id=session_id)
    conn.send_frame({"type": "hello", "token": token, "session_id": session_id})
    return conn
