"""Rigorous, COM-free tests for the loopback-TCP frame transport.

Everything here runs against real sockets on 127.0.0.1 with OS-assigned ports.
The headline test is the SEND-LOCK + correlation check: many threads hammer a
single connection concurrently (plus a heartbeat thread) and every response
must still match its request_id with no corrupted/interleaved frames.
"""
from __future__ import annotations

import socket
import struct
import threading
import time

import pytest

from origin_pro_mcp import transport
from origin_pro_mcp.transport import (
    Connection,
    FrameError,
    TcpServer,
    Transport,
    connect,
)

TOKEN = "test-token-abc123"


@pytest.fixture
def server():
    """A bound TcpServer on an ephemeral 127.0.0.1 port, closed on teardown."""
    srv = TcpServer(TOKEN, host="127.0.0.1", port=0)
    try:
        yield srv
    finally:
        srv.close()


def _socketpair_connections():
    """Two Connections wrapping the ends of a connected socket pair."""
    a, b = socket.socketpair()
    return Connection(a), Connection(b)


# --------------------------------------------------------------------------- #
# Framing                                                                      #
# --------------------------------------------------------------------------- #


def test_protocol_runtime_checkable():
    conn, other = _socketpair_connections()
    try:
        assert isinstance(conn, Transport)
    finally:
        conn.close()
        other.close()


def test_length_prefix_round_trip():
    a, b = _socketpair_connections()
    try:
        frame = {"type": "request", "request_id": "1", "name": "noop",
                 "kwargs": {"x": 5, "s": "héllo with spaces"}}
        a.send_frame(frame)
        assert b.recv_frame() == frame
    finally:
        a.close()
        b.close()


def test_clean_eof_returns_none():
    a, b = _socketpair_connections()
    try:
        a.close()  # clean shutdown at a frame boundary
        assert b.recv_frame() is None
    finally:
        b.close()


def test_truncated_frame_raises():
    a, b = _socketpair_connections()
    try:
        # Claim 100 bytes of body, then send only a few and hang up.
        a._sock.sendall(struct.pack(">I", 100) + b"partial")
        a.close()
        with pytest.raises(FrameError):
            b.recv_frame()
    finally:
        b.close()


def test_garbage_body_raises():
    a, b = _socketpair_connections()
    try:
        body = b"this is not json"
        a._sock.sendall(struct.pack(">I", len(body)) + body)
        with pytest.raises(FrameError):
            b.recv_frame()
    finally:
        a.close()
        b.close()


def test_send_frame_requires_type():
    a, b = _socketpair_connections()
    try:
        with pytest.raises(FrameError):
            a.send_frame({"request_id": "1"})  # no 'type'
    finally:
        a.close()
        b.close()


# --------------------------------------------------------------------------- #
# SECURITY 1 — oversize frame length is rejected BEFORE allocating             #
# --------------------------------------------------------------------------- #


def test_oversize_frame_rejected_before_alloc():
    """A header declaring more than MAX_FRAME bytes must raise FrameError
    immediately — without reading/allocating the (unbounded) body."""
    from origin_pro_mcp.transport import MAX_FRAME

    a, b = _socketpair_connections()
    try:
        b.settimeout(2.0)  # so an UNFIXED build (which blocks on the body) fails
        # Send ONLY the 4-byte length header claiming an oversize body.
        a._sock.sendall(struct.pack(">I", MAX_FRAME + 1))
        with pytest.raises(FrameError):
            b.recv_frame()
    finally:
        a.close()
        b.close()


def test_max_frame_is_eight_mib():
    from origin_pro_mcp.transport import MAX_FRAME

    assert MAX_FRAME == 8 * 1024 * 1024


# --------------------------------------------------------------------------- #
# SECURITY 2 — handshake timeout: a silent peer can't wedge the accept loop    #
# --------------------------------------------------------------------------- #


def test_silent_peer_does_not_wedge_accept():
    """A peer that connects and sends nothing must NOT block accept() beyond the
    handshake timeout, so a subsequent legitimate connect still gets served."""
    srv = TcpServer(TOKEN, host="127.0.0.1", port=0, handshake_timeout=0.3)
    try:
        silent = socket.create_connection((srv.host, srv.port))
        t0 = time.monotonic()
        result = srv.accept()  # silent peer: must time out -> None, not wedge
        elapsed = time.monotonic() - t0
        assert result is None
        assert elapsed < 3.0, f"accept wedged for {elapsed:.1f}s on a silent peer"
        silent.close()

        # A legitimate client still connects fine afterwards.
        holder: list = []
        th = threading.Thread(target=lambda: holder.append(srv.accept()))
        th.start()
        good = connect(srv.host, srv.port, TOKEN, session_id="ok")
        th.join(timeout=3.0)
        assert holder and holder[0] is not None
        assert holder[0].session_id == "ok"
        good.close()
    finally:
        srv.close()


# --------------------------------------------------------------------------- #
# LOW — missing/None token is rejected (constant-time compare path)            #
# --------------------------------------------------------------------------- #


def test_missing_token_rejected(server):
    accepted: list = []
    t = threading.Thread(target=lambda: accepted.append(server.accept()))
    t.start()

    raw = socket.create_connection((server.host, server.port))
    conn = Connection(raw)
    conn.settimeout(3.0)
    conn.send_frame({"type": "hello", "session_id": "x"})  # NO token field
    assert conn.recv_frame() is None  # server rejected + closed
    conn.close()

    t.join(timeout=3.0)
    assert accepted == [None]


# --------------------------------------------------------------------------- #
# Token handshake                                                              #
# --------------------------------------------------------------------------- #


def test_wrong_token_rejected_and_closed(server):
    accepted: list = []
    t = threading.Thread(target=lambda: accepted.append(server.accept()))
    t.start()

    conn = connect(server.host, server.port, "WRONG-TOKEN")
    conn.settimeout(3.0)
    # Server must reject (close) the connection -> client sees a clean EOF.
    assert conn.recv_frame() is None
    conn.close()

    t.join(timeout=3.0)
    assert not t.is_alive()
    assert accepted == [None]  # accept() returned None for the bad token


def test_correct_token_accepted_and_echoes(server):
    holder: list = []
    ready = threading.Event()

    def serve():
        conn = server.accept()
        holder.append(conn)
        ready.set()
        if conn is None:
            return
        req = conn.recv_frame()
        conn.send_frame({"type": "response", "request_id": req["request_id"],
                         "ok": True, "result": req["name"], "error": None})

    t = threading.Thread(target=serve)
    t.start()

    conn = connect(server.host, server.port, TOKEN, session_id="s-1")
    conn.settimeout(3.0)
    conn.send_frame({"type": "request", "request_id": "r1", "name": "ping",
                     "kwargs": {}})
    resp = conn.recv_frame()

    t.join(timeout=3.0)
    assert holder[0] is not None
    assert holder[0].session_id == "s-1"  # hello carried the session id
    assert resp == {"type": "response", "request_id": "r1", "ok": True,
                    "result": "ping", "error": None}
    conn.close()


# --------------------------------------------------------------------------- #
# SEND-LOCK + correlation under concurrency                                    #
# --------------------------------------------------------------------------- #


def test_concurrent_sends_and_heartbeats_never_interleave(server):
    """100+ concurrent request senders + a heartbeat flood on ONE connection.

    The server reads frames with a single reader and echoes each request_id
    back. If the per-connection SEND-LOCK failed, two senders' bytes would
    interleave, the length prefix would desync, and the server's recv_frame
    would raise (recorded in ``server_errors``) and/or responses would be
    missing/mismatched. We assert: zero server-side frame errors, every
    request_id answered exactly once, and every result echoes its own id.
    """
    n_requests = 250
    server_errors: list = []
    heartbeats_seen = [0]
    server_conn: list = []

    def serve():
        conn = server.accept()
        server_conn.append(conn)
        seen = 0
        try:
            while seen < n_requests:
                frame = conn.recv_frame()
                if frame is None:
                    break
                if frame["type"] == "heartbeat":
                    heartbeats_seen[0] += 1
                    continue
                if frame["type"] == "request":
                    seen += 1
                    conn.send_frame({
                        "type": "response",
                        "request_id": frame["request_id"],
                        "ok": True,
                        "result": frame["request_id"],  # echo id into result
                        "error": None,
                    })
        except FrameError as exc:  # interleaved bytes -> desync -> raise here
            server_errors.append(str(exc))

    server_thread = threading.Thread(target=serve)
    server_thread.start()

    client = connect(server.host, server.port, TOKEN)
    client.settimeout(10.0)

    # Single reader thread on the client correlates responses by request_id.
    responses: dict[str, str] = {}
    reader_errors: list = []

    def reader():
        try:
            while len(responses) < n_requests:
                frame = client.recv_frame()
                if frame is None:
                    break
                if frame["type"] == "response":
                    responses[frame["request_id"]] = frame["result"]
        except FrameError as exc:
            reader_errors.append(str(exc))

    reader_thread = threading.Thread(target=reader)
    reader_thread.start()

    # Heartbeat flood: keep writing pings on the same socket the whole time.
    stop_heartbeat = threading.Event()

    def heartbeat():
        while not stop_heartbeat.is_set():
            client.send_frame({"type": "heartbeat", "ts": time.time()})
            time.sleep(0.0005)

    hb_thread = threading.Thread(target=heartbeat)
    hb_thread.start()

    # Fan out: one thread per request, all sharing the single connection.
    def send_one(i: int):
        client.send_frame({"type": "request", "request_id": f"req-{i}",
                           "name": "noop", "kwargs": {"i": i}})

    senders = [threading.Thread(target=send_one, args=(i,))
               for i in range(n_requests)]
    for s in senders:
        s.start()
    for s in senders:
        s.join(timeout=10.0)

    reader_thread.join(timeout=10.0)
    stop_heartbeat.set()
    hb_thread.join(timeout=3.0)
    server_thread.join(timeout=3.0)

    client.close()
    if server_conn and server_conn[0] is not None:
        server_conn[0].close()

    assert server_errors == [], f"frame corruption on server: {server_errors}"
    assert reader_errors == [], f"frame corruption on client: {reader_errors}"
    assert len(responses) == n_requests
    # Correlation: every response carried its OWN request_id, none crossed.
    for i in range(n_requests):
        rid = f"req-{i}"
        assert responses[rid] == rid
