"""Attach mode: at most ONE session attaches to the user's shared Origin
(attach_factory / Origin.ApplicationSI); further attach requests fall back to an
isolated instance. COM-free via tagged fake factories."""
from fakes import FakeOrigin
from origin_pro_mcp.daemon import Pool


def _registry():
    from origin_pro_mcp import server  # noqa: F401 — registers tools
    from origin_pro_mcp.app import mcp
    return {n: t.fn for n, t in mcp._tool_manager._tools.items()}


def _factories():
    def iso():
        o = FakeOrigin()
        o.tag = "iso"
        return o

    def att():
        o = FakeOrigin()
        o.tag = "att"
        return o
    return iso, att


def test_attach_uses_shared_factory_only_once():
    iso, att = _factories()
    pool = Pool(iso, _registry(), max_size=5, attach_factory=att)
    try:
        s1 = pool.acquire("A", attach=True)
        assert s1.instance.tag == "att"        # attached to the shared instance
        # second attach request -> guard -> falls back to an isolated instance
        s2 = pool.acquire("B", attach=True)
        assert s2.instance.tag == "iso"
        # a non-attach session is always isolated
        s3 = pool.acquire("C", attach=False)
        assert s3.instance.tag == "iso"
    finally:
        pool.stop_all()


def test_non_attach_never_uses_shared_factory():
    iso, att = _factories()
    pool = Pool(iso, _registry(), max_size=3, attach_factory=att)
    try:
        assert pool.acquire("A", attach=False).instance.tag == "iso"
    finally:
        pool.stop_all()


def test_attach_slot_frees_on_discard():
    iso, att = _factories()
    pool = Pool(iso, _registry(), max_size=3, attach_factory=att)
    try:
        assert pool.acquire("A", attach=True).instance.tag == "att"
        pool.discard("A")                       # frees the single attach slot
        # a new session can now attach to the shared instance again
        assert pool.acquire("B", attach=True).instance.tag == "att"
    finally:
        pool.stop_all()


def test_attach_disabled_when_no_attach_factory():
    iso, _ = _factories()
    pool = Pool(iso, _registry(), max_size=3, attach_factory=None)
    try:
        # attach requested but unavailable -> isolated
        assert pool.acquire("A", attach=True).instance.tag == "iso"
    finally:
        pool.stop_all()


def test_shim_attach_flag_reads_env(monkeypatch):
    from origin_pro_mcp.shim import ShimClient
    monkeypatch.delenv("ORIGIN_PRO_MCP_ATTACH", raising=False)
    assert ShimClient(allow_spawn=False)._attach is False
    monkeypatch.setenv("ORIGIN_PRO_MCP_ATTACH", "1")
    assert ShimClient(allow_spawn=False)._attach is True
    monkeypatch.setenv("ORIGIN_PRO_MCP_ATTACH", "off")
    assert ShimClient(allow_spawn=False)._attach is False
