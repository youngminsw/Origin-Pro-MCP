"""Tests for skills-as-a-first-class-MCP-feature.

Skills are bundled markdown shipped inside the package. They are exposed via two
COM-free tools — ``list_skills`` / ``get_skill`` — registered on BOTH the
in-process server and the daemon-backed shim (so a connecting agent discovers
them regardless of transport). All assertions here are WSL-safe (no Origin/COM).
"""
from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from origin_pro_mcp import skills


# --------------------------------------------------------------------------- #
# discover_skills                                                             #
# --------------------------------------------------------------------------- #


def test_discover_finds_publication_figure():
    found = skills.discover_skills()
    by_name = {s["name"]: s for s in found}
    assert "publication-figure" in by_name
    skill = by_name["publication-figure"]
    assert skill["title"].strip()           # non-empty H1 title
    assert skill["summary"].strip()         # non-empty one-line summary
    assert skill["content"].strip()         # full markdown body
    # Title comes from the first H1.
    assert skill["title"] == "Publication Figure"
    # Summary is derived from the "When To Use" section.
    assert "manuscript figure" in skill["summary"]


def test_discover_is_cached():
    assert skills.discover_skills() is skills.discover_skills()


# --------------------------------------------------------------------------- #
# register_skills — tools                                                     #
# --------------------------------------------------------------------------- #


def _registry(mcp: FastMCP) -> dict:
    return {name: tool.fn for name, tool in mcp._tool_manager._tools.items()}


def test_register_adds_both_tools():
    mcp = FastMCP("t")
    skills.register_skills(mcp)
    reg = _registry(mcp)
    assert "list_skills" in reg
    assert "get_skill" in reg


def test_list_skills_mentions_publication_figure():
    mcp = FastMCP("t")
    skills.register_skills(mcp)
    out = _registry(mcp)["list_skills"]()
    assert "publication-figure" in out
    assert "Publication Figure" in out  # the title


def test_get_skill_returns_full_markdown():
    mcp = FastMCP("t")
    skills.register_skills(mcp)
    out = _registry(mcp)["get_skill"]("publication-figure")
    assert "# Publication Figure" in out          # the H1
    assert "## When To Use" in out                # a known section
    # The full content, not just the summary.
    assert out == skills.discover_skills()[0]["content"] or "# Publication Figure" in out


def test_get_skill_unknown_is_actionable():
    mcp = FastMCP("t")
    skills.register_skills(mcp)
    out = _registry(mcp)["get_skill"]("nope")
    assert "nope" in out
    assert "publication-figure" in out  # lists valid names


# --------------------------------------------------------------------------- #
# Registration on both entrypoints                                            #
# --------------------------------------------------------------------------- #


def test_in_process_server_exposes_skill_tools():
    from origin_pro_mcp import server  # noqa: F401 — registers all tools
    from origin_pro_mcp.app import mcp

    names = set(mcp._tool_manager._tools)
    assert "list_skills" in names
    assert "get_skill" in names


def test_shim_exposes_skill_tools():
    from origin_pro_mcp import shim

    client = shim.ShimClient(heartbeat_interval=0)
    shim_server = shim.build_shim_server(client)
    names = set(shim_server._tool_manager._tools)
    assert "list_skills" in names
    assert "get_skill" in names
