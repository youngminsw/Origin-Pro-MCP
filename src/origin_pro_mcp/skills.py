"""Skills as a first-class MCP feature.

Skills are plain markdown documents bundled inside the package
(``origin_pro_mcp/skills/*.md``). They ship in the wheel, so they are loadable
whether the server runs from a source checkout or a ``pip``/``uvx`` install.

Two COM-free tools expose them so any connecting agent discovers them with zero
manual setup:

* ``list_skills`` — a compact catalogue (name + title + summary) so an agent
  autonomously learns which skills exist and when to use each;
* ``get_skill`` — the full markdown of a named skill.

Both tools read bundled markdown only (no Origin/COM), so they are registered
*locally* on every server — the in-process server AND the daemon-backed shim —
and must never be forwarded to the daemon.
"""
from __future__ import annotations

import re
from importlib import resources
from typing import List, Optional

# Names of the tools this module registers. The shim uses this to register them
# locally instead of building daemon-forwarders for them.
SKILL_TOOL_NAMES = ("list_skills", "get_skill")

# Server-level instructions sent to the client at initialize: which bundled
# skill to read for which task.
SERVER_INSTRUCTIONS = (
    "This server controls OriginLab Origin Pro through COM automation and "
    "bundles skills that record how Origin behaves when scripted: verified "
    "recipes, silent no-ops, and unit traps. Before building or styling a "
    "figure for a paper, thesis, or talk, read the `publication-figure` skill "
    "with `get_skill`. Before writing a non-trivial `run_labtalk` script, or "
    "when a styling call reports success but the exported image shows no "
    "change, read `labtalk-gotchas`. `list_skills` shows every bundled skill "
    "and when to use it. Prefer the typed tools; use `run_labtalk` for "
    "operations no typed tool covers."
)

_SKILLS_CACHE: Optional[List[dict]] = None


def _first_h1(text: str) -> str:
    """Return the text of the first ``# H1`` heading, else an empty string."""
    for line in text.splitlines():
        m = re.match(r"^#\s+(.*\S)\s*$", line)
        if m:
            return m.group(1).strip()
    return ""


def _summary(text: str) -> str:
    """One-line description.

    Prefer the first content line of a ``## When To Use`` section; otherwise the
    first non-heading, non-empty paragraph line.
    """
    lines = text.splitlines()
    # Look for a "When To Use" section (any heading level).
    for i, line in enumerate(lines):
        if re.match(r"^#{1,6}\s+when to use\s*$", line.strip(), re.IGNORECASE):
            for follow in lines[i + 1:]:
                stripped = follow.strip()
                if stripped and not stripped.startswith("#"):
                    return stripped
            break
    # Fallback: first non-heading, non-empty line.
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            return stripped
    return ""


def discover_skills() -> List[dict]:
    """Discover every bundled ``*.md`` skill.

    Returns a cached, name-sorted list of
    ``{"name", "title", "summary", "content"}`` dicts. ``name`` is the filename
    stem; ``title`` the first H1; ``summary`` a one-line description; ``content``
    the full markdown.
    """
    global _SKILLS_CACHE
    if _SKILLS_CACHE is not None:
        return _SKILLS_CACHE

    skills: List[dict] = []
    skills_dir = resources.files(__package__).joinpath("skills")
    for entry in skills_dir.iterdir():
        if not entry.name.endswith(".md"):
            continue
        content = entry.read_text(encoding="utf-8")
        name = entry.name[: -len(".md")]
        skills.append({
            "name": name,
            "title": _first_h1(content) or name,
            "summary": _summary(content),
            "content": content,
        })

    skills.sort(key=lambda s: s["name"])
    _SKILLS_CACHE = skills
    return _SKILLS_CACHE


def register_skills(mcp) -> None:
    """Register ``list_skills`` / ``get_skill`` (and ``skill://`` resources).

    Idempotent-ish: callers register on a fresh FastMCP per process. Resource
    registration degrades gracefully if the installed FastMCP lacks it.
    """

    def list_skills() -> str:
        """List the bundled origin-pro skills with when to use each.

        Skills record how Origin behaves when scripted (verified recipes,
        silent no-ops, unit traps) and the figure conventions this server is
        tuned for. Load one with ``get_skill(name)``. Reads bundled files
        only; does not touch Origin.
        """
        skills = discover_skills()
        if not skills:
            return "No skills are bundled with this server."
        lines = ["Available origin-pro skills (call get_skill(name) for details):", ""]
        for s in skills:
            lines.append(f"- {s['name']}: {s['title']}")
            if s["summary"]:
                lines.append(f"    {s['summary']}")
        return "\n".join(lines)

    def get_skill(name: str) -> str:
        """Return the full markdown of a bundled skill by name.

        Use the names from ``list_skills``. On an unknown name, returns an error
        that lists the valid skill names.
        """
        for s in discover_skills():
            if s["name"] == name:
                return s["content"]
        valid = ", ".join(s["name"] for s in discover_skills()) or "(none)"
        return (
            f"Unknown skill: {name!r}. No skill by that name is bundled. "
            f"Call list_skills, then get_skill with one of: {valid}."
        )

    mcp.add_tool(list_skills, name="list_skills")
    mcp.add_tool(get_skill, name="get_skill")

    _register_resources(mcp)


def _register_resources(mcp) -> None:
    """Register each skill as a ``skill://<name>`` resource if supported.

    The tools are the priority; if this FastMCP build can't register resources
    cleanly, skip silently and keep the tools working.
    """
    try:
        from mcp.server.fastmcp.resources import FunctionResource
        from pydantic import AnyUrl
    except Exception:
        return

    for s in discover_skills():
        content = s["content"]

        def _read(_content=content) -> str:
            return _content

        try:
            mcp.add_resource(FunctionResource(
                uri=AnyUrl(f"skill://{s['name']}"),
                name=f"skill:{s['name']}",
                description=s["summary"] or s["title"],
                mime_type="text/markdown",
                fn=_read,
            ))
        except Exception:
            # Resource registration is best-effort; tools remain the priority.
            return
