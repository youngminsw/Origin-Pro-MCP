"""Doc-vs-registry consistency check.

Guards against documentation drift: every tool name referenced in README.md
and the bundled publication-figure skill (call-style `name(...)` snippets or
markdown tool-table rows) must actually exist in the live MCP tool registry.
"""

import re
from pathlib import Path

import pytest

pytest.importorskip("mcp")

from origin_pro_mcp.app import mcp  # noqa: E402
import origin_pro_mcp.server  # noqa: E402,F401

REPO_ROOT = Path(__file__).resolve().parent.parent
README = REPO_ROOT / "README.md"
SKILL = REPO_ROOT / "src" / "origin_pro_mcp" / "skills" / "publication-figure.md"

# Identifiers that appear as `name(...)` in the docs but are NOT MCP tools
# (LabTalk/COM helpers, rich-text markup, or generic prose functions).
NOT_TOOLS = {
    "b",            # \b(...) bold markup
    "x",            # \x(...) markup (rejected — see publication-figure.md)
    "col",          # LabTalk col(n)
    "load",         # layer.cmap.load(...)
    "updateScale",  # layer.cmap.updateScale()
    "FindGraphLayer",
    "font",         # LabTalk font(Arial)
    "color",        # LabTalk color(r,g,b)
}

CALL_RE = re.compile(r"\b([a-zA-Z_][a-zA-Z0-9_]*)\(")
TABLE_ROW_RE = re.compile(r"^\|\s*`([a-zA-Z_][a-zA-Z0-9_]*)`\s*\|", re.MULTILINE)


def _registered_tool_names() -> set:
    return set(mcp._tool_manager._tools.keys())


def _tool_table_section(text: str) -> str:
    """Slice out the '## Available Tools' section (tool-name table rows only;
    other tables in README, e.g. plot types / env vars / color palette, use
    `name` for non-tool values and must not be scanned as tool rows)."""
    match = re.search(r"## Available Tools.*?(?=\n## |\Z)", text, re.DOTALL)
    return match.group(0) if match else ""


def _referenced_names(text: str, *, table_text: str = "") -> set:
    names = {m for m in CALL_RE.findall(text) if m not in NOT_TOOLS}
    names |= set(TABLE_ROW_RE.findall(table_text))
    return names


def test_readme_tool_references_are_registered():
    registered = _registered_tool_names()
    text = README.read_text(encoding="utf-8")
    referenced = _referenced_names(text, table_text=_tool_table_section(text))
    unknown = referenced - registered
    assert not unknown, f"README.md references unregistered tool names: {sorted(unknown)}"


def test_skill_tool_references_are_registered():
    registered = _registered_tool_names()
    text = SKILL.read_text(encoding="utf-8")
    referenced = _referenced_names(text)
    unknown = referenced - registered
    assert not unknown, (
        f"publication-figure.md references unregistered tool names: {sorted(unknown)}"
    )


def test_readme_tool_count_matches_registry():
    registered = _registered_tool_names()
    match = re.search(r"Available Tools \((\d+) total\)", README.read_text(encoding="utf-8"))
    assert match, "README.md is missing the 'Available Tools (N total)' header"
    assert int(match.group(1)) == len(registered)
