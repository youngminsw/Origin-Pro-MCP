"""Direct command-line control of Origin Pro — no MCP client required.

The MCP tools are plain functions registered with FastMCP. This CLI
reflects over that single registry and exposes every tool as a
subcommand, so the repo alone is enough to drive Origin:

    python -m origin_pro_mcp.cli list
    python -m origin_pro_mcp.cli list_worksheets
    python -m origin_pro_mcp.cli apply_publication_style --graph_name FigA
    python -m origin_pro_mcp.cli apply_publication_style --json '{"graph_name": "FigA", "x_label": "Temperature (K)"}'

Use --json for any argument containing spaces or special characters
(axis labels, file paths). Like the MCP server, this must run under
Windows Python with pywin32 and a licensed Origin install.
"""
from __future__ import annotations

import inspect
import json
import sys
import types
import typing


def _tools() -> dict:
    """Map of tool name -> underlying callable, from the FastMCP registry."""
    from . import server  # noqa: F401 — importing registers all tools
    from .app import mcp

    manager = mcp._tool_manager
    return {name: tool.fn for name, tool in manager._tools.items()}


def _unwrap_optional(annotation):
    """For `X | None` / `Optional[X]`, return X; otherwise the annotation."""
    origin = typing.get_origin(annotation)
    union_types = (typing.Union, getattr(types, "UnionType", None))
    if origin in union_types:
        non_none = [a for a in typing.get_args(annotation) if a is not type(None)]
        if non_none:
            return non_none[0]
    return annotation


def _coerce(value: str, annotation):
    """Coerce a CLI string to the parameter's annotated type."""
    target = _unwrap_optional(annotation)
    if target is bool:
        return str(value).strip().lower() in ("1", "true", "yes", "on")
    if target is int:
        return int(value)
    if target is float:
        return float(value)
    return str(value)


def _parse_kv(args: list, sig: inspect.Signature) -> dict:
    """Parse `--key value` / `--key=value` pairs against a tool signature."""
    kwargs: dict = {}
    i = 0
    while i < len(args):
        token = args[i]
        if not token.startswith("--"):
            raise ValueError(f"unexpected argument: {token!r} (expected --name value)")
        key = token[2:]
        if "=" in key:
            key, raw = key.split("=", 1)
            i += 1
        else:
            if i + 1 >= len(args):
                raise ValueError(f"missing value for --{key}")
            raw = args[i + 1]
            i += 2
        if key not in sig.parameters:
            allowed = ", ".join(sig.parameters)
            raise ValueError(f"unknown argument --{key}; valid: {allowed}")
        kwargs[key] = _coerce(raw, sig.parameters[key].annotation)
    return kwargs


def _format_tools(tools: dict) -> str:
    lines = ["Available tools (python -m origin_pro_mcp.cli <tool> [args]):", ""]
    for name in sorted(tools):
        fn = tools[name]
        params = ", ".join(str(p) for p in inspect.signature(fn).parameters.values())
        doc = (fn.__doc__ or "").strip().splitlines()
        summary = doc[0] if doc else ""
        lines.append(f"  {name}({params})")
        if summary:
            lines.append(f"      {summary}")
    return "\n".join(lines)


def main(argv: list | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    tools = _tools()

    if not argv or argv[0] in ("-h", "--help", "help", "list", "--list"):
        print(_format_tools(tools))
        return 0

    name, rest = argv[0], argv[1:]
    if name not in tools:
        print(f"Unknown tool: {name!r}. Run 'list' to see options.", file=sys.stderr)
        return 2

    fn = tools[name]
    sig = inspect.signature(fn)
    try:
        if rest and rest[0] == "--json":
            if len(rest) < 2:
                raise ValueError("--json requires a JSON object argument")
            kwargs = json.loads(rest[1])
            if not isinstance(kwargs, dict):
                raise ValueError("--json argument must be a JSON object")
        else:
            kwargs = _parse_kv(rest, sig)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"Argument error: {exc}", file=sys.stderr)
        return 2

    try:
        result = fn(**kwargs)
    except TypeError as exc:
        print(f"Argument error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # tool/Origin runtime failure
        print(f"ERROR ({type(exc).__name__}): {exc}", file=sys.stderr)
        return 1

    print(result if isinstance(result, str) else json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
