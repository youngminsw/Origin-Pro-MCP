from __future__ import annotations

import os
import re
import sys
from collections.abc import Iterable
from typing import Final

_NAME_RE: Final = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_WSL_PATH_RE: Final = re.compile(r"^/mnt/([A-Za-z])(/.*)?$")
_VARIABLE_RE: Final = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*[$]?$")
_STRING_BLOCKLIST: Final = {'"', "\r", "\n"}

# Confirm-list: real LabTalk command tokens that can touch files, the project,
# or the system. Everything else (incl. save, saveAs, open, expGraph, file) is
# allowed by default. Each entry is (human label, pattern matched against the
# string/comment-stripped script). Anchored to whole-token / command syntax so
# identifiers merely *containing* a gated word (e.g. `delete_me`, `myrun`,
# `system_id`) never trigger.
_LABTALK_CONFIRM_PATTERNS: Final = (
    ("system.", re.compile(r"\bsystem\s*\.", re.IGNORECASE)),
    ("system(", re.compile(r"\bsystem\s*\(", re.IGNORECASE)),
    ("run.section", re.compile(r"\brun\s*\.\s*section\b", re.IGNORECASE)),
    ("run -", re.compile(r"\brun\s*-\s*[A-Za-z]", re.IGNORECASE)),
    ("dll", re.compile(r"\bdll\b", re.IGNORECASE)),
    ("dde", re.compile(r"\bdde\b", re.IGNORECASE)),
    ("getfilename", re.compile(r"\bgetfilename\b", re.IGNORECASE)),
    ("getsavename", re.compile(r"\bgetsavename\b", re.IGNORECASE)),
    ("doc -s", re.compile(r"\bdoc\s*-\s*s\b", re.IGNORECASE)),
    ("doc -n", re.compile(r"\bdoc\s*-\s*n\b", re.IGNORECASE)),
    ("del/delete", re.compile(r"\b(?:del|delete)\b", re.IGNORECASE)),
    ("win -c/-cd/-ct", re.compile(r"\bwin\s*-\s*c[dt]?\b", re.IGNORECASE)),
    ("label -r", re.compile(r"\blabel\s*-\s*r\b", re.IGNORECASE)),
)


def labtalk_name(value: str, field: str) -> str:
    if not _NAME_RE.fullmatch(value):
        msg = (
            f"{field} must start with a letter or underscore and contain only "
            "letters, numbers, and underscores."
        )
        raise ValueError(msg)
    return value


def labtalk_string(value: str, field: str) -> str:
    if any(char in value for char in _STRING_BLOCKLIST):
        msg = f'{field} cannot contain double quotes or line breaks.'
        raise ValueError(msg)
    return f'"{value}"'


# Origin text-object markup escapes that render safely. Anything else — most
# importantly ``\q()``, which invokes LaTeX and pops a blocking "Get MiKTeX
# Path" modal that WEDGES Origin (and the daemon) — must never reach a text
# object. Confirmed safe on Origin 2020: b, i, u, +, -, g, f (\f:Font).
_SUPPORTED_TEXT_ESCAPES = frozenset("biu+-gf")
_TEXT_ESCAPE_RE = re.compile(r"\\([^\s(])")


def validate_text_escapes(value: str, field: str) -> None:
    """Reject unsupported Origin text-markup escapes in label/title/annotation
    text. ``\\q()`` in particular triggers a blocking LaTeX/MiKTeX modal that
    wedges Origin, so it is refused BEFORE it ever reaches the COM layer."""
    for match in _TEXT_ESCAPE_RE.finditer(value):
        ch = match.group(1)
        if ch not in _SUPPORTED_TEXT_ESCAPES:
            raise ValueError(
                f"{field} contains an unsupported Origin text escape '\\{ch}'. "
                "Supported: \\b (bold), \\i (italic), \\u (underline), "
                "\\+ (superscript), \\- (subscript), \\g (Greek/symbol), "
                "\\f:Font (font). ('\\q' invokes LaTeX and hangs Origin.) "
                "Remove or replace it."
            )


def labtalk_text(value: str, field: str) -> str:
    """``labtalk_string`` for TEXT-OBJECT content (axis labels, titles,
    annotations): additionally rejects unsupported markup escapes that can
    wedge Origin via a modal dialog."""
    validate_text_escapes(value, field)
    return labtalk_string(value, field)


def labtalk_formula(value: str, field: str) -> str:
    """Validate a column-formula expression (e.g. 'col(1)^2 + col(2)').

    Blocks statement separators and quotes so the expression cannot break
    out of `col(N) = <expr>;`.
    """
    if any(char in value for char in _STRING_BLOCKLIST) or ";" in value:
        msg = f"{field} cannot contain quotes, line breaks, or ';'."
        raise ValueError(msg)
    if not value.strip():
        msg = f"{field} cannot be empty."
        raise ValueError(msg)
    return value


def labtalk_variable(value: str, field: str) -> str:
    if not _VARIABLE_RE.fullmatch(value):
        msg = f"{field} must be a LabTalk variable name."
        raise ValueError(msg)
    return value


def windows_path(value: str, field: str) -> str:
    """Normalize a user-supplied path to a Windows-writable path.

    Strips stray quotes/whitespace and converts WSL-style paths
    (/mnt/c/Users/...) to Windows form (C:\\Users\\...) so agents running
    in WSL can pass their native paths. A bare POSIX path outside /mnt (e.g.
    /tmp/x, no drive letter available) is translated to the UNC form
    \\\\wsl.localhost\\<distro>\\... (probe-confirmed on Origin 2020: expGraph
    can write there and the file genuinely lands on the WSL side) when a
    distro name is known via ORIGIN_PRO_MCP_WSL_DISTRO or (if the daemon
    inherited it) WSL_DISTRO_NAME; otherwise it is rejected with a hint to set
    that env var.
    """
    path = value.strip().strip('"').strip("'")
    if not path:
        msg = f"{field} cannot be empty."
        raise ValueError(msg)
    match = _WSL_PATH_RE.fullmatch(path)
    if match:
        drive = match.group(1).upper()
        rest = (match.group(2) or "/").replace("/", "\\")
        path = f"{drive}:{rest}"
    elif path.startswith("/") and sys.platform == "win32":
        # On the Windows daemon a POSIX path like /tmp/x silently resolves
        # against the current drive (C:\tmp\x): Origin would "successfully"
        # write somewhere the WSL-side agent can never see. Translate it to
        # a \\wsl.localhost\<distro>\... UNC path when a distro name is known
        # (this only matters on the real Windows daemon — on WSL/Linux itself
        # the path is already directly accessible, so it is left untouched).
        distro = os.environ.get("ORIGIN_PRO_MCP_WSL_DISTRO") or os.environ.get("WSL_DISTRO_NAME")
        if distro:
            path = f"\\\\wsl.localhost\\{distro}" + path.replace("/", "\\")
        else:
            msg = (
                f"{field} is a Linux/WSL path ({path}) that Origin — a Windows "
                "process — cannot access. Use a Windows path (C:\\...) or a "
                "/mnt/<drive>/... path, or set ORIGIN_PRO_MCP_WSL_DISTRO=<distro> "
                "to auto-map to \\\\wsl.localhost."
            )
            raise ValueError(msg)
    return path


def labtalk_path(value: str, field: str) -> str:
    if any(char in value for char in _STRING_BLOCKLIST):
        msg = f'{field} cannot contain double quotes or line breaks.'
        raise ValueError(msg)
    escaped = value.replace("\\", "\\\\")
    return f'"{escaped}"'


def labtalk_choice(value: str, allowed: Iterable[str], field: str) -> str:
    allowed_values = set(allowed)
    if value not in allowed_values:
        msg = f"{field} must be one of: {', '.join(sorted(allowed_values))}."
        raise ValueError(msg)
    return value


def positive_column(value: int, field: str) -> int:
    if value < 1:
        msg = f"{field} must be a 1-based column index."
        raise ValueError(msg)
    return value


def positive_int(value: int, field: str) -> int:
    if value < 1:
        msg = f"{field} must be positive."
        raise ValueError(msg)
    return value


def _strip_strings_and_comments(script: str) -> str:
    """Blank out string literals and comments so command keywords inside them
    never trigger the confirm gate.

    LabTalk strings are double-quoted; comments are `//` to end-of-line and
    `/* ... */` blocks. Stripped characters are replaced with spaces (newlines
    preserved) so token boundaries and match offsets are unaffected.
    """
    out: list[str] = []
    i = 0
    n = len(script)
    state = "normal"  # normal | string | line_comment | block_comment
    while i < n:
        ch = script[i]
        nxt = script[i + 1] if i + 1 < n else ""
        if state == "normal":
            if ch == '"':
                state = "string"
                out.append(" ")
                i += 1
            elif ch == "/" and nxt == "/":
                state = "line_comment"
                out.append("  ")
                i += 2
            elif ch == "/" and nxt == "*":
                state = "block_comment"
                out.append("  ")
                i += 2
            else:
                out.append(ch)
                i += 1
        elif state == "string":
            # An unescaped double-quote closes the string.
            if ch == '"':
                state = "normal"
            out.append("\n" if ch == "\n" else " ")
            i += 1
        elif state == "line_comment":
            if ch == "\n":
                state = "normal"
                out.append("\n")
            else:
                out.append(" ")
            i += 1
        else:  # block_comment
            if ch == "*" and nxt == "/":
                state = "normal"
                out.append("  ")
                i += 2
            else:
                out.append("\n" if ch == "\n" else " ")
                i += 1
    return "".join(out)


def classify_labtalk_script(script: str) -> tuple[bool, bool, str]:
    """Classify a LabTalk script for the run_labtalk confirm gate.

    Allow-by-default: most operations (including save, saveAs, open, expGraph,
    file) run freely. Only a narrow confirm-list of real command tokens that can
    touch files, the project, or the system requires explicit confirmation.

    String literals and comments are stripped before scanning, so keywords
    inside `"..."`, `//`, or `/* ... */` never trigger.

    Returns:
        (ok, requires_confirm, reason). ``ok`` is always True — nothing is
        hard-blocked. ``requires_confirm`` is True when a confirm-list token is
        present, and ``reason`` names the earliest such token (else "").
    """
    cleaned = _strip_strings_and_comments(script)
    best_start: int | None = None
    best_label = ""
    for label, pattern in _LABTALK_CONFIRM_PATTERNS:
        match = pattern.search(cleaned)
        if match is not None and (best_start is None or match.start() < best_start):
            best_start = match.start()
            best_label = label
    if best_start is None:
        return (True, False, "")
    return (True, True, best_label)


def split_labtalk_statements(script: str) -> list[str]:
    """Split a LabTalk script into top-level statements on `;`.

    A `;` inside a string literal, a `//`/`/* */` comment, or a `{...}` block
    is NOT a split point — brace depth and string/comment state are tracked
    via the same scan :func:`_strip_strings_and_comments` uses, so this stays
    in sync with the confirm-gate's tokenization instead of a naive
    ``script.split(';')``.

    Returns the original (unstripped) substrings, each including its
    trailing `;` where present. Whitespace-only fragments are dropped, so a
    script with no real statements (or a single trailing `;`) yields at most
    one entry.
    """
    cleaned = _strip_strings_and_comments(script)
    statements: list[str] = []
    start = 0
    depth = 0
    for i, ch in enumerate(cleaned):
        if ch == "{":
            depth += 1
        elif ch == "}":
            if depth > 0:
                depth -= 1
        elif ch == ";" and depth == 0:
            statements.append(script[start:i + 1])
            start = i + 1
    tail = script[start:]
    if tail.strip():
        statements.append(tail)
    return [s for s in statements if s.strip()]


def safe_labtalk_script(value: str) -> str:
    """Backward-compatible pass-through shim.

    The old denylist hard-blocked core operations; gating now flows through
    :func:`classify_labtalk_script` (used by run_labtalk with a confirm flag).
    Nothing is hard-blocked here, so the script is returned unchanged.
    """
    return value
