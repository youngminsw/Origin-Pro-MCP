from __future__ import annotations

import re
from collections.abc import Iterable
from typing import Final

_NAME_RE: Final = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_WSL_PATH_RE: Final = re.compile(r"^/mnt/([A-Za-z])(/.*)?$")
_VARIABLE_RE: Final = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*[$]?$")
_STRING_BLOCKLIST: Final = {'"', "\r", "\n"}
_LABTALK_BLOCKED_PATTERNS: Final = (
    re.compile(r"\b(delete|del|save|saveas|open|file|exit|dll|dde|run|expgraph)\b", re.IGNORECASE),
    re.compile(r"\b(doc)\s*-\s*[sn]\b", re.IGNORECASE),
    re.compile(r"\b(win|window)\s*-\s*c[dt]?\b", re.IGNORECASE),
    re.compile(r"\blabel\s*-\s*r\b", re.IGNORECASE),
    re.compile(r"\b(getfilename|getsavename)\b", re.IGNORECASE),
    re.compile(r"\b(system\.|system\s*\()", re.IGNORECASE),
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


def labtalk_variable(value: str, field: str) -> str:
    if not _VARIABLE_RE.fullmatch(value):
        msg = f"{field} must be a LabTalk variable name."
        raise ValueError(msg)
    return value


def windows_path(value: str, field: str) -> str:
    """Normalize a user-supplied path to a Windows path.

    Strips stray quotes/whitespace and converts WSL-style paths
    (/mnt/c/Users/...) to Windows form (C:\\Users\\...) so agents running
    in WSL can pass their native paths.
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


def safe_labtalk_script(value: str) -> str:
    for pattern in _LABTALK_BLOCKED_PATTERNS:
        match = pattern.search(value)
        if match is not None:
            msg = f"LabTalk command is blocked by the safety guard: {match.group(0)}"
            raise ValueError(msg)
    return value
