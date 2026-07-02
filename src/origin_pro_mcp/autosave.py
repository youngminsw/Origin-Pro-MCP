"""Autosave-before-destructive-op: snapshot a recoverable project copy BEFORE a
destructive operation so a wedge / crash / user mistake can be rolled back.

Design (from the downgrade plan, Feature 2):

* Two snapshot triggers, both evaluated at DAEMON DISPATCH (there is deliberately
  NO ``execute_labtalk`` hook — that would double-fire on the many internal
  scripts tools issue):
    - typed destructive tools (``delete_graph``, ``remove_plot``,
      ``manage_columns``, ``load_project``, ``new_project``) — snapshot at dispatch;
    - the OVERWRITE-only writers ``set_worksheet_data`` / ``set_matrix_data`` —
      snapshot ONLY when the target sheet already holds data (an in-worker
      emptiness probe); ``worksheet_to_matrix`` is EXCLUDED;
    - ``run_labtalk`` whose script carries a destructive LabTalk token AND runs
      with ``confirm=True`` (i.e. it will actually execute).
* A conservative has-work predicate: any successful non-read-only tool marks the
  session as having work worth protecting; ``new_project`` resets it. A snapshot
  is skipped when there is no work (a fresh/empty project is not worth backing up).
* The actual save primitive (:func:`save_copy`) is PLUGGABLE and its live
  correctness on Origin 2020 is gated on a save-copy spike; everything else here
  is COM-free and unit-tested. Runtime activation in the daemon is default-OFF
  until that spike confirms the primitive (see ``ORIGIN_PRO_MCP_AUTOSAVE``).

All classification/policy/has-work logic in this module is COM-free.
"""
from __future__ import annotations

import os
import time
from typing import Optional

from .labtalk_safe import _LABTALK_CONFIRM_PATTERNS, _strip_strings_and_comments

# Tools that never mutate recoverable project state — they must NOT set has-work
# and never trigger a snapshot.
READONLY_TOOLS = frozenset({
    "get_labtalk_variable", "get_matrix_data", "get_plot_info", "get_plot_names",
    "get_worksheet_data", "list_worksheets", "list_fitting_functions",
    "find_plot_column", "stats", "export_all_graphs", "export_graph",
    "export_graph_to_file", "export_worksheet", "save_graph_template",
    "save_project",
})

# Typed tools whose effect can destroy/replace project objects — snapshot before.
DESTRUCTIVE_TYPED_TOOLS = frozenset({
    "delete_graph", "remove_plot", "manage_columns", "load_project", "new_project",
})

# Writers that OVERWRITE existing cells — snapshot only when the target is
# non-empty (an emptiness probe). Keyed to the kwarg that names the book.
OVERWRITE_TOOLS = frozenset({"set_worksheet_data", "set_matrix_data"})

# Explicitly excluded (a derived-output tool that does not destroy its source).
EXCLUDED_TOOLS = frozenset({"worksheet_to_matrix"})

# Destructive LabTalk tokens (the project/data-destroying subset of the confirm
# gate): delete, window-close/-delete, and new-document. system/run/dll/dde/
# getfilename/getsavename/doc -s/label -r are not project-data destroyers.
_DESTRUCTIVE_LABTALK_LABELS = frozenset({"del/delete", "win -c/-cd/-ct", "doc -n"})
_DESTRUCTIVE_LABTALK_PATTERNS = tuple(
    pat for label, pat in _LABTALK_CONFIRM_PATTERNS
    if label in _DESTRUCTIVE_LABTALK_LABELS
)


def _falsey_env(raw: Optional[str]) -> bool:
    return raw is None or raw.strip().lower() in ("off", "false", "no", "0", "")


class AutosavePolicy:
    """Env-driven autosave configuration.

    ``ORIGIN_PRO_MCP_AUTOSAVE``           opt-out (default ON) — set off/false/no/0 to disable.
    ``ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED``  default ON — when a required snapshot
                                          fails, surface an error instead of
                                          silently proceeding with the destructive op.
    ``ORIGIN_PRO_MCP_AUTOSAVE_RETENTION`` how many backups to keep (default 3).
    ``ORIGIN_PRO_MCP_AUTOSAVE_DIR``       backup directory (default alongside the
                                          project, or CWD when the project is unsaved).
    """

    def __init__(self, enabled: bool = True, required: bool = True,
                 retention: int = 3, backup_dir: Optional[str] = None):
        self.enabled = enabled
        self.required = required
        self.retention = max(1, retention)
        self.backup_dir = backup_dir

    @classmethod
    def from_env(cls, environ: Optional[dict] = None) -> "AutosavePolicy":
        env = environ if environ is not None else os.environ
        # Opt-OUT: enabled unless explicitly set to an off value (unset -> ON).
        raw = env.get("ORIGIN_PRO_MCP_AUTOSAVE")
        enabled = not (raw is not None
                       and raw.strip().lower() in ("off", "false", "no", "0"))
        # REQUIRED defaults ON (unset -> required); explicit off/false/no/0 clears it.
        req_raw = env.get("ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED")
        required = True if req_raw is None else not _falsey_env(req_raw)
        try:
            retention = int(env.get("ORIGIN_PRO_MCP_AUTOSAVE_RETENTION", "3"))
        except (TypeError, ValueError):
            retention = 3
        backup_dir = env.get("ORIGIN_PRO_MCP_AUTOSAVE_DIR") or None
        return cls(enabled=enabled, required=required,
                   retention=retention, backup_dir=backup_dir)


def is_readonly(name: str) -> bool:
    return name in READONLY_TOOLS


def classify_autosave_labtalk(script: str) -> bool:
    """True when a LabTalk script carries a project/data-destroying token
    (delete, win -c/-cd/-ct, doc -n). String literals and comments are stripped
    first (reusing the confirm-gate masking) so tokens inside "..."/// never
    trigger a false snapshot."""
    if not script:
        return False
    cleaned = _strip_strings_and_comments(script)
    return any(pat.search(cleaned) for pat in _DESTRUCTIVE_LABTALK_PATTERNS)


def _target_has_data(name: str, kwargs: dict, origin) -> bool:
    """Emptiness probe for the OVERWRITE writers: True when the target sheet
    already holds data (so the write overwrites it). Conservative: if the probe
    cannot determine emptiness, treat it as having data (snapshot)."""
    if origin is None:
        return True
    book = kwargs.get("book_name")
    if not book:
        return True
    try:
        if name == "set_worksheet_data":
            sheet = kwargs.get("sheet_name") or "Sheet1"
            data = origin.GetWorksheet(f"[{book}]{sheet}")
        else:  # set_matrix_data
            data = origin.GetMatrix(f"[{book}]MSheet1")
    except Exception:
        return True
    if isinstance(data, int):  # HRESULT error => target not found/empty
        return False
    # A tuple/list of columns/rows: non-empty when any cell holds a value.
    try:
        return any(any(cell is not None for cell in row) for row in data)
    except TypeError:
        return bool(data)


def should_snapshot(name: str, kwargs: dict, origin=None) -> bool:
    """True when a snapshot is warranted BEFORE this dispatch (ignores has-work,
    which the caller checks separately)."""
    kwargs = kwargs or {}
    if name in EXCLUDED_TOOLS:
        return False
    if name in DESTRUCTIVE_TYPED_TOOLS:
        return True
    if name in OVERWRITE_TOOLS:
        return _target_has_data(name, kwargs, origin)
    if name == "run_labtalk":
        return bool(kwargs.get("confirm")) and classify_autosave_labtalk(
            kwargs.get("script", ""))
    return False


class HasWorkTracker:
    """Per-session flag: has any recoverable work been produced since the last
    fresh project? Marked by any successful non-read-only tool; reset on
    new_project. A snapshot is skipped while this is False."""

    def __init__(self):
        self._has_work = False

    @property
    def has_work(self) -> bool:
        return self._has_work

    def record_success(self, name: str) -> None:
        if name == "new_project":
            self._has_work = False
            return
        if not is_readonly(name):
            self._has_work = True

    def reset(self) -> None:
        self._has_work = False


def backup_path(policy: AutosavePolicy, remembered_path: Optional[str],
                now: Optional[float] = None) -> str:
    """Compute the timestamped backup path for a snapshot."""
    stamp = time.strftime("%Y%m%d-%H%M%S", time.localtime(now))
    if remembered_path:
        base = os.path.splitext(os.path.basename(remembered_path))[0]
        default_dir = os.path.dirname(remembered_path) or os.getcwd()
    else:
        base = "untitled"
        default_dir = os.getcwd()
    directory = policy.backup_dir or default_dir
    return os.path.join(directory, f"{base}.autosave-{stamp}.opju")


def prune_backups(policy: AutosavePolicy, remembered_path: Optional[str],
                  lister=None, remover=None) -> list:
    """Keep only the newest ``policy.retention`` autosave copies for this project.
    ``lister``/``remover`` are injectable for COM-free testing (default: os)."""
    lister = lister or (lambda d: [os.path.join(d, f) for f in os.listdir(d)])
    remover = remover or os.remove
    if remembered_path:
        base = os.path.splitext(os.path.basename(remembered_path))[0]
        directory = policy.backup_dir or (os.path.dirname(remembered_path) or os.getcwd())
    else:
        base = "untitled"
        directory = policy.backup_dir or os.getcwd()
    marker = f"{base}.autosave-"
    try:
        candidates = [p for p in lister(directory)
                      if os.path.basename(p).startswith(marker)]
    except OSError:
        return []
    candidates.sort()  # timestamp in the name sorts chronologically
    removed = []
    while len(candidates) > policy.retention:
        victim = candidates.pop(0)
        try:
            remover(victim)
            removed.append(victim)
        except OSError:
            pass
    return removed


def save_copy(origin, dest_path: str, remembered_path: Optional[str]) -> bool:
    """Save a recoverable backup of the current project to ``dest_path`` (inside
    the backup directory) — and NOTHING ELSE.

    N5 (data destruction) root cause: this function used to also
    ``Save(remembered_path)`` to "restore the original binding" after the backup
    Save rebinds project identity. But when a flaky empty-load had blanked the
    in-memory project, that second Save wrote the EMPTY project over the user's
    real 579 KB .opju, destroying it (450 bytes). The fix:

      * NEVER re-save the user's original file — the backup lives in the backup
        directory only (``remembered_path`` is accepted for signature/naming
        compat but is deliberately never written).
      * NEVER back up an EMPTY project (0 windows): a flaky empty-load must not
        cause any file write at all.

    ``o.Save(dest_path)`` rebinds the active project's identity to the backup
    path; that is harmless — the user's original file is never touched by
    autosave, so it can never be destroyed by a snapshot. Returns True only when
    a backup was actually written.
    """
    del remembered_path  # intentionally never written (see N5 above)
    try:
        pages = (origin.WorksheetPages.Count + origin.GraphPages.Count
                 + origin.MatrixPages.Count)
    except Exception:
        pages = -1
    if pages == 0:
        return False  # empty project -> nothing worth backing up; write NOTHING
    try:
        return bool(origin.Save(dest_path))
    except Exception:
        return False
