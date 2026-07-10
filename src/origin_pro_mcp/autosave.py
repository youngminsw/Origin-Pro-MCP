"""Autosave-before-destructive-op AND periodic autosave: save the open project
IN PLACE (its own file, same name) so a wedge / crash / user mistake can be rolled
back by reloading the file. Autosave NEVER writes a differently-named copy.

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
* The actual save primitive (:func:`save_in_place`) does a plain ``Save`` of the
  project to its own path (never a copy), guarded so an empty/blanked project can
  never overwrite a real file (N5). Runtime activation is via ``ORIGIN_PRO_MCP_AUTOSAVE``.

All classification/policy/has-work logic in this module is COM-free.
"""
from __future__ import annotations

import os

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
    ``ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED``  default ON — when a required in-place
                                          save fails, surface an error instead of
                                          silently proceeding with the destructive op.

    Autosave saves the project IN PLACE (to its own file, same name) — it never
    writes a differently-named copy.
    """

    def __init__(self, enabled: bool = True, required: bool = True):
        self.enabled = enabled
        self.required = required

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
        return cls(enabled=enabled, required=required)


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


def save_in_place(origin, remembered_path: Optional[str] = None) -> Optional[bool]:
    """Save the project to ITS OWN file — in place, same name. NOT a copy.

    This is what "autosave" means: a plain Save of the open project, so the user
    ends up with their project up to date, NOT a pile of differently-named
    ``*.autosave-*.opju`` copies (which also rebound the project's identity to the
    backup name — the behavior this replaces).

    Two safety guards keep the N5 data-destruction incident from recurring:
      * NEVER save an EMPTY/blanked project (0 windows, e.g. after a flaky
        empty-load) — that is exactly what once wrote a 450-byte file over a real
        579 KB .opju.
      * ONLY overwrite an EXISTING file — never create a new-named file. If the
        project has no on-disk file yet (never saved), autosave is a no-op.

    ``origin.Save("")`` does NOT save when a path exists (verified on Origin 2020),
    so we save to the project's actual full path: the remembered load/save path
    when known, else reconstructed from the ``%X`` (folder) + ``%G`` (name)
    LabTalk registers. Saving to the current path is in place — no rename, no
    rebind.

    Returns a TRI-STATE result (issue #12 fix — a caller must be able to tell
    "nothing to protect" from "a real save attempt failed"):
      * ``True``  — a save actually happened.
      * ``None``  — nothing on disk to protect: an empty/blanked project, OR a
        never-saved project with no on-disk file yet. Not a failure — there was
        no in-place target to guard, so a preflight gate must PROCEED, not block.
      * ``False`` — a real save attempt was made (an on-disk file exists) and it
        failed. A REQUIRED preflight gate should block on this.
    """
    try:
        pages = (origin.WorksheetPages.Count + origin.GraphPages.Count
                 + origin.MatrixPages.Count)
    except Exception:
        pages = -1
    if pages <= 0:
        return None  # empty/blanked project -> nothing to protect, never overwrite (N5)
    path = remembered_path
    if not path:
        try:
            folder = origin.LTStr("%X")
            name = origin.LTStr("%G")
            if folder and name:
                path = os.path.join(folder, f"{name}.opju")
        except Exception:
            path = None
    if not path or not os.path.isfile(path):
        return None  # never-saved project -> no in-place target; nothing to protect
    try:
        return bool(origin.Save(path))
    except Exception:
        return False