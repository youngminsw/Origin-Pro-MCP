# Work Log

## 2026-07-11 — g2-g8 sweep confirmed-fixes round (agent: impl-g2g8)

Scope: the three CONFIRMED items assigned from the g2-g8 sweep ("NEW ISSUES
from the g2-g8 sweep" section of `ORIGIN_MCP_ISSUES.md`): #18 (missing values
in `set_worksheet_data`), the Misc `create_worksheet`-on-existing-book bug,
and a docs batch for #16/#17a/#17c/#20/Misc. Items #15/16b/17b/19/21/20's
"impossible" verdicts were explicitly out of scope (a separate agent
re-probed those in parallel).

### Task 1 — #18: `set_worksheet_data` accepts JSON null / NaN as Origin missing values

- **What:** `set_worksheet_data` (`src/origin_pro_mcp/tools/worksheet.py`)
  now accepts `null` (and a bare `NaN` token, which `json.loads` parses to
  `float('nan')`) in the `columns` payload. Each null/NaN cell is written
  with a numeric placeholder in the bulk `PutWorksheet` call, then
  overwritten with Origin's real missing-value sentinel via LabTalk
  `col(<c>)[<r>]=0/0;` — the reporter's confirmed live recipe (`NANUM` is a
  no-op; `0/0` works). Docstring updated to describe this and note the old
  "split every series at gaps" workaround is obsolete.
- **Commit:** `a341849` — "fix: set_worksheet_data accepts null/NaN as Origin
  missing values (#18)"
- **HOW VERIFIED:**
  - Fake tests (`tests/test_worksheet_ops.py`):
    `test_set_worksheet_data_accepts_null_as_missing_value`,
    `test_set_worksheet_data_accepts_nan_as_missing_value`,
    `test_set_worksheet_data_no_missing_cells_emits_no_labtalk` — all assert
    the exact `col(N)[R]=0/0;` LabTalk is emitted for exactly the null/NaN
    positions and nowhere else.
  - WSL fake suite after this task: `587 passed, 29 skipped` (0 fails).
  - Live test (`tests/test_live_worksheet.py::test_set_worksheet_data_null_becomes_missing_value`,
    `requires_origin`, isolated `live_origin` fixture): writes
    `[[1, 2, None, 4]]` to a fresh worksheet, reads it back via
    `get_worksheet_data`, and asserts the result is
    `[[1.0, 2.0, None, 4.0]]` — the null round-trips to a real Origin
    missing value and the control cells are untouched. Run on the Windows
    side (`C:\Users\swym4\opm_dev`, rsynced from this WSL working tree):
    ```
    tests/test_live_worksheet.py::test_set_worksheet_data_null_becomes_missing_value PASSED
    ```

### Task 2 — Misc: `create_worksheet(book, sheet)` on an existing book adds a sheet instead of a new book

- **What:** `create_worksheet` (`src/origin_pro_mcp/tools/worksheet.py`) now
  checks `workbook_names()` first. If `book_name` is already open: raise if
  `sheet_name` already exists in that book (via the existing `_sheet_names`
  helper + a new `_find_workbook_page` lookup), otherwise activate the book
  and run LabTalk `newsheet name:="<sheet>" cols:=2;` to add the sheet to the
  SAME book, returning `added_to_existing_book: true`. Unchanged behavior
  (new `CreatePage`, `added_to_existing_book: false`) when the book doesn't
  exist yet.
- **Commit:** `9634246` — "fix: create_worksheet adds a sheet to an existing
  book instead of a new one"
- **HOW VERIFIED:**
  - Live probe first (`/mnt/c/Users/swym4/probe_newsheet.py` /
    `probe_newsheet2.py`, deleted after use — not committed): confirmed
    `newsheet name:=Sheet2 cols:=2;` (and the quoted form the tool actually
    emits, `newsheet name:="Sheet2" cols:=2;`) adds a second sheet to an
    EXISTING workbook when run against the activated book window, on a
    throwaway isolated `DispatchEx` Origin instance. Verified via the
    `layer$(k).name$` enumeration returning `Sheet1|Sheet2|`.
  - Fake tests (`tests/test_worksheet_ops.py`):
    `test_create_worksheet_adds_sheet_to_existing_book` (asserts
    `added_to_existing_book: true` and the `newsheet name:="Sheet2"` LabTalk
    call), `test_create_worksheet_existing_sheet_raises` (asserts
    `ValueError` with "already has a sheet" when the sheet name collides).
    Existing tests updated for the new `added_to_existing_book` field.
  - WSL fake suite after this task: `589 passed, 30 skipped` (0 fails).
  - Live test (`tests/test_live_worksheet.py::test_create_worksheet_existing_book_adds_sheet_not_new_book`):
    creates a book, calls `create_worksheet(book, "Sheet2")` again on the
    SAME name, asserts `added_to_existing_book is True` and `renamed is
    False`, then confirms via `list_worksheets` that there is exactly ONE
    workbook with that name and its `sheets` set is `{"Sheet1", "Sheet2"}`
    (i.e. no second auto-renamed book was created). Run on the Windows side:
    ```
    tests/test_live_worksheet.py::test_create_worksheet_existing_book_adds_sheet_not_new_book PASSED
    ```

### Task 3 — Docs batch: encode g2-g8's verified recipes (#16, #17a, #17c, #20, Misc)

- **What (docs-only, no behavior change):**
  - `add_second_y_axis` docstring (`src/origin_pro_mcp/tools/graph.py`): the
    right-axis TITLE recipe via the `YR` text object (`yr.text$="\b(...)"`
    for bold, `YR.color`, `YR.fsize`) — `yl` targets layer 2's hidden LEFT
    title, not the visible right one; `YR.bold=1` is a no-op.
  - `src/origin_pro_mcp/skills/publication-figure.md` — Origin COM Notes
    table gained rows for: (a) the YR right-axis-title recipe + which parts
    of the right axis remain unreachable (tick labels, axis line); (b)
    `layer.y.inc` needing the graph active + `doc -uw`; (c) colorbar overflow
    fixed via LAYER geometry (`layer.left/width/top/height`), not the
    unaddressable colorbar object; (d) the matplotlib in-corner tick-label
    nudge being moot in Origin (labels don't collide the way matplotlib's
    do); (e) the tick-anchor-at-multiples-of-`inc` workaround
    (`layer.x.label.formula$="x+1"`) and that `\x()` is rejected in axis
    label text (already enforced by `validate_text_escapes`) — use the
    unicode character directly. Also added inline pointers from the
    "Multi-panel and axis control" and "Surfaces, contours, heatmaps"
    sections to these notes.
  - `tests/test_doc_registry.py`: added `"x"` to the `NOT_TOOLS` allowlist
    (the new `\x()` markup mention in the skill doc was being parsed as a
    reference to a nonexistent tool named `x` by the doc-registry guard).
- **Commit:** `e3f3278` — "docs: encode g2-g8 sweep's verified recipes (#16
  #17a #17c #20 + Misc)"
- **HOW VERIFIED:** docs-only change; verification is that the WSL suite
  (including `tests/test_doc_registry.py`, which parses these exact docs for
  tool-name references) stays green: `589 passed, 30 skipped` (0 fails),
  `ruff check` clean on the touched files.

### Non-negotiables checked

- WSL fake suite (venv at `/tmp/opm_venv`, Python 3.11): green after every
  commit, final count `589 passed, 30 skipped` (baseline before this round
  was `584 passed, 28 skipped`).
- `ruff check` clean on every file touched across all three tasks.
- Full live suite re-run after all three commits, on the Windows side against
  `C:\Users\swym4\opm_dev` (rsynced from this WSL tree, `C:\Users\swym4\Origin-Pro-MCP`
  was never touched): `pytest -m requires_origin tests/ -v` →
  `1 failed, 29 passed` (589 non-live tests deselected). The 1 failure,
  `tests/test_live_styling.py::test_axis_tick_top_none_removes_marks_keeps_bottom_labels`,
  is OUT OF SCOPE for this round (it belongs to the prior styling-report
  round's per-side-tick-removal task) and touches code this round never
  modified (`axis`/tick handling in `graph.py`/`style.py`) — the failure is a
  byte-identical baseline-vs-after export (the tick-mark change was too
  subtle to move any pixel at this test's default frame state), consistent
  with that test's own docstring caveat about subtle changes. Not
  investigated further here; flagging for whoever owns that task. All 2
  live tests written in THIS round passed.
- Zero leftover Origin processes: baseline was 4 pre-existing `Origin64.exe`
  processes (not spawned by this work, left untouched) before, during, and
  after every live-test run in this round — confirmed via `tasklist` after
  each isolated-fixture test run and after the final full live-suite run.
- Did NOT push (per instructions).
