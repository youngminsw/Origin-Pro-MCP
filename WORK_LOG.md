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

## 2026-07-11 — g2-g8 A-verdict implementations (agent: probe-g2g8)

Scope: implement the phase-1 re-probe findings (three A-verdicts were broken by
live re-probe; receipts in `C:\Users\swym4\probe_out\g2g8\`). Six tasks, one
commit each, TDD + live acceptance. Base = 97dfd5b. Did NOT push.

Live testing: isolated `DispatchEx("Origin.Application")` only (never
ApplicationSI), `origin.Exit()` in finally. WSL edits rsynced into
`C:\Users\swym4\Origin-Pro-MCP` (the clone was released to me for this round);
`mcp__origin-pro__*` tools were NOT used (they bind a shared daemon Origin).

### Task 1 — #16: right-axis color on `add_second_y_axis`

- **What:** new optional `color="r,g,b"` on `add_second_y_axis`
  (`tools/graph.py`). The visible right axis is layer 2's **`y2`** (positional
  right), not its `y` (layer 2's hidden LEFT axis — the reporter's trap). Emits
  `win -a <g>; page.active=<layer>; layer.y2.color=color(...);
  layer.y2.label.color=color(...); doc -uw;` (separate statements — a 3+
  `layer.*` batch can fail as a whole). Docstring rewritten to the working
  recipe (was previously "NOT reachable via LabTalk — PPT only", which was wrong).
- **Commit:** `5c538c1`
- **HOW VERIFIED:** fake tests (`tests/test_graph_layers.py`:
  `test_add_second_y_axis_colors_right_axis`,
  `test_add_second_y_axis_no_color_leaves_axis_black`,
  `test_add_second_y_axis_bad_color`) assert the exact `layer.y2.color` /
  `layer.y2.label.color` scripts are/are-not emitted. LIVE
  (`accept_phase2.py` → `probe_out/g2g8/accept_phase2.json`): right-axis strip
  near-red pixels **0 → 532**; message "colored right axis line + tick labels
  color(214,96,77)". PNGs `acc16_base.png` / `acc16_colored.png`.

### Task 2 — #17b: `colormap(levels=)` for continuous maps

- **What:** new optional `levels:int` on `colormap` + helper
  `_set_colormap_level_count_impl` (`tools/graph.py`). Emits
  `layer.cmap.numMinorLevels=0; numMajorLevels=<n>; setLevels(1);
  updateScale();` then reads `layer.cmap.numColors` back and RAISES on mismatch.
  Docstring records that no-arg `setLevels()` after `numColors=` is a silent
  no-op, and that a `palette` recolors even when the PNG byte-size is unchanged.
- **Commit:** `2e2a65d`
- **HOW VERIFIED:** fake tests (`test_colormap_levels_sets_and_verifies`,
  `test_colormap_levels_raises_on_readback_mismatch`,
  `test_colormap_levels_too_few`) assert the emitted LabTalk + the read-back
  raise. LIVE (`accept_phase2.json`): `layer.cmap.numColors` **8.0 → 32.0**;
  message "Set colormap to 32 continuous levels". PNGs `acc17_base.png` (8
  bands) / `acc17_lv32.png` (smooth).

### Task 3 — remove_plot per-index fix

- **What:** `remove_plot` (`tools/graph.py`) now removes ONLY the indexed plot
  via the COM `gl.DataPlots.Item(com_index).Destroy()` (mapping the 1-based
  DATA-plot index past any error-bar plots). The old `layer -e <dataset>`
  addressed by dataset NAME, so it removed EVERY plot of a duplicated dataset.
  Spike (`spike_rm2.json`) confirmed `Destroy()` is the only per-index route
  (DataPlot COM has no `Remove`/`Delete`; `range r; delete r` no-ops). Return
  message now lists the survivors. Fake `FakePlot.Destroy()` added to model it.
- **Commit:** `d5a1f14`
- **HOW VERIFIED:** fake tests (`test_remove_plot_destroys_only_indexed_duplicate`,
  `test_remove_plot_indexes_over_data_plots_skipping_errors`). LIVE
  (`accept_phase2.json`): built `[ARM_B, ARM_B, ARM_C]`, `remove_plot(1)` →
  `[ARM_B, ARM_C]` (only one B removed).

### Task 4 — #15 bar hatch (timeboxed theme attempt → hard limitation)

- **What:** timeboxed live attempt at the remaining untested route(s). The
  plot-object tree `layer.plotN.pattern.*` (`spike_rm_theme.json`) and indexed
  patterns `set -pfpd/-pfpi` (`spike_rm2.json`) BOTH no-op (returned true,
  byte-identical export, zero rendered hatch), on top of the phase-1 no-ops for
  `set -pfp/-pfw/-pfc` and the range tree. Documented #15 as a HARD LIMITATION
  (PPT/Illustrator workaround) in `publication-figure.md` (Origin COM Notes) and
  `README.md` (new "Hard limitations" table). No code exposed.
- **Commit:** `708c591` (docs)
- **HOW VERIFIED:** spike receipts `spike_rm_theme.json` / `spike_rm2.json`
  (`plot1_tree` all 0; `after_pfpd`/`after_pfpi` byte-identical to base).

### Task 5 — #19 / #21 / cross-cutting docs

- **What:** `publication-figure.md` + `README.md`: #19 transparency documented
  as a HARD LIMITATION (all `-paap/-paal/-paas` + `rr.transparency` no-op;
  unconfirmed future candidate = originpro `Plot.transparency()`); #21 layer-2
  `set -w` documented as deterministic-with-protocol (`page.active` + ~2s settle
  + `doc -uw`; saturates ~`w=8`, scaled-copy workaround only beyond that).
  `run_labtalk` docstring (`tools/labtalk.py`): the first 1-3 raw `set`/`layer.*`
  commands right after `create_graph` can silently no-op (settle race the typed
  tools guard against). `test_doc_registry.py` NOT_TOOLS allowlist extended
  (`setLevels`/`Destroy`/`transparency` are LabTalk/COM/originpro helpers).
- **Commit:** `708c591` (same docs commit)
- **HOW VERIFIED:** phase-1 receipts (`PROBE_G2G8_FINDINGS.md`, `probe_log.json`);
  `tests/test_doc_registry.py` green (3 passed) so the new doc tokens don't
  break the doc-vs-registry guard.

### Task 6 — harden the flaky top-tick live test

- **What:** `test_axis_tick_top_none_removes_marks_keeps_bottom_labels`
  (`tests/test_live_styling.py`) dropped the whole-image byte-diff
  (`baseline != after`) that cried wolf. On this graph state removing the
  top/right ticks does not move a pixel — the top-band dark-pixel count is
  IDENTICAL (measured `3056 == 3056`) — and a settle race could also leave the
  exports byte-identical. Now: a `doc -uw` + settle before the after-export;
  the bottom/left tick-LABEL strip must survive non-blank and ~unchanged (the
  real #4 regression guard, strict); the top band may only LOSE ink, never gain
  it (tolerant).
- **Commit:** `2601162`
- **HOW VERIFIED:** LIVE, run 3× standalone → 3/3 passed (`1 passed in ~8s`
  each). The prior round recorded this exact test as the "1 failed" in its full
  live-suite run.

### Non-negotiables checked

- WSL fake suite (`uv run --with pytest pytest -q`): **596 passed, 30 skipped**
  (0 failed) at final HEAD `2601162`. Baseline before this round: 596 passed.
- FULL live suite (Windows, `pytest -m requires_origin`, C:\ clone rsynced from
  this WSL tree): **30 passed, 596 deselected in 138s** — 0 failed, INCLUDING
  the task-6 test that was failing before. Log:
  `probe_out/g2g8/live_suite_final.txt`.
- `ruff check` on every source/test file touched this round: clean (the one
  remaining `F401 threading` is a PRE-EXISTING unused import in
  `tests/fakes.py`'s unrelated `ThreadGuardedFake`, present at base HEAD — not
  introduced here; the repo has 15 pre-existing ruff findings and no ruff config).
- Zero leftover Origin: baseline `0` `Origin.exe` before every run; `0` after
  every script and after both full live-suite runs (`tasklist` checked). All
  isolated `DispatchEx` instances `.Exit()`'d in finally. User Origin untouched.
- One commit per task — 5 implementation commits on base `97dfd5b`: `5c538c1`
  (#16), `2e2a65d` (#17b), `d5a1f14` (remove_plot), `708c591` (docs for tasks
  4+5, which are both documentation), `2601162` (task 6) — plus this WORK_LOG
  receipts commit.
- Did NOT push (per instructions).
