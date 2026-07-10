# Styling-Report (13 Issues) Fix Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix/close all 13 issues in `/home/swym4/01.Project/02.Tool_development/03.Figure_Skills/ORIGIN_MCP_ISSUES.md` (4 BLOCKER, 4 BUG, 5 PAPERCUT) reported against Origin-Pro-MCP v0.2.0 by a Windows agent reproducing a matplotlib publication figure.

**Architecture:** Empirical-probe-first. Every issue here is of the "LabTalk silently no-ops" class, so code reasoning alone is insufficient: Phase 0 runs a live probe matrix on real Origin 2020 (via Windows python from WSL interop) and records ground truth; Phase 1 implements fixes conditioned on those findings; every fix ships with a live pixel-verified test plus WSL fake tests.

**Tech Stack:** Python 3.10+, LabTalk over COM (win32com), pytest (`requires_origin` marker for live tests), dev repo in WSL at `/home/swym4/01.Project/02.Tool_development/02.Origin_MCP`, live runtime clone at `C:\Users\swym4\Origin-Pro-MCP`.

## Global Constraints

- Do NOT add new MCP tools — extend existing tools with optional params or docstrings only (standing user constraint: no tool/context bloat).
- Never touch `Origin.ApplicationSI` in tests/probes — always isolated `win32com.client.DispatchEx("Origin.Application")` (the `live_origin` fixture pattern in `tests/test_live_loaded_graph.py:25-55`). Always `origin.Exit()` in teardown and verify your spawned `Origin.exe` PIDs are gone; never kill Origin processes you did not spawn.
- Live commands run from WSL as: `cd /mnt/c/Users/swym4/Origin-Pro-MCP && /mnt/c/ProgramData/anaconda3/python.exe -m pytest -m requires_origin <file> -v` (WSL interop restored 2026-07-09; if "Exec format error" reappears: `sudo sh -c "echo ':WSLInterop:M::MZ::/init:PF' > /proc/sys/fs/binfmt_misc/register"`).
- To live-test UNCOMMITTED WSL changes: rsync the WSL worktree over the C: clone working tree (`rsync -a --delete --exclude .git /home/swym4/01.Project/02.Tool_development/02.Origin_MCP/src/ /mnt/c/Users/swym4/Origin-Pro-MCP/src/` plus the tests dir), run there, iterate. After the final WSL commit+push: `cd /mnt/c/Users/swym4/Origin-Pro-MCP && git fetch && git reset --hard origin/main`.
- WSL fake-based suite must stay green: `python3 -m pytest -q` (556+ tests) in the WSL repo. Fakes live in `tests/fakes.py`.
- Commits in the WSL repo only, one per task, message style matches `git log`; end with `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`. Do not push until all tasks pass.
- Windows may have user-owned Origin windows open. Probes/tests must not activate, modify, save, or close them.

## Issue → verdict map (from code inspection, pre-probe)

| # | Verdict | Root cause / key fact |
|---|---------|----------------------|
| 1 cap width | **Knob exists, not exposed.** | `set <err> -erwc <n>` already used & verified in `apply_publication_style` (style.py:420-423). Reporter used `-ew` (wrong flag). |
| 2 whisker width | **Knob exists, not exposed.** | `set <err> -erw <pt>` (POINTS, unlike `-w`) same place. Reporter used `-w`. |
| 3 frame width | **Probe needed.** | Code uses `layer.x.thickness` (style.py:399), reporter tried `layer.x.thick` (wrong name) — but reporter also says apply_publication_style didn't thicken, so `thickness` may itself be a no-op or mis-scaled. |
| 4 top/right ticks | **Probe needed.** | Reporter's `layer.x2.majorTicks=0` corrupted labels. Candidates: `layer.x2.ticks=0` (our tick impl uses `ticks` for direction 1/2/3), `layer.x2.majorLen=0`. |
| 5 `set -w` "no-op" | **Not a bug — units.** | `-w` is ~200 units/pt (style.py:43). Reporter's `-w 5` vs `-w 19` = 0.025 pt vs 0.095 pt — both hairline. Document. |
| 6 color wipe | **Real bug, cause identified.** | `set_plot_style` always sends `-k <auto-shape> -z 8 -kf 0 -w 500` even when the caller only wanted one aspect (defaults are values, not None), and never sends `-cf` for symbol plots — solid interiors fall back to black. |
| 7 pub-style corruption | **Partially by design + probe.** | Palette overwrite is by design (document). Nothing in the code groups plots or opens symbols — needs live repro. |
| 8 fill on grouped | **Expected Origin behavior, docs wrong.** | Group override; docstring claims "line_width/symbol still apply while grouped" — verify and correct. |
| 9 `-k` table mismatch | **Docs bug.** | `symbol_shape` passes raw `-k` code; docstring table (1=square, 4=diamond…) contradicts reporter's live observation (-k 4=down-triangle, -k 5=diamond). Probe the real table, fix docs. |
| 10 active-window | **Real footgun.** | `run_labtalk` executes globally; `layer.*`/`col()` hit whatever window is active. Add optional `window` param + docstring warning. |
| 11 unreadable props | **Document + translate sentinel.** | `-1.23456789e-300` is Origin's missing sentinel; translate in `capture` output, document which props are readable. |
| 12 autosave gate blocks delete | **Real bug.** | `save_in_place` returns False BOTH for "never-saved project — nothing on disk to protect" (autosave.py:216-217) and for real save failures; preflight (daemon.py:430) blocks on any False. Tri-state fix. |
| 13 WSL export paths | **Probe needed.** | S6 rejects POSIX paths by design. If Origin's expGraph can write to `\\wsl.localhost\<distro>\...`, translate; else improve the error hint. |

## Task 0 verdicts (COMPLETE — full detail in `/mnt/c/Users/swym4/probe_out/PROBE_FINDINGS.md`, PNGs beside it)

- **P1:** `layer.x/y/x2/y2.thickness` DOES control frame width (≈points; readable). Task 2's "no knob" ValueError fallback is NOT needed.
- **P2:** `layer.x2.majorTicks=0` confirmed corrupts number labels on ALL axes — never emit it; add a regression guard. `layer.x2.ticks=0` and `layer.x2.majorLen=0` are label-safe candidates, but tick-mark removal itself was too subtle at 1200px export — Task 2 must re-verify at larger export width and pick one.
- **P3:** the -k 1..12 table from pass 1 is INVALID (settle bug). Only `-k 3` = solid triangle-up re-confirmed. Task 5 must re-run the loop cleanly (single-flag `set <ds> -k <n>;`, 2s settle) before writing any docstring table.
- **P4:** on a GROUPED plot, `-kf`/`-w`/`-z` are ALL no-ops (extends issue #8). The style.py:168 claim "line_width/symbol still apply while grouped" is WRONG — correct it in Task 3. Also: a grouped multi-Y `plotxy` enumerates as ONE plot in `get_plot_info`.
- **P5:** issue #7 NOT reproduced but INCONCLUSIVE (precondition colors never applied due to settle bug). Task 3 must re-run the repro with the Task 0.5 fix in place; #7 may well be the reporter's own timing artifact.
- **P6:** `-erw`/`-erwc` CONFIRMED working; plot order confirmed adjacent `[data, err, data, err]` for both y_error_col and set_error_bars routes — Task 1's adjacency design holds.
- **P7:** INCONCLUSIVE — UNC `\\wsl.localhost\...` export didn't raise, but WSL-side file arrival was NOT verified. Task 6 must check `/tmp/probe_wsl.png` actually exists before implementing translation.
- **P8 (root cause of #6):** combining `-c` and `-cf` in ONE `set` command wipes the plot to BLACK (either order); two SEPARATE `set` calls work perfectly. Also `-k`+`-kf`+`-z` combined can blank the symbol. **Rule for Task 1: every flag gets its own `set <ds> -flag val;` call — never batch flags in one string.** The style.py:212-214 batching "optimization" is the bug.
- **P9:** primary `layer.x/y.*` props (from/to/inc/thickness/ticks/majorLen) read back real values; the `-1.23456789e-300` missing-sentinel is mainly an x2/y2 (opposite-axis) phenomenon — scope Task 5's docs accordingly.
- **Cross-cutting (NEW ISSUE → Task 0.5):** a freshly created graph page silently ignores or partially applies LabTalk styling/reads/exports issued too soon after `create_graph`/`plotxy`/`add_plot_to_graph` — no exception, ~2s settle fixes it. Distinct from the loaded-.opju freeze. Must be fixed in the codebase, not left to callers.

---

### Task 0: Live probe matrix (ground truth before any code)

**Files:**
- Create: `scratch/probe_styling.py` in the session scratchpad (NOT the repo) — run against the C: clone's installed v0.2.0 (`PYTHONPATH=C:\Users\swym4\Origin-Pro-MCP\src`).
- Produce: `scratch/PROBE_FINDINGS.md` + exported PNGs.

**Interfaces:**
- Produces: PROBE_FINDINGS.md with one section per probe P1–P9, each ending in a single-line verdict `P<n>: <answer>`.

- [ ] **Step 1: Write the probe script** — standalone (no pytest), mirroring the `live_origin` fixture: `pythoncom.CoInitialize()`, `DispatchEx("Origin.Application")`, `set_session_origin(origin, factory=...)`, `try/finally: origin.Exit(); clear_session_origin()`. Build the reporter's repro graph once per probe group: workbook via `create_worksheet`, 4 columns X/Y/Yerr, `create_graph(..., plot_type="line+symbol", y_error_col=3)` + `add_plot_to_graph` for a 2nd series. Export via `export_graph(name, r"C:\Users\swym4\probe_<id>.png")` and compare bytes/pixels between variants. Probes:
  - **P1 frame width:** export baseline; `layer.x.thickness=8; layer.y.thickness=8; layer.x2.thickness=8; layer.y2.thickness=8` (via `graph_layer_execute`); export; also try value 0.5. Verdict: does `thickness` change rendered frame width, what unit scale, and does the `opposite`-side line follow `layer.x.thickness` or need `layer.x2.thickness`?
  - **P2 per-side ticks:** on a closed-frame graph with visible tick labels, try in order (re-export + check labels after each, restoring between): (a) `layer.x2.ticks=0; layer.y2.ticks=0`, (b) `layer.x2.majorLen=0; layer.y2.majorLen=0` (+minor), (c) reporter's `layer.x2.majorTicks=0` to reproduce the label corruption. Verdict: which removes top/right tick MARKS while keeping bottom/left NUMBER LABELS?
  - **P3 symbol table:** one scatter plot; loop `set <ds> -k <n>` for n=1..12, export each, then LOOK at the PNGs (Read tool renders images) and catalog shape + open/solid per code. Verdict: the real `-k` table for Origin 2020.
  - **P4 fill vs grouping:** build a GROUPED plot (single `plotxy` with two Y ranges, `plot:=202`), try `set <ds> -kf 0` and `-kf 1` per curve, export. Verdict: does `-kf` apply while grouped? Do `-w`/`-z` apply while grouped?
  - **P5 pub-style repro (#7):** 4-series line+symbol with error cols, apply custom colors via `set -c/-cf`, then `apply_publication_style(...)`, export. Check: symbols open or solid? shapes per-plot or uniform? Then try `set_plot_style(open_symbol=False)` on plot 1 and check refill. Verdict: exactly which corruption reproduces and what LabTalk state causes it (inspect with `get <ds> -k`/`-kf` reads).
  - **P6 error-bar knobs (#1/#2):** on the y_error_col-built graph: `set <errds> -erw 2.5;` export; `-erw 5` export; `-erwc 8` vs `-erwc 20` export. Pixel-verify each changes whiskers/caps; note the unit scale (reporter wants caps ≈60% of symbol width). Also confirm `get_plot_info` plot ordering: is it `[data, err, data, err]` (adjacent) for the y_error_col build AND after `set_error_bars`?
  - **P7 WSL export (#13):** `export_graph(name, r"\\wsl.localhost\Ubuntu\tmp\probe_wsl.png")` — does the file appear in WSL `/tmp/probe_wsl.png`?
  - **P8 color wipe repro (#6):** color plot 1 red via `set <ds> -c color(255,0,0) -cf color(255,0,0)`, export; call current `set_plot_style(g, 1, line_width=5.0)` (defaults active), export. Did the curve turn black? Then isolate: send ONLY `set <ds> -w 1000`, export — color preserved? Then `-kf 0` alone — is it the fill reset? Verdict: the minimal flag set that wipes color.
  - **P9 `layer.*` reads (#11):** `capture` reads of `layer.x.from/to/inc/thickness/ticks/majorLen` on a live graph — which return real values vs 0 vs -1.23456789e-300.
- [ ] **Step 2: Run it** (`/mnt/c/ProgramData/anaconda3/python.exe scratch/probe_styling.py`), write PROBE_FINDINGS.md with verdicts + PNG paths. Confirm zero leftover probe-spawned Origin.exe processes.

---

### Task 0.5: New-page settle barrier (cross-cutting fix, prerequisite for all live tests)

**Files:**
- Modify: `src/origin_pro_mcp/tools/graph.py` (`create_graph`, `add_plot_to_graph`), `src/origin_pro_mcp/tools/style_helpers.py` (shared helper), possibly `src/origin_pro_mcp/tools/style.py` (`ungroup_plots` rebuild loop)
- Test: fake tests (helper called after plotting) + a live regression test

**Interfaces:**
- Produces: `settle_new_plots(graph_name: str, expected_min_plots: int, timeout_s: float = 4.0) -> None` in style_helpers — polls `get_plot_info` until at least N plots enumerate, then applies a short fixed settle; called by create_graph/add_plot_to_graph/ungroup_plots after their plotxy.

- [ ] **Step 1:** Live failing repro first (this is the probe's exact failure): `create_graph(...)` then IMMEDIATELY `set <ds> -c color(255,0,0);` via graph_layer_execute, export, assert red pixels present (Pillow count > 1000). Without the fix this intermittently/reliably fails (probe evidence: p5_before_pubstyle.png). Calibrate the minimal reliable barrier (poll + settle; prefer poll-until-enumerated over blind sleep, add the smallest fixed tail that makes the repro pass 3/3 runs).
- [ ] **Step 2:** Implement the helper + call sites. Keep it cheap: skip the fixed tail when the poll already found the plots on its first try (page was already settled).
- [ ] **Step 3:** WSL fake suite green (fakes enumerate instantly, so the helper must no-op fast against fakes). Live repro passes 3/3. **Commit** (`fix: settle barrier after page creation — new plots silently ignored immediate styling`).

---

### Task 1: `set_plot_style` — None-defaults + `-cf` + error-bar params (#1 #2 #6)

**Files:**
- Modify: `src/origin_pro_mcp/tools/style.py:136-246`
- Test: `tests/test_style_tools.py` (existing fake tests — update expectations), new live test in `tests/test_live_styling.py`

**Interfaces:**
- Produces (MCP-visible signature): `set_plot_style(graph_name, plot_index=1, line_width: float|None=None, symbol_size: int|None=None, symbol_shape: int|None=None, color="", rgb="", open_symbol: bool|None=None, error_bar_width: float|None=None, error_cap_width: float|None=None)` — **None/"" = leave that aspect untouched** (breaking change from value-defaults, intentional: partial styling must not reset other aspects).

**P8 HARD RULE (probe-verified):** every LabTalk flag is sent as its OWN `set <ds> -flag val;` call. Combining `-c` + `-cf` in one command wipes the plot to black; `-k`+`-kf`+`-z` combined can blank the symbol. Delete the batching code and its style.py:212-214 rationale comment — that "optimization" is the root cause of issue #6. One short settle (`time.sleep(0.2)`) after the LAST call is enough; do not sleep per flag.

- [ ] **Step 1: Write failing fake tests**: (a) `set_plot_style(g, 1, line_width=5)` must emit ONLY `set <ds> -w 1000;` (no `-k`, no `-z`, no `-kf`); (b) `rgb="255,0,0"` on a symbol plot must emit `set <ds> -c color(255,0,0);` and `set <ds> -cf color(255,0,0);` as TWO SEPARATE Execute strings — assert no single executed string contains both `-c ` and `-cf `; (c) no aspect given → ValueError "nothing to change"; (d) `error_bar_width=2.5, error_cap_width=12` must emit `set <errds> -erw 2.5;` and `set <errds> -erwc 12;` (separate calls, same P8 rule) against the error plot(s) immediately following the chosen data plot in `get_plot_info` order (P6 confirmed adjacency for both construction routes) — fall back to ALL error plots on the layer with a note in the return message if none are adjacent; (e) `error_*` given but graph has no error plots → ValueError naming set_error_bars/y_error_col.
- [ ] **Step 2: Run tests, verify they fail.**
- [ ] **Step 3: Implement.** Core shape (each entry becomes its own `graph_layer_execute(g, f"set {pname} {flag};")` call):

```python
flags: list[str] = []
if c is not None:
    flags.append(f"-c {c}")
if line_width is not None:
    flags.append(f"-w {int(line_width * _WIDTH_UNITS_PER_POINT)}")
touch_symbols = (symbol_shape is not None or symbol_size is not None
                 or open_symbol is not None)
has_symbols = _plot_has_symbols(pname)
if touch_symbols and has_symbols:
    if symbol_shape is not None:
        shape = symbol_shape if symbol_shape > 0 else SYMBOL_SHAPES.get(plot_index, 2)
        flags.append(f"-k {shape}")
    if symbol_size is not None:
        flags.append(f"-z {symbol_size}")
    if open_symbol is not None:
        flags.append(f"-kf {1 if open_symbol else 0}")
if c is not None:
    flags.append(f"-cf {c}")  # symbol interior / bar fill follows the curve color
for flag in flags:
    _set(flag)                # ONE flag per LabTalk `set` call — P8 hard rule
time.sleep(0.2)
```
plus the error-bar block (adjacency via `infos`; separate `-erw` / `-erwc` calls), the empty-call guard, and a docstring rewrite: state None-semantics, the `-erw`(points)/`-erwc` units from P6, and the corrected grouping caveat from P4 (while grouped, color AND width AND symbol changes are ALL overridden — ungroup_plots first).
- [ ] **Step 4: WSL suite green** (`python3 -m pytest -q`).
- [ ] **Step 5: Live test** (new `tests/test_live_styling.py`, `requires_origin`, `live_origin` fixture copied from test_live_loaded_graph.py): build 2-series line+symbol with y_error_col; color plot 1 red; `set_plot_style(g, 1, line_width=6.0)`; export before/after; assert pixels changed AND red pixels still present (Pillow: count px with r>200,g<80,b<80). Second test: `set_plot_style(g, 1, error_bar_width=4.0, error_cap_width=16)` changes pixels vs baseline. Run live; iterate until green.
- [ ] **Step 6: Commit** (`fix: set_plot_style partial styling + error-bar width/cap exposure`).

---

### Task 2: Frame line width + per-side tick removal (#3 #4)

**Files:**
- Modify: `src/origin_pro_mcp/tools/graph.py` (`axis` op="frame" → accept `frame_width`; op="tick" → accept axis values `"top"`/`"right"` mapping to `layer.x2`/`layer.y2`, and `tick_direction="none"`), `src/origin_pro_mcp/tools/style.py` (`_set_tick_style_impl` gains the x2/y2 axes + `none` direction; `apply_publication_style` frame thickness uses the P1-verified knob/scale)
- Test: fake tests + live pixel tests in `tests/test_live_styling.py`

**Interfaces:**
- Produces: `axis(g, op="frame", frame="closed", frame_width=2.0)` — frame_width in points (or P1's verified unit, documented); `axis(g, op="tick", axis="top"|"right", tick_direction="none")` removes that side's tick marks only.

**Probe outcomes to build on:** P1 confirmed `layer.<ax>.thickness` works and is readable (≈points) — frame_width is a straight implementation, no fallback needed. P2: candidates are `layer.x2.ticks=0` or `layer.x2.majorLen=0` (both label-safe); NEVER emit `majorTicks` — probe-confirmed it wipes number labels on all four axes (see `p2c_after_majorTicks0.png`). Add a fake test asserting the emitted LabTalk never contains `majorTicks` for tick ops.

- [ ] **Step 1:** Decide the P2 knob first with a 5-minute live check at LARGE export width (`export_graph` then crop the top-frame strip; probe's 1200px exports were too small to see which candidate removes the marks): pick whichever of `ticks=0` / `majorLen=0` (+`minor=0`) visibly removes top/right tick marks while the bottom-label strip stays intact. Record the choice + PNG evidence.
- [ ] **Step 2:** Fake tests for the new param plumbing (exact LabTalk per Step 1's choice), including the no-`majorTicks` guard. Run, verify fail.
- [ ] **Step 3:** Implement (`frame_width` on op="frame" applying `layer.x/x2/y/y2.thickness`; axis="top"/"right" + tick_direction="none" on op="tick").
- [ ] **Step 4:** WSL suite green.
- [ ] **Step 5:** Live tests: (a) frame_width 0.5 vs 6 exports differ (and read-back `layer.x.thickness` equals what was set — P1 confirmed readable); (b) `axis(op="tick", axis="top", tick_direction="none")` → export shows change AND bottom labels still render (crop the bottom-label strip, assert non-blank / unchanged vs baseline).
- [ ] **Step 6: Commit** (`feat: axis frame width + per-side tick control`).

---

### Task 3: `apply_publication_style` corruption (#7) + grouped-fill truth (#8)

**Files:**
- Modify: `src/origin_pro_mcp/tools/style.py` (`apply_publication_style`, docstrings of it and `set_plot_style`/`ungroup_plots`)
- Test: live repro test per P5, fake tests for any code change

**Probe outcomes to build on:** P5 was inconclusive (its custom-color precondition silently failed — the Task 0.5 settle bug); apply_publication_style itself worked correctly on a clean ungrouped build (frame, ticks, legend all fine, symbols stayed solid). P4 CONFIRMED that on a grouped plot `-kf`/`-w`/`-z` are all no-ops — the style.py:168 claim "line_width/symbol still apply while grouped" is wrong and must be corrected regardless of the P5 rerun's outcome. Note: apply_publication_style still batches multi-flag `set` commands (style.py:432-441) — apply the P8 one-flag-per-call rule here too; per P8 that batching may itself be what produced the reporter's "colors reset" in issue #7.

- [ ] **Step 1:** Re-run the P5 repro CLEANLY (with Task 0.5's settle fix in place): 4-series line+symbol + error cols, custom colors via separate `-c`/`-cf` calls, verify colors actually applied (red-pixel count), then apply_publication_style, then `set_plot_style(open_symbol=False)` refill. If corruption reproduces, fix the actual mechanism; if not, issue #7 was the reporter's own timing/flag-batching artifact — then this task is the P8-rule refactor of apply_publication_style's `set` calls + docstring truth: "applies the pastel palette (overwrites custom colors) — do custom colors AFTER this call", and correct style.py:168 per P4 (while grouped, color/width/symbol/fill changes are ALL overridden).
- [ ] **Step 2:** Live test: reporter's minimal repro (4-series + error bars + custom colors → apply_publication_style → set_plot_style(open_symbol=False) refill works; shapes stay per-plot). Export-and-inspect assertions.
- [ ] **Step 3:** WSL suite green. **Commit** (`fix: apply_publication_style symbol/grouping integrity`).

---

### Task 4: Autosave tri-state — stop blocking deletes on never-saved projects (#12)

**Files:**
- Modify: `src/origin_pro_mcp/autosave.py` (`save_in_place`), `src/origin_pro_mcp/daemon.py:402-434` (`_autosave_preflight`)
- Test: `tests/test_autosave.py` (existing fake tests file — extend)

**Interfaces:**
- Produces: `save_in_place(origin, remembered_path) -> bool | None` — `True` saved, `None` = "nothing on disk to protect" (never-saved project, or empty project), `False` = a real save attempt failed. Preflight: `None` → proceed (never block), `False` + required → block.

- [ ] **Step 1: Failing fake test:** never-saved project (no remembered path, `%X`/`%G` empty or file absent) + `delete_graph` dispatch → op PROCEEDS (no "Autosave before … failed" error). Existing behavior preserved: real save failure (path exists, `origin.Save` returns False) still blocks.
- [ ] **Step 2:** Run, verify fail.
- [ ] **Step 3:** Implement: in `save_in_place`, `pages <= 0` → `None`; `not path or not os.path.isfile(path)` → `None`; `origin.Save` exception/False → `False`. In `_autosave_preflight`: `ok = _autosave.save_in_place(...)`; `if ok or ok is None: return None`. Extend the block message: "…(set ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED=0 to proceed without saving, or ORIGIN_PRO_MCP_AUTOSAVE=off if this session must never write the project)".
- [ ] **Step 4:** WSL suite green. **Commit** (`fix: autosave gate no longer blocks destructive ops on never-saved projects`).

---

### Task 5: Documentation + small footguns (#5 #9 #10 #11)

**Files:**
- Modify: `src/origin_pro_mcp/tools/labtalk.py` (`run_labtalk`: new optional `window: str = ""` param + docstring; sentinel translation in capture), `src/origin_pro_mcp/tools/style.py` (symbol table docstrings), `src/origin_pro_mcp/skills/publication-figure.md`, `README.md`

- [ ] **Step 1: `run_labtalk` window param:** when `window` given → `labtalk_name` validate, `activate_window(safe_window, "window")` before executing (reuses the COM activation that works on loaded graphs, unlike bare `win -a`). Docstring adds: "GOTCHA: `layer.*`, `col()`, `%C` target the ACTIVE window — pass `window=<graph>` or prefer the typed tools (they take graph_name)."
- [ ] **Step 2: capture sentinel:** in the capture loop, numeric values `< -1e299` → return the string `"missing"` (documented: Origin's sentinel for unset/unreadable properties — P9 showed it mainly on x2/y2 opposite-axis reads; primary layer.x/y props read back real values). Fake test.
- [ ] **Step 3: Re-run the symbol table probe cleanly** (P3's original 1-12 loop was invalidated by the settle bug): with Task 0.5's fix, loop single-flag `set <ds> -k <n>;` for n=1..12 on a line+symbol build, export each, LOOK at the PNGs, and record the real Origin 2020 table (only k=3=solid-triangle-up is currently confirmed). Do NOT copy any table from the old docstring, the reporter, or memory.
- [ ] **Step 4: Units + tables docs** (from P3-rerun/P6/P9): run_labtalk docstring gets a compact block — `set -w` ≈200 units/pt (500=2.5pt); error bars use `-erw <points>` / `-erwc <cap>` NOT `-ew`/`-w`; NEVER combine multiple `set` flags in one command (P8 — silently corrupts color/symbols); NEVER `layer.x2.majorTicks=0` (wipes all labels); the verified `-k` symbol table; which `layer.*` props are readable (primary axes yes, x2/y2 often sentinel). Fix the `set_plot_style`/`SYMBOL_SHAPES` docstring table to match the P3 rerun. Mirror the essentials in publication-figure.md ("error-bar width/cap → set_plot_style(error_bar_width=, error_cap_width=)"; "never style via raw `set -w` — units differ"; per-side ticks; frame width) and a short README "LabTalk gotchas" subsection.
- [ ] **Step 4:** WSL suite green. **Commit** (`docs+fix: run_labtalk window param, sentinel translation, LabTalk unit/symbol tables`).

---

### Task 6: WSL export paths (#13, conditional on P7)

**Files:**
- Modify: `src/origin_pro_mcp/labtalk_safe.py` (`windows_path`)
- Test: `tests/test_tool_guards.py`

- [ ] **Step 0:** Settle P7 first (probe left it inconclusive — export "did not raise" but WSL-side arrival was never checked): live-export to `\\wsl.localhost\Ubuntu\tmp\p7_check.png`, then verify `/tmp/p7_check.png` exists AND is a valid PNG from the WSL side.
- [ ] **Step 1:** If Step 0 shows Origin CAN write `\\wsl.localhost\...`: translate bare-POSIX paths to `\\wsl.localhost\<distro>\<path>` when a distro name is known (env `ORIGIN_PRO_MCP_WSL_DISTRO`, else `WSL_DISTRO_NAME` if the daemon inherited it); otherwise keep the rejection but append: "…or set ORIGIN_PRO_MCP_WSL_DISTRO=<distro> to auto-map to \\\\wsl.localhost". If P7 shows Origin cannot write there, docstring/README note only ("export to /mnt/<drive>/... then copy") — the current clear error already beats a fake success. Fake tests for whichever branch; live test if translating.
- [ ] **Step 2:** WSL suite green. **Commit.**

---

### Task 7: Finish — full verification, deploy, report

- [ ] **Step 1:** Full WSL suite `python3 -m pytest -q` green; `ruff check src/` clean on changed files.
- [ ] **Step 2:** Full live suite on C: clone (after rsync): `pytest -m requires_origin tests/ -v` — all green; confirm no leftover probe/test Origin.exe processes.
- [ ] **Step 3:** Bump `pyproject.toml` version to `0.2.1` (cache-bust for uvx configs). Commit.
- [ ] **Step 4:** STOP — do NOT push. The orchestrator runs an independent fresh-eyes review first; push + clone sync (`git reset --hard origin/main` in `/mnt/c/Users/swym4/Origin-Pro-MCP` and `"/mnt/g/My Drive/VDLab_Google/14.Agent/01.Projects/06.MCP_development/Origin_Pro_MCP/repo"`) happen after review approval.
- [ ] **Step 5:** Final report: per-issue table (fixed / documented-limitation / not-reproduced) with evidence paths (live test names, exported PNG paths, commit hashes).
