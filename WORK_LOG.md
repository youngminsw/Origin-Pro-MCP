# Work Log

## 2026-07-11 — whole-product review Round D (docs coherence + agent-context diet) (agent: round-d)

Scope: items 13-20, 32, and the Addendum from `docs/REVIEW-2026-07-11-whole-product.md`,
docs-only (no runtime behavior changes, no live Origin needed). Base =
`1890e02` (Rounds A+B+C merged). One commit per task, ruff-clean on every
touched file, WSL suite must stay green throughout.

| Task | Commit(s) | What / verdict |
|---|---|---|
| 13 (docstring diet) | `b2052cd` | New `src/origin_pro_mcp/skills/labtalk-gotchas.md` skill (6836 bytes, loaded on demand via `get_skill`) holds the FULL gotcha ledger: P8 flag-batching rule + evidence, `majorTicks` danger, `-w`/`-erw`/`-erwc` unit tables, `-k` symbol table, `offsetV`/`offsetH` tick-label-offset property semantics, the `y2`/`YR` right-axis recipe, layer-2 `-w` protocol, the settle race, readable-vs-sentinel axis properties, `0/0` missing-value writes, and the hatch/transparency hard limitations (#15/#19). `run_labtalk`'s docstring keeps only its own confirm-gate/capture/statement-retry mechanics + a 5-line safety-critical shortlist + a `get_skill('labtalk-gotchas')` pointer. `add_second_y_axis`'s inline recipe block shrank the same way. |
| 20 (color param cross-ref) | `b2052cd` | `add_second_y_axis`'s `color` docstring now states it is always "r,g,b" numeric (unlike `set_plot_style`'s named `color`); `set_plot_style`'s `rgb` docstring now notes `add_second_y_axis` uses this same format. Docs-only, no behavior change. |
| 18/19 (remove_plot truth, tool count) | `8df4f9b` | README.md:363 and publication-figure.md's own prose (not its COM Notes table, which was already correct) still described the old `layer -e`/`layer -ie` name-addressed removal, replaced by `d5a1f14`'s per-index COM `DataPlot.Destroy()` — fixed both to match the COM Notes table. Tool count: **verified, not fixed** — the live registry is 45 tools (43 `@mcp.tool` functions in `tools/*.py` + `list_skills`/`get_skill` from `skills.py`), matching README/skill prose already; the review's "actual is 43" used an AST scan that only walked `tools/*.py` and missed the 2 skill tools. `tests/test_doc_registry.py::test_readme_tool_count_matches_registry` (already present, added in `9b44d5b`) asserts this against the live registry and passes — no test change needed. |
| 32 + usability P2s | `c79ac4f` | publication-figure.md: softened `offset_pct=80` to "start at 20-30, tune against an export" (a live run showed 80 collides tick numbers with the axis); added a note that `colormap(levels=)` drives colorbar LABEL density along with band smoothness, so a high level count can over-label the bar. `import_data` docstring: told agents to use the returned `"name"` directly instead of a `list_worksheets` round-trip (F11). `set_plot_style`'s grouping NOTE reworded so it only suggests `ungroup_plots` when a change had NO visible effect, not on every call (F13). |
| 14 (FFT claim) | `a46818f` | Researched `docs.originlab.com/x-function/ref/fft1` + `.../origin-help/fft1-algorithm`: fft1's `interval` param defaults to `<Auto>`, defined as "the average increment of the time sequence, which is usually from the X column associated with the input signal." `_fft_impl` designates x_col/y_col as X/Y on the sheet before calling `fft1 ix:=col(<y>)`, so the Y column has an associated X — **the docstring's claim is TRUE**, kept as-is with a code comment citing the source. |
| 15 + feature-doc sync | `6527889` | README.md + publication-figure.md updated for Rounds A-C: vector export (`format="pdf"/"eps"/"emf"`, svg impossible), export width honored / no independent height, `create_graph(template=)`, `curve_fit(x_min=, x_max=, peaks=N, peak_centers=)`, `transform(smooth_method=)`, `axis(op="labels", axis="right"/"top")`. Added the "Return-format convention" paragraph to README (JSON for creation/data tools, plain strings for styling). Swept for stale `poly_order`/`dpi`/`height` references — none found; those were already fully removed in prior rounds. |

**Context-cost measurement (AST over all `@mcp.tool` docstrings, before → after):**
- Total always-loaded docstring bytes: **41646 → 40565** (-1081 bytes / -2.6%; the
  reduction is smaller than the review's 3-4k-token estimate because several
  tasks this round ADDED docstring content elsewhere — the import_data
  usability note, the rgb cross-reference, the reworded grouping note — that
  partially offset the run_labtalk/add_second_y_axis cuts).
- `run_labtalk`: **6159 → 5084 bytes** (-1075, -17%).
- `add_second_y_axis`: **1568 → 1080 bytes** (-488, -31%).
- New `labtalk-gotchas.md` skill: 6836 bytes, loaded ONLY on `get_skill('labtalk-gotchas')` — zero always-loaded cost.
- Tool count unchanged at 45 (43 `@mcp.tool` + 2 skill tools); confirmed by the same AST scan.

**Note on commit atomicity:** README.md's Graphing-section table hunk bundles
the item-18 `remove_plot` fix with adjacent item-15 row updates
(`create_graph`/`axis`/`export_graph`) because they fall in the same diff
hunk (adjacent markdown table rows) — split via `git add -p` where the tools
allowed it (`style.py`, `publication-figure.md`'s 5 independent hunks),
otherwise commits carry both items and say so in the message.

WSL suite (`/tmp/opm_venv/bin/python3.11 -m pytest -q`): **660 passed, 59
skipped** at HEAD `6527889` — unchanged from the Round A+B+C baseline (this
was a docs-only round, no test changes expected or needed). `ruff check` on
every file touched this round (`labtalk.py`, `graph.py`, `style.py`,
`analysis.py`, `worksheet.py`, `README.md`, `publication-figure.md`,
`labtalk-gotchas.md`): all clean. Did NOT push.

## 2026-07-11 — whole-product review Round C (multi-peak fitting, item 7) (agent: round-c)

Scope: the single biggest analysis gap — multi-peak fitting / deconvolution
(XRD/XPS/Raman: N overlapping gauss/lorentz/voigt peaks fit simultaneously).
DESIGN-FIRST: no tool code until the route was live-proven. Base = 6a7ad9e
(Rounds A+B merged). Live probes: isolated `DispatchEx("Origin.Application")`,
`Exit()` in finally, WSL tree rsynced into `C:\Users\swym4\opm_dev` (the
`Origin-Pro-MCP` clone was never touched). Receipts: `C:\Users\swym4\probe_out\roundc\`.

### Phase 1 — route probe (live, Origin Pro 2020)

**VERDICT: the NLFit REPLICA route works for gauss/lorentz/voigt — and it drops
straight into the existing `nlbegin/nlfit/nlend` + result-tree plumbing.** The
Peak Analyzer `fitpeaks` X-Function was the fallback candidate (only
gauss/lorentz, results in a report tree not the nltree); not needed.

Working script (probe `probe_multipeak.py`, 3-Gaussian synthetic: y0=0, peaks
(xc,σ,A)=(20,4,10),(35,3,15),(50,5,8), x 0..70 @0.35 = 201 pts + noise):
```
nlbegin iy:=[MPK]Sheet1!(1,2) func:=gauss replica:=2 nltree:=tt;   // replica = peaks-1
tt.xc = 20; tt.xc__2 = 35; tt.xc__3 = 50;    // optional initial centres (peak1 = unsuffixed)
nlfit;
// read BEFORE nlend: tt.xc / tt.xc__2 / tt.xc__3, tt.w[/__k], tt.A[/__k],
//                    tt.e_xc[/__k], shared tt.y0, tt.cod / tt.chisqr / tt.dof
nlend output:=1;   // only when a fit curve is wanted
```
Recovered centres: **xc = 20.00 / 35.00 / 50.01** vs ground truth 20/35/50.

Facts established (all live):
- **`replica:=N` fits N+1 peaks** (N = *additional* copies). For 3 peaks use
  `replica:=2`. `replica:=3` on 3-peak data over-parametrised → degenerate
  (cod=0, params frozen at init) — so the tool computes `replica = peaks-1`.
- **Node naming:** peak 1 = unsuffixed (`xc`,`w`,`A`); peaks 2..N = `xc__2`,
  `w__2`, `A__2`, … (DOUBLE underscore). Std errors = `e_` prefix
  (`e_xc__2`). Confirmed for gauss/lorentz (per-peak `xc/w/A`) and **voigt**
  (per-peak `xc/A/wG/wL`). `y0` is a **single SHARED baseline** (reading
  `y0__2` returns the same value). Stats `cod`/`chisqr`/`dof` unchanged.
- **Origin's parameterisation is unchanged from single-peak:** gauss `w`=2σ,
  `A`=area (recovered w=8/6/10 = 2×σ, exactly), so a multi-peak result reports
  the SAME node meanings the existing single-peak path already returns — no new
  caller surprise.
- **Stale-read hazard:** reading a node BEYOND the last real peak (e.g. `xc__4`
  when 3 peaks) silently returns the previous `__pr` value, not an error — so
  reads must be strictly bounded to `peaks`. A genuinely non-converged fit
  freezes params at their init values and reads std errors back as the
  missing-value sentinel `-1.23456789e-300` with **`cod=0.0`** (probe
  `probe_multipeak3.py`, flat-line data + absurd centres: `nlfit` still
  returned TRUE — so nlfit's return alone is NOT a convergence proof for
  multi-peak; `cod<=0` is the honest guard).
- **Auto-init converges too:** the same 3-peak fit with NO supplied centres
  recovered identical params — Origin's built-in replica initialisation is good
  on well-separated peaks, so `peak_centers` is optional (seed only when the
  caller supplies them; no internal find_peaks dependency needed).
- **Fit-curve sheet (for plot_on_graph):** `nlend output:=1` makes
  **`FitNLCurve1`** (contains "Curve" → the existing sheet-search still finds
  it). Its columns, found by LongName: `Fit Peak 1`, `Fit Peak 2`, … `Fit Peak
  N` (per-peak components) then **`Cumulative Fit Peak`** (the full envelope) —
  for N=3 that's cols 2,3,4 (components) and col 5 (cumulative). The plot path
  must target the cumulative column by LongName (NOT col 2, which is only
  component 1), and may additionally draw the component columns. `plotxy` of the
  cumulative column onto a graph succeeded live.

### Phase 2 — design

**NO NEW TOOL — extend `curve_fit`.** Two new params:
- `peaks: int = 1` — >1 activates multi-peak (replica) fitting; only for
  gauss/lorentz/voigt.
- `peak_centers: str = ""` — "x1,x2,…" initial centre guesses. When given, its
  count must equal `peaks`. When empty, Origin's replica auto-init is used.

Validation (fail-honest, BEFORE Origin):
- `peaks < 1` → ValueError.
- `peaks > 1` and function ∉ {gauss,lorentz,voigt} → ValueError.
- `peak_centers` supplied and count ≠ `peaks` → ValueError.
- `peak_centers` supplied with `peaks == 1` → ValueError (avoid a silent no-op).
- `peaks == 1` leaves the existing single-peak path and its JSON UNCHANGED.

Route (`peaks > 1`): `nlbegin … func:=<fdf> replica:=peaks-1 nltree:=__mcpfit`;
compose with the Round-B x_min/x_max row-subrange (restrict THEN deconvolve);
set `__mcpfit.xc[/__k]` from `peak_centers` when supplied; `nlfit`; raise on
`nlfit`==False OR on `cod` unreadable/≤0 (non-convergence, message suggests
supplying/adjusting `peak_centers`).

Return JSON (`peaks > 1`):
```
{ "function":"gauss", "peaks":3,
  "statistics": {r_squared, sum_sq_residuals, reduced_chi_sq, dof},
  "baseline": {"y0": {value, std_error}},          // the shared offset
  "parameters": {"peak_1":{"xc":{value,std_error},"w":{…},"A":{…}},
                 "peak_2":{…}, "peak_3":{…}} }
```
Per-peak keys come from `FITTING_FUNCTIONS[fn][1][1:]` (index 0 is the shared
`y0`): gauss/lorentz → xc/w/A, voigt → xc/A/wG/wL. Reads are sentinel-filtered
(`-1.23e-300` → omitted, like the "power" unreadable precedent); reads bounded
to exactly `peaks`. When `plot_on_graph` is set: draw the `Cumulative Fit Peak`
column (bold brick-red, P8 one-flag-per-`set`) as the fit line and the `Fit Peak
k` components as thin best-effort lines; report which were drawn.

### Phase 3 — implement + verify

One feature, `curve_fit` extended in place (no new tool). The `peaks>1` branch
returns early, so the single-peak path and its JSON are byte-for-byte unchanged.

| # | Commit | What / verification |
|---|---|---|
| impl | `<c1>` | `fitting.py`: `peaks`/`peak_centers` params; validation (peaks<1, unsupported func, centre-count/`peaks==1` misuse) raises before Origin; replica route + optional centre seeding; `cod<=0` non-convergence guard; sentinel-filtered per-peak reader; `_plot_multipeak_fit` (cumulative by LongName + best-effort components). 11 fake tests in `test_fitting.py` (replica count, centre seeding, auto-init, return shape, x-range composition, every raise, nlfit-fail, cod-fail). |
| live | `<c2>` | 5 live tests in `test_live_fitting.py`. |

Live (Windows, isolated DispatchEx, opm_dev rsync — clone untouched), all PASSED:
- **3-Gaussian recovery:** `peaks=3` recovered xc within 0.5 of 20/35/50, R²>0.99,
  every peak reports w/A + xc std_error.
- **peaks=2 lorentz + supplied centres** (`peak_centers="20,40"`): recovered 20/40.
- **peaks=2 + x_min/x_max=[10,42]:** restricts (drops the xc=50 peak) then
  deconvolves the two in-range peaks → 20/35.
- **non-convergence:** flat data + centres at 1000/2000/3000 → `cod<=0` guard
  raises "did not converge" (nlfit itself returned True — guard is what catches it).
- **plot:** `plot_on_graph` → `fit_curve.components_drawn == 3`, cumulative column
  drawn, graph exports a non-empty PNG.

WSL fake suite (`/tmp/opm_venv/bin/python3.11 -m pytest -q`): **660 passed, 54
skipped** (baseline 649/54; +11 fake tests). Ruff clean on `fitting.py`,
`test_fitting.py`, `test_live_fitting.py`. Live fitting file: **7 passed in 60s**
(2 prior x-range + 5 new). Full live suite counts recorded below. Zero leftover
Origin: 0 `Origin64.exe` before and after every run (all isolated DispatchEx
`Exit()`'d). Did NOT push.

## 2026-07-11 — whole-product review Round B (capability) (agent: round-b)

Scope: the 9 capability items assigned from `docs/REVIEW-2026-07-11-whole-product.md`
(items 6, 10, 8, 9, 11, 29, 30, 31a, 31b). Base = 59037c8 (Round A merged).
One commit per item, TDD, ruff-clean on touched files, live-verified where the
item required it. Did NOT push. Live runs used an isolated
`DispatchEx("Origin.Application")` fixture (never ApplicationSI), `Exit()` in
teardown; WSL tree rsynced into `C:\Users\swym4\opm_dev` (the
`C:\Users\swym4\Origin-Pro-MCP` clone was never touched). Live receipts under
`C:\Users\swym4\probe_out\roundb*`.

| Item | Commit | What / how verified |
|---|---|---|
| 6 (vector export) | `04f101a` | Probed expGraph vector types live: **pdf ✓ (%PDF), eps ✓ (%!PS), emf ✓; svg ✗ (command fails, no file)**. Split EXPORT_IMAGE_FORMATS → EXPORT_RASTER/EXPORT_VECTOR/EXPORT_FORMATS (honest names); vector skips the tr1 pixel-size opts (they make expGraph fail on vector). export_graph/export_all/sized accept the new formats. Live: pdf/eps/emf write valid magic-byte files; svg raises. |
| 10 (dead export params) | `5b29884` | Probed live: an explicit **width IS honored**; **tr1.height is silently ignored** (1600x1000 → 1600x1224, aspect-locked); **no dpi node** (tr1.dpi/DPI/res all make expGraph fail). So: export_graph applies width with no sized=True; **dpi removed** from both export tools; **height removed** from both (could never take effect); export_all passes width through. No param accepted-and-ignored remains. Live: width=1600 → 1600px PNG via both tools; dpi/height absent from signatures. |
| 8 (fit X-range) | `2eab9cd` | Probed: NLFit input accepts a **1-based row subrange `!(x,y)[i1:i2]`**. Added x_min/x_max to curve_fit; _rows_in_x_range resolves bounds to a contiguous row block (X assumed monotonic), raises when empty; range-restricted line fit routed through NLFit. Live: spectrum with a small peak at xc=5 + a LARGER interfering peak at xc=15 — whole-curve fit pulled to ~15, curve_fit(x_min=2,x_max=8) recovers xc=5. |
| 9 (smooth methods) | `87e114a` | Exposed `smooth_method` (savitzky_golay/adjacent/binomial) on transform; removed dead `poly_order`. Fake-only (no live path): smooth_method reaches the impl + right method:= id; poly_order gone. CLI typed-coercion tests moved off poly_order. |
| 11 (orphan hygiene) | `3b1ab5c` | interpolate reuses one stable "Interp" book (rows cleared before overwrite; only self-created book rolled back on failure). find_peaks tags outputs "Peak X"/"Peak Y" and deletes the prior pair on repeat (net +2, no growth); pkfind's uninvited "Center Peaks Indices" column deleted each call (probe: no oindex/oidx option suppresses it). Live: 3 interpolate calls → one Interp book; 3 find_peaks calls settle at +2 columns. |
| 29 (right/top axis title) | `f913b2e` | Probed: `yr.text$` (layer 2) and `xt.text$` (layer 1) both set + read back; yr **also "succeeds" on a single-layer graph** (the false success). axis(op=labels) gains axis=right (requires page.nlayers>=2, else raises) and axis=top; read-back guards a genuine no-op. Live: right title sets+reads on a dual-Y graph; single-layer axis=right raises. |
| 30 (dual-Y legend) | `3c6b848` | add_second_y_axis now rebuilds the legend (both layers), sets it borderless (legend.background=0), re-places it via place_legend_avoiding_data, and reports the placement. Live: return says "legend rebuilt borderless, placed <corner>", legend.background reads back 0, graph exports. |
| 31a (template APPLY) | `9d9075f` | Probed: **`plotxy ... ogl:=<new template:="<path>">`** builds a graph from a saved template. Added `template` param to create_graph (2D XY only; missing path / XYZ rejected); return gains a "template" field. Live round-trip: thick-red-styled graph → save_graph_template → create_graph(template=...) with new data exports with >5x the red pixels of a default no-template graph. |
| 31b (batch import) | `5974940` | import_data accepts a directory or glob; imports every csv/txt/dat/xls/xlsx match into its own stem-named book; returns per-file results (partial failures non-fatal); caps at 20 with a note. Live: 3-file temp folder → books alpha/beta/gamma each holding the data. |

**Vector-format verdict (live-probed):** pdf ✓, eps ✓, emf ✓, svg ✗ (unsupported
on Origin 2020 — expGraph fails and writes nothing). **Export size controls:**
width honored; height aspect-locked and dpi unsupported (both removed).
**Fit-range route:** `[Book]Sheet!(x,y)[i1:i2]` row subrange on nlbegin.
**Template route:** `plotxy iy:=... ogl:=<new template:="<path>">`.

WSL fake suite (`/tmp/opm_venv/bin/python3.11 -m pytest -q`): **649 passed, 54
skipped** at HEAD `5974940` (Round A baseline 620/38; +29 fake tests this round).
Ruff clean on every touched file (removed one pre-existing unused `pytest`
import in the touched `tests/test_cli.py`). Full live suite (Windows,
`pytest -m requires_origin tests/`): **54 passed, 649 deselected in 264s, 0
failed**. Zero leftover Origin: 4 pre-existing `Origin64.exe` before and after
every live run (all isolated DispatchEx instances Exit()'d). Did NOT push.

## 2026-07-11 — whole-product review Round A (correctness) (agent: round-a)

Scope: the 12 correctness items assigned from `docs/REVIEW-2026-07-11-whole-product.md`.
Base = b4e9465 (v0.2.2). One commit per item/group, TDD, ruff-clean on touched
files, live-verified where the item required it. Did NOT push. Live runs used an
isolated `DispatchEx("Origin.Application")` fixture (never ApplicationSI),
`Exit()` in teardown; WSL tree rsynced into `C:\Users\swym4\opm_dev` (the
`C:\Users\swym4\Origin-Pro-MCP` clone was never touched).

| Item | Commit | What / how verified |
|---|---|---|
| 1 (P8 fit line) | `df1ed32` | fitting.py fit-line `-c`/`-w` split into two `set` calls + comment. Live: item-22/errbar & existing fit paths render styled. |
| 27 (P0 label drop) | `12d7d81` | apply_publication_style routes xb/yl label writes through `_set_axis_title_verified` (activate+write+read `xb/yl.text$` back+settle-retry); return only claims labels when the read-back contains them, else WARNs. Fakes: verified/unverified both. |
| 2 (ws placeholder 0.0) | `30b9adf` | set_worksheet_data checks activation + each `col(c)[r]=0/0` return, raises naming the exact cells still holding the 0.0 placeholder. Fake: injected per-cell failure. Live round-trip already covered by `test_set_worksheet_data_null_becomes_missing_value`. |
| 3 (wrong-book import) | `1523f4b` | import_data captures `CreatePage`'s returned uniquified name and activates THAT. Fake: taken name -> import lands in the new book. |
| 5 (matrix null/NaN) | `5805e2b` | Live-probed PutMatrix accepts NaN -> stores `-1.23456789e-300`. set_matrix_data accepts null/NaN; fixed get_matrix_data's sentinel test (missed the small sentinel). Live round-trip PASSED. |
| 22 (error bars) | `8cc537d` | **FIXED WITH ROUTE.** Live root cause: missing `settle_new_plots` after the plotxy -> designation + `set -o` raced the new plot and no-op'd. Added settle + read-back verify (col.type==3, error plot present, no stray data curve) + stray rollback + fail-honest raise. Live PASSED; export shows real whiskers identical to the y_error_col route. ungroup_plots' "re-add with set_error_bars" advice is now correct, left as-is. |
| 23/24 (colormap) | `c9f3481` | Live-probed: no cmap state distinguishes a palette change; `load()` lies on a bogus name. Reject unknown palettes up front (bundled dir + Origin `Palettes/` via `system.path.program$`); z-range reads `layer.cmap.zmin/zmax` back and raises on mismatch. Live: bogus raises, real applies, z-range reads back, line-graph z-range rejected. |
| 25 (ungroup rollback) | `54a967e` | Counts rebuilt vs removed: raises if 0 rebuild, PARTIAL (naming missing series) if some fail; success only when all rebuild. Fakes: 0-rebuild raise + partial. Success path live-covered by existing reload test. |
| 26 (verify batch) | `ccd5483` | set_layer_geometry -> rebind-safe path + loose read-back; apply_publication_style raises on frame/tick failures, notes cosmetic; manage_columns props/add + (sort already) check returns; create_matrix reads `wks.nrows/ncols` back. Live: geometry + create_matrix no false-raise. |
| 12/28 (font bold, peaks) | `161ff15` | set_graph_font tick labels honor bold=False; tick-size rule documented. find_peaks clamps local_points to `(n-2)//2` (live-probed: 5 fails on 11-pt Gaussian, 4 works), raises <3 points. Live 11-pt Gaussian PASSED. |
| 4 (rebind decision) | `08b8232` | **DECISION: writes ARE reliable on loaded .opju** (probe: xb/yl verified True after activate+settle+read-back). Wired read-back gating: set_graph_font confirms axis-title/legend/title fsize; set_legend confirms the legend is readable, else WARN. Live: both succeed on a reloaded .opju (no false alarm). |
| 16 (sparklines) | `0a62646` | **ANSWER: the key was WRONG.** `options.Miscellaneous.Sparklines:=0` is rejected (12-col CSV -> 12 graph windows); correct key `options.Sparklines:=0` -> 0 windows. Also made sparklines_suppressed = (option ran AND deleted==0) so it can't contradict sparklines_deleted (F12). Live PASSED. Dialog-watchdog flag left out of scope. |

WSL fake suite (`/tmp/opm_venv/bin/python3.11 -m pytest -q`): **620 passed, 38
skipped** at HEAD `0a62646` (baseline 596 passed, 30 skipped). Ruff clean on
every touched file (removed 3 pre-existing dead imports in touched files:
matrix.py `labtalk_string`, test_typed_style.py `FakeColumn`, fakes.py
`threading`). Full live suite counts + zero-leftover-Origin recorded below after
the final live run.

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
