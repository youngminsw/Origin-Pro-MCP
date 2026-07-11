# Whole-product review — 2026-07-11 (v0.2.2, HEAD b4e9465)

Direct review by the orchestrator (not delegated). Method: full read of every
tool implementation + docstring in `src/origin_pro_mcp/tools/` (style, graph,
worksheet, fitting, analysis, matrix, project, labtalk, style_helpers),
autosave/daemon dispatch paths, publication-figure skill, README structure;
plus a measured context-cost scan (AST over all `@mcp.tool` docstrings) and a
targeted grep sweep for the P8 multi-flag class. Live-unverified items are
marked SUSPECT; everything else is confirmed by code reading.

## P0/P1 — correctness: silent failures & data integrity

1. **[CONFIRMED] P8-rule violation left in `curve_fit`** — `fitting.py:231`
   styles the fit line with `set <plot> -c color(...) -w 400;` (two flags,
   one command; the comment even still carries the old batching rationale).
   The 2026-07-10 probe proved flag-batching is what silently corrupts plots
   on Origin 2020; `set_plot_style`/`apply_publication_style` were refactored
   but fitting.py was missed. Fix: two separate `set` calls. Cheap; live-check
   that the fit line stays styled.
2. **[CONFIRMED] `set_worksheet_data` can leave a placeholder 0.0 as real
   data** — `worksheet.py:155-159`: null/NaN cells are bulk-written as 0.0
   then overwritten per-cell with `col(c)[r]=0/0;`, but the activation and
   per-cell `execute_labtalk` returns are IGNORED. If activation or any cell
   write fails, the user's "missing" cell silently becomes numeric 0 — worst
   class of bug (wrong data, success message). Fix: check returns, raise
   naming the cells; or read back one written cell.
3. **[CONFIRMED] `import_data(book_name=...)` can import into the WRONG
   book** — `worksheet.py:250-253`: `o.CreatePage(2, requested, "origin")`'s
   return (the ACTUAL, possibly uniquified name) is discarded, then
   `win -a <requested>;` activates whatever already owns that name, and
   `impasc` imports into the ACTIVE window. If the requested name existed,
   the import lands in the pre-existing book. Fix: capture CreatePage's
   return and activate that.
4. **[SUSPECT, known-documented] `set_graph_font` / `set_legend` wrong-window
   writes on loaded graphs** — both use global text-object LabTalk
   (`xb.*`, `yl.*`, `legend.*`) after `activate_window`; LabTalk returns true
   even when the active window is not the target, so the existing `if not
   execute_labtalk` guards cannot catch a silent wrong-target/no-op on
   .opju-loaded graphs. Accepted-gap since the S2 round; this review re-flags
   it because every OTHER styling path is now rebind-protected and these two
   are the odd ones out.
5. **[CONFIRMED] `set_matrix_data` rejects null/NaN** — inconsistent with
   `set_worksheet_data` (which now accepts them); heatmaps with missing cells
   are a real materials-science case. Fix: accept null → Origin missing (via
   NaN in PutMatrix if it works live, else per-cell LabTalk).

## P1 — capability gaps that force workarounds

6. **No vector export** — `EXPORT_IMAGE_FORMATS = {png,jpg,tif,bmp}` only.
   Journal submission needs SVG/EPS/PDF; Origin's expGraph supports vector
   types. Biggest single usability win available. (export_graph param, no new
   tool.)
7. **No multi-peak fitting** — `curve_fit` is single-peak only
   (gauss/lorentz/voigt); XRD/XPS/Raman deconvolution impossible. Known
   deferred feature; remains the largest analysis gap.
8. **No fit X-range restriction** — can't fit a sub-range (exclude baseline
   /neighboring features); standard need. (curve_fit params x_min/x_max.)
9. **`transform(method="smooth")` hides its methods** — adjacent/binomial are
   implemented (`_SMOOTH_METHODS`) but unreachable from the tool signature;
   `poly_order` param is declared "Reserved; not used" (schema noise +
   confusion). Fix: add `smooth_method`, drop or wire `poly_order`.
10. **Dead export params mislead agents** — `export_all_graphs(dpi/width/
    height)` and `export_graph(dpi)` are accepted and ignored ("kept for API
    compatibility"). An agent that sets dpi=600 believes it worked (the
    IGNORED note only appears when values differ from defaults, and
    export_all_graphs has no note at all — it silently ignores). Fix: honor
    them via the sized path or remove.
11. **Orphan-object accumulation** — `transform(interpolate)` creates a new
    "Interp" book EVERY call; `find_peaks` appends 2 columns per call;
    `curve_fit(plot_on_graph=...)` leaves report sheets (documented). Repeat
    calls bloat the project the same way sparklines used to. Fix: reuse a
    stable output book/columns, or note + cleanup guidance.
12. **`set_graph_font` surprises** — tick labels are FORCE-bolded even with
    bold=False, and tick size is silently font_size-4 (min 16) — neither is
    in the docstring. Fix: honor bold for ticks (set_tick_labels already
    can), document the -4 rule or add a tick_size param.

## P2 — docs, consistency, context cost

13. **Tool docstrings cost ~9.1k tokens per session** (36.5KB across 43
    tools, measured by AST). `run_labtalk` alone is 6.1KB (17%) and grows
    every round as the gotcha ledger. Proposal: keep one-line warnings in
    docstrings, move the detailed recipes/tables (symbol table, unit tables,
    protocol recipes) into the `publication-figure` skill or a new
    `labtalk-gotchas` skill served by get_skill (loaded on demand), and state
    "see get_skill('...')" — est. 3-4k tokens saved per session for every
    agent, no safety loss.
14. **[SUSPECT] FFT docstring claim** — "uses the X column spacing as the
    sampling interval" but the command is `fft1 ix:=col(<y>)` with no X
    passed (`analysis.py:185`); the Frequency column units claim needs a live
    check or a docstring correction.
15. **Return-format inconsistency** — some tools return JSON objects
    (create_*, import, fits), others plain strings (styling), `stats`/`
    transform` mix both across ops. Documented per-tool, so P2; a convention
    note in README suffices.
16. **Remaining LIVE-UNVERIFIED flags** — impASC `Sparklines:=0` option key
    (worksheet.py:270), DialogWatchdog WM_CLOSE on real mid-session dialogs,
    attach-mode notices end-to-end. Each has a defensive fallback, but they
    have never been exercised on real Origin. One deliberate live session
    would clear or fix all three.
17. **Minor**: `create_matrix`'s `mdim` return unchecked (confusing later
    error if it fails); WSL suite showed one unidentified intermittent
    failure (1-in-5 runs, vanished on rerun) — worth catching with `-p
    no:randomly`-style triage if it recurs; `matrix.py:16` pre-existing F401.

## Addendum — verified findings from the (cancelled) docs-audit agent

Its report arrived after cancellation; the items below were RE-VERIFIED
directly before inclusion (grep receipts in session log):

18. **[CONFIRMED] `remove_plot` documented three contradictory ways** —
    README.md:363 and publication-figure.md:411-413 still describe the OLD
    `layer -e`/`layer -ie` mechanism replaced by d5a1f14; the skill's own COM
    Notes table (line 459) correctly describes the new per-index
    `DataPlot.Destroy()` — the skill contradicts itself 48 lines apart. Fix:
    two doc edits.
19. **[CONFIRMED] stale tool count** — README.md:315 "45 total" and
    publication-figure.md:337 "45 tools"; actual is 43 (AST-verified).
    test_doc_registry checks names, not the count — add the count to the test
    or drop the number from prose.
20. **[CONFIRMED] `color` param means different formats in different tools** —
    set_plot_style: `color`=named ("blue"), `rgb`="r,g,b"; add_second_y_axis:
    `color`="r,g,b". Same name, different format, no cross-reference. Fix:
    accept both formats in one param going forward, or cross-reference.
21. (Context-cost cross-check: docs agent measured 46.3KB incl. signatures /
    ~11.6k tokens vs my 36.5KB docstring-only — same conclusion, run_labtalk
    is ~15-17% of the always-loaded budget and its gotcha block is ~2KB of
    duplicated content.)

## Addendum 2 — live-confirmed findings from the (cancelled) silent-failure agent

Receipts: `C:\Users\swym4\probe_out\review\silent_failures.md` (+ probe
scripts alongside). Its foundational fact matches the codebase's own history:
`execute_labtalk()` returns "script parsed", NOT "had effect".

22. **[CONFIRMED-LIVE P1] `set_error_bars` (post-hoc attach) does not attach**
    — the emitted `plotxy`→`set <er> -o <yr>` sequence live-tested: the whole
    -script path fails (tool raises truthfully) but leaves a STRAY CURVE +
    partial column-type mutation with no rollback; the split path returns
    True while DataPlots stays 2 and col.type reads 0 — i.e. the tool's
    success message can be a lie and its failure path leaves damage. NOTE:
    the `y_error_col` route in create_graph/add_plot_to_graph is the one that
    works; ungroup_plots' advice "re-add with set_error_bars" points at the
    broken path. Remedy: settle after plotxy, read-back (col.type==3 +
    DataPlots is_error), rollback the stray plot on failure.
23. **[CONFIRMED-LIVE P1] `colormap(palette=)` no-verify** — `layer.cmap.
    load()` returns True for a BOGUS palette name and on a non-colormap
    graph; tool reports "Applied palette" doing nothing. Remedy: read-back
    (e.g. cmap.colorlist / numColors delta) like the levels branch already
    does.
24. **[CONFIRMED-LIVE P1] `colormap(z_min,z_max)` no-verify** — same class;
    its sibling levels branch got a read-back in the g2-g8 round, this branch
    was not updated.
25. **[SUSPECTED P1] `ungroup_plots` no-rollback** — removes ALL plots before
    re-plotting; a replot failure yields "rebuilt 0 plots" over an emptied
    graph. Remedy: verify plotxy successes and restore/raise loudly.
26. Its suspects S5-S11 overlap items 4/12/17 above and add:
    `set_layer_geometry` uses the unprotected global route (unlike its
    axis-range sibling), `apply_publication_style` steps 1-6 ignore returns,
    `manage_columns`/`sort_worksheet`/`create_matrix` no read-back — fold
    into Round A/B as cheap `graph_layer_execute`+verify upgrades.

## Addendum 3 — live-confirmed findings from the (cancelled) usability agent

Receipts: `C:\Users\swym4\probe_out\review\usability.md` (+ wf?_log.json,
wf?_*.png). Five workflows run end-to-end on live Origin.

27. **[CONFIRMED-LIVE P0] `apply_publication_style` silently drops the axis
    LABELS it claims to set** — titles stay `%(?X)`/`%(?Y)` (column names) in
    3/3 runs while the return says "Arial bold labels". Root cause: the label
    step uses the unbarriered active-window `execute_labtalk('xb.text$=...')`
    route while every other step uses `graph_layer_execute` — the same settle
    race documented in Task 0.5. Fix: route the label writes through the
    barriered path + read back `xb.text$`.
28. **[CONFIRMED-LIVE P1] `transform(find_peaks)` default local_points=10
    fails on short spectra** with an unactionable error (11-pt Gaussian:
    lp<=3 finds the peak instantly). Fix: clamp to data length + actionable
    error.
29. **[CONFIRMED-LIVE P1] `axis(op="labels", axis="y")` falsely reports
    success for a right-axis title** — only the raw `yr.text$` recipe works;
    no typed path exists. Fix: add right-title support (add_second_y_axis
    param or axis op) + stop the false success.
30. **[CONFIRMED-LIVE P1] dual-Y legend overlaps + border returns after
    `add_second_y_axis`** (wfB_dualy_final.png). Fix: re-run the borderless
    placement after the second layer lands.
31. **[P1 gaps, corroborating my items 6-8]** no template-APPLY tool (save
    exists, reuse doesn't), no batch/folder import, no multi-peak fit — the
    three "reuse/scale" paths that force raw LabTalk.
32. **[CONFIRMED-LIVE P2] skill advice `offset_pct=80` collides tick numbers
    with the axis** on the reference workflow — soften the skill's number
    (docstring-only). Also P2: high `colormap(levels=)` over-labels the
    colorbar; import_data auto-names force a list_worksheets round-trip;
    sparklines telemetry fields can contradict each other; grouping NOTE
    nudges needless ungroup calls.
33. **[OBSERVATION] stale `Origin64.exe -Embedding` processes accumulate**
    (4 present, ages 14h-2d, predating all review probes) — consistent with
    detach-keep by design, but worth an eventual "orphan report" surface so
    users notice; already partially covered by the session-ledger notices.

## Suggested execution order

Round A (correctness, small): items 1, 2, 3, 5, 12 + docstring fixes from 10.
Round B (capability): 6 (vector export), 8 (fit range), 9 (smooth methods),
11 (orphan hygiene), 10 (honor dpi/width).
Round C (big feature): 7 (multi-peak fitting) — needs design (probe NLFit
multi-peak / `pa` X-Function routes first).
Round D (docs/CTX): 13 (docstring→skill split), 14-16 sweeps.
Item 4 needs a decision: either rebind-protect the two tools (text objects
via layer handle, if possible on Origin 2020 — probe) or loudly document.
