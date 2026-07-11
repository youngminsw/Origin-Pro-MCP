# LabTalk Gotchas

The full, probe-verified ledger of LabTalk/COM quirks discovered while
building and hardening this server (Origin Pro 2020). `run_labtalk`'s
docstring keeps only a short safety-critical list at the point of use; this
skill is the detailed reference — read it before writing non-trivial raw
LabTalk, or when a typed tool's result looks wrong and you need to know why.

## When To Use

Use this before writing raw `run_labtalk` scripts, or when a styling/analysis
call through any tool silently no-ops or looks wrong on a loaded/reloaded
graph — most of these gotchas explain a "success" message that didn't do
anything.

## P8: never batch `-flag`s in one `set` command

Probe-confirmed on Origin 2020: combining multiple `-flag`s in ONE
`set <dataset> ...` command (e.g. `-c` + `-cf`, or `-k` + `-kf` + `-z`)
silently corrupts the plot — colors reset to BLACK, or the symbol blanks out.
This reproduces reliably; it is not a timing fluke. Evidence: the 2026-07-10
flag-batching probe found every combined-flag script produced a visibly wrong
plot even though `execute_labtalk` returned success. Send exactly ONE flag per
`set <ds> -flag val;` call. `set_plot_style` and `apply_publication_style`
both do this; never batch flags yourself via `run_labtalk` either.

## `layer.x2/y2.majorTicks` — never write it

`layer.x2.majorTicks` / `layer.y2.majorTicks` set to 0 wipes the NUMBER LABELS
on ALL FOUR axes, not just the opposite side. Use `layer.<ax>.ticks = 0` to
remove tick MARKS instead (`axis(op="tick", tick_direction="none")` does
this). Never write `majorTicks` directly from `run_labtalk`.

## `-w`/`-erw`/`-erwc` unit tables

Line width and error-bar width are NOT the same unit scale — this is the
single most common mistake when styling from raw LabTalk:

| LabTalk flag | What it sets | Unit |
| --- | --- | --- |
| `set <ds> -w <n>` | data-line width | ~200 units per point (500 ≈ 2.5pt) |
| `set <er> -erw <n>` | error-bar LINE width | POINTS directly |
| `set <er> -erwc <n>` | error-bar CAP (whisker) width | POINTS directly |

Passing a `-w`-scale value (e.g. 550) into `-erw` makes bars explode visually
— that is a units mistake, not an Origin bug. `-ew` is a silent no-op on an
error plot; use `-erw`.

## Symbol shape `-k` codes

Re-verified live on Origin 2020:

| Code | Shape |
| --- | --- |
| 1 | square |
| 2 | circle |
| 3 | triangle-up |
| 4 | triangle-down |
| 5 | diamond |
| 6 | plus |
| 7 | x / cross |
| 8 | asterisk |
| 9-12 | render as a dash/vertical-bar/literal glyph — not useful shapes, avoid |

## Tick-label offset property names (`offsetV`/`offsetH`)

`set_tick_labels(offset_pct=...)` writes `layer.x.label.offsetV` for the X
axis and `layer.y.label.offsetH` for the Y axis — the property name swaps
axis letter for the PERPENDICULAR direction (an X-axis label moves
vertically; a Y-axis label moves horizontally), which is not obvious from
the property name alone. Positive values pull labels TOWARD the axis;
negative pushes them away; 0 is Origin's default.

## `y2`/`YR` right-axis recipe (dual-Y color/title)

The visible right axis on a dual-Y graph is the SECOND layer's **`y2`**
(positional right) axis — NOT its `y` (that is layer 2's hidden LEFT axis;
writing `layer.y.*` there is a silent no-op, the trap that made the right
axis look unreachable). With the graph active and layer 2 selected:

```text
win -a <graph>; page.active=2;
layer.y2.color=color(214,96,77);        // right axis LINE + its ticks
layer.y2.label.color=color(214,96,77);  // right tick-NUMBER labels
doc -uw;
```

`add_second_y_axis(..., color="r,g,b")` emits exactly this. The right-axis
TITLE is a separate object, the `YR` text object (not `yl`):

```text
yr.text$="\b(Magnetization (emu/g))";  // title text + BOLD via \b()
YR.color=color(214,96,77);              // title color (YR.bold=1 no-ops)
YR.fsize=35;                             // title size
```

For layer 1, the LEFT title colors via `YL.color=color(...)` (`layer.y.color`
alone colors only layer 1's line + ticks, not its title).

## Layer-2 `set -w` protocol (connecting-line width on the right axis)

Layer-2 line width via `set <ds> -w <n>` looks non-deterministic ("flaky")
unless you follow the full protocol: `win -a <graph>; page.active=2;
set <ds> -w <n>; doc -uw;` plus a ~2s settle (5/5 repeats identical live). It
SATURATES around `-w 8` (`w=8` and `w=20` render identically) — to exceed
that thickness, plot a scaled copy of the data on layer 1 and style it there
with `set_plot_style(line_width=...)` instead.

## Settle race after page creation

A `set`/`layer.*` command issued through `run_labtalk` immediately after
`create_graph`/`add_plot_to_graph`/`plotxy` can silently no-op (return
`false`, or return `true` and apply nothing) for the FIRST 1-3 commands,
because the new page has not finished settling. The typed tools
(`set_plot_style`, `axis`, etc.) go through a settle barrier that prevents
this; raw `run_labtalk` does not. If you must style right after building,
prefer the typed tool, or issue a throwaway `doc -uw;` and re-issue, or pass
`window=<graph>` and expect the first call may need a retry.

## Readable properties vs. the missing-value sentinel

`layer.x/y.*` properties (from/to/inc/thickness/ticks/majorLen) read back
real values; `layer.x2/y2.*` (opposite-axis) reads more often return
Origin's "missing" sentinel — a tiny near-zero float (e.g.
`-1.23456789e-300`), not a huge negative number, and not the same as a
genuine zero. `run_labtalk`'s `capture` translates that sentinel to the
string `"missing"` instead of leaking the raw float.

## Missing-value cells: `0/0`, not `NANUM`

To write an Origin missing-value cell from LabTalk, use
`col(<c>)[<r>]=0/0;` — `NANUM` is a no-op on this build. `set_worksheet_data`
does this for you when a cell is `null`/`NaN` in the input; it also checks
the write actually landed before reporting success.

## Hard limitations (no scriptable route exists)

Two Origin 2020 features have NO scriptable route — every documented and
undocumented flag was tried live and no-ops (byte-identical export, zero
visible effect), so don't spend time hunting for one:

- **Bar/column fill PATTERN (hatch)**: `set <ds> -pfp/-pfw/-pfc` (the
  documented flags), indexed patterns `-pfpd/-pfpi`, the range tree
  `rp.pattern.*`, and the plot-object tree `layer.plotN.pattern.*` all
  no-op. Apply hatching in PPT/Illustrator, or distinguish series by color
  instead.
- **Plot TRANSPARENCY / alpha**: `set <ds> -paap/-paal/-paas`,
  `rr.transparency`, `layer.transparency`, `page.transparency` all return
  success but no-op. Use fully-opaque colors, or composite the alpha in PPT.
  (Unconfirmed future candidate: the `originpro` Python package's
  `Plot.transparency()` method — not installed in this environment to test.)
