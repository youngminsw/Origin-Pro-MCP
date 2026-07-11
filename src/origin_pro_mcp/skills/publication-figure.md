# Publication Figure

Create paper-grade figures in Origin Pro with the `origin-pro` MCP server. Treat this skill as a generic starting point: copy it, tune it, and keep your own lab or journal presets close to the projects where you use them.

## When To Use

Use this when the user asks for a manuscript figure, publication-quality graph, journal-ready plot, reference-paper style, or polished Origin export.

## Requirements

- Windows with a licensed Origin/OriginPro installation
- `origin-pro` MCP launched by the MCP client through Windows Python
- Prefer typed MCP tools over raw LabTalk
- Use `run_labtalk` for advanced graph and analysis commands; destructive and file-writing commands are blocked. If a multi-statement script fails as a whole, the server automatically retries it statement-by-statement and reports each one's OK/FAILED status — proactively splitting a script into small chunks is still fine, but no longer required to see which statement caused a failure
- Verified on Origin Pro 2020 only; other versions may work but are not verified by this skill

## First Pass

Ask only for missing choices that materially change the figure:

- Data source: arrays, worksheet, CSV, or existing Origin workbook
- Figure type: scatter, line+symbol, column, histogram, box, contour, or another Origin plot
- Series mapping: X, Y, error bars, groups, panels
- Target: general manuscript, specific journal, conference slide, or thesis
- Export path and format

If unspecified, choose a conservative manuscript default: line+symbol for ordered data, scatter for independent observations, white background, colorblind-safe colors, readable labels with units, no grid, inward ticks, closed frame, and PNG export.

## Design Defaults

- Use a restrained scientific style, not decorative styling.
- Put units in axis labels: `Temperature (K)`, `Current density (mA cm^-2)`.
- Use the muted pastel palette (`apply_publication_style` default): soft
  steel blue, muted rose, muted teal, soft amber, soft purple, gray cyan.
  Avoid pure primary colors — they tire the eyes and dominate the figure.
- Avoid red and green as the only two series colors.
- `apply_publication_style` defaults are tuned for print: 3.0 pt data
  lines, size-14 symbols, and 2.5 pt error-bar lines with whisker caps
  matched to the symbol size.
- **Axis rules (strict — `apply_publication_style` enforces these):**
  - **No empty axis gap on a linear graph.** The axis range is TIGHT to the
    data — no padding before the first or after the last point. If the data
    starts at 0, the axis starts at 0 (0 is the Y-intercept; the curve touches
    the Y axis). Never leave a blank stretch where there is no data.
  - **Major ticks capped at 4–6 per axis.** Never crowd the axis with ticks;
    the tool picks a round increment giving ≤6 labelled ticks, with sparse
    minor ticks. If you set ranges manually, keep within 4–6 major ticks.
- **Keep the project lightweight — delete bad graphs.** When a figure comes
  out wrong (mis-scaled, wrong plot type, rejected attempt), remove it with
  the `delete_graph` tool instead of leaving it in the project. A project full
  of discarded graphs gets heavy and slow.
- **The legend must never overlap the data.** `apply_publication_style`,
  `set_legend`, and the fit-curve path auto-place the legend in the
  emptiest corner (the requested `legend_position` is only a preference,
  overridden when data sits there). **When every inside corner still
  overlaps data, the legend is automatically moved OUTSIDE the frame — to
  the right of the plot, vertically centered — and the plot is shrunk to
  make room.** The tool reports `outside-right` and "moved outside the frame
  to avoid the data" when this happens; no manual step is needed. (A dense
  scatter that fills all four corners is the usual trigger.)
- Legends: borderless with bold entries (the style tools do this).
- Use distinct symbols when colors alone may not survive grayscale printing.
- Keep axis ranges tight but honest; do not crop important outliers.
- Use legends only when labels cannot be placed directly or series count requires it.
- Remove grid lines unless the plot type genuinely needs them.
- Use a closed frame for journal-style single-panel graphs.
- Export once, inspect the image, then adjust text size, legend placement, and axis range.
- Supported plot types via `create_graph`: scatter, line, line+symbol,
  column, bar, area, pie, histogram (single Y), box (single Y),
  contour and 3d_scatter (need `z_col`). Matrix plots — 3D surface,
  contour, heatmap, image — use `create_matrix_plot`. Colormapped plots
  take `colormap(...)`. Annotate with `annotate(kind="text"/"line"/"arrow"/
  "reference_line", ...)`; save reusable layouts with `save_graph_template`.

## 3D and Colormap Design Defaults

Same spirit as the 2D defaults (clean, honest, colorblind-safe, units
everywhere, readable bold labels), adapted to surfaces, contours,
heatmaps, and 3D scatter.

- **The colormap is the 3D palette — choose it as carefully as 2D colors.**
  Use a perceptually-uniform, colorblind- and grayscale-safe map. The server
  **bundles the gold-standard scientific colormaps** (viridis/cividis were
  added to Origin only after 2020, so they ship with the MCP and load by name
  via `colormap(graph_name, palette=...)`): `Viridis` (default, best general sequential),
  `Cividis` (optimized for red-green CVD), `Plasma`/`Inferno`/`Magma`
  (high-contrast sequential). Origin's own colorblind-safe built-ins also
  work: `Heatmap4ColorBlind`, `GrayScale` (print-safe), `RedWhiteBlue`
  (diverging). **Avoid `Rainbow`/`Jet` AND `BlueGreenYellow`** for
  quantitative data — they are not perceptually uniform, invent false
  boundaries, and fail in grayscale. Verify by eye: a good map reads as a
  smooth light→dark ramp with no banding.
- **Pastel/muted aesthetic:** to keep colormaps in the same soft, low-saturation
  family as the figA/B/C pastel series, use the bundled `PastelViridis` /
  `PastelCividis` (viridis/cividis blended toward white) instead of the punchy
  full-saturation versions. They stay perceptually-ordered and colorblind-safe
  but read soft, matching muted 2D series colors.
- **Sequential vs. diverging:** sequential map for one-directional
  magnitude (0→max); diverging map with the midpoint at 0 for signed data
  (deviation, difference, charge).
- **A labeled color scale is the legend of a colormap plot — never ship a
  heatmap/contour/surface without it.** `create_matrix_plot` now plots every
  type (surface/contour/heatmap/image) from the Origin system template that
  carries a data-linked color scale, so the scale appears automatically and
  always reflects the real palette and Z range. To match the rest of the
  figure (figA-style bold Arial), drive the `Spectrum1` object via
  `run_labtalk`. **The numeric labels only honour `bold`/`font` after you turn
  off auto-display first:**
  `Spectrum1.labels.autodisp=0; Spectrum1.labels.bold=1; Spectrum1.labels.fsize=14; Spectrum1.labels.font=font(Arial); Spectrum1.barthick=130;`
  Keep the numeric labels SHORT — set the format explicitly so you don't get
  `0.5000`/`1.000`: `Spectrum1.labels.numdisp=1; Spectrum1.labels.decplaces=1;`
  (→ `0.0, 0.5, 1.0`). Pin the over/under-range colors to the map's ends so the
  bar has no white speck on top or black block on the bottom:
  `layer.cmap.colorAbove=color(254,243,146); layer.cmap.colorBelow=color(162,128,170);`.
  Bold the scale title with the `\b(...)` escape (a plain `.title$` is not
  bold): `Spectrum1.title$="\b(Intensity (a.u.))";`.
- **Set the Z range honestly** with `colormap(graph_name, z_min=, z_max=)` to
  the real data range; don't clip features into saturation just to boost
  contrast, and state the range.
- **Label every axis with units, including Z — and make EVERY text element
  bold Arial, not just X/Y.** The skill's reference figures bold all titles
  and tick labels; a half-bolded figure looks inconsistent. On 3D plots the
  axis-title objects are `xb` (X), `yl` (Y) and **`zf` (Z)** — note `zl`/`zt`
  silently no-op for the OpenGL Z title — so bold the Z title with
  `zf.text$="\b(Intensity (a.u.))";`. Bold the tick labels through the layer:
  `layer.x.label.bold=1; layer.y.label.bold=1; layer.z.label.bold=1;`. The
  matrix long name (set via `create_matrix_plot(..., z_label=)`) still seeds
  both the Z-axis title and the color-scale title; X/Y titles also via
  `axis(graph_name, op="labels", axis="x"/"y", label=...)`.
- **Prefer a 2D contour or heatmap over a 3D surface when exact values
  matter** — a top-down colormap is easier to read off than a tilted
  surface. Use the 3D surface for shape/intuition, the contour for
  quantitative reading; a contour-projected surface gives both.
- **Keep surfaces uncluttered:** a light mesh or a smooth color-mapped
  surface reads better than a dense wireframe; pick one clear viewing
  angle and keep it consistent across a figure set.
- **3D scatter:** the default Origin 3D scatter is tiny red dots with red
  droplines — restyle it to the figure-set palette and a readable size with
  `run_labtalk`: `set %C -c color(93,143,179); set %C -z 12; set %C -kf 1;`
  (steel blue, size 12, solid fill). Changing the symbol color also recolors
  the droplines, so they fade into a subtle guide instead of dominating. Rely
  on Z color/height — not tiny dots — to carry the third dimension. Its X/Y/Z
  axis titles come from the source column long names, so name the columns with
  units up front (e.g. `column_names="X (mm),Y (mm),Signal (a.u.)"`).
- **Match the rest of the figure set:** same fonts, same export pixel size
  (`export_graph(..., sized=True, width=...)`), same labeling conventions
  as the 2D panels so a mixed figure looks like one family.
- **Keep the 3D box honest and uncrowded.** Give X and Y the *same* range and
  origin so the floor reads as a square grid (e.g. both 0→10); a matrix's
  default 1→N coordinates look lopsided, so set them with
  `run_labtalk("range mm=[Book]MSheet1; mm.x1=0; mm.x2=10; mm.y1=0; mm.y2=10;")`
  before plotting. Use few, short major ticks (3 per axis reads cleanest on a
  small 3D cube — e.g. `layer.x.from=0; layer.x.to=10; layer.x.inc=5;` and
  `layer.x.minor=0; layer.x.majorLen=3;`), and bold the tick numbers
  (`layer.x.label.bold=1`). Keep the color scale compact so it doesn't dwarf
  the cube: `Spectrum1.barthick=130;` plus 3 labels via
  `Spectrum1.levels.major=3; Spectrum1.levels.from=0; Spectrum1.levels.to=1; Spectrum1.levels.inc=1; Spectrum1.levels.inc$=0.5;`.
- **Put the origin (0,0) at the front corner so X=0 and Y=0 meet at one point**
  (common-sense layout). Origin draws X/Y tick labels on the two front edges,
  so reverse the X axis to bring its 0 to the front:
  `layer.x.from=10; layer.x.to=0; layer.x.inc=5;` with Y normal
  (`layer.y.from=0; layer.y.to=10`). Pick a viewing-friendly data layout (e.g.
  tall feature far from the origin) so the front peak doesn't hide the rest.
- **Origin 2020 OpenGL caveat (verified):** axis rotation/tilt
  (`xrotate`/`zrotate`/`psi`), tick-label repositioning, and per-tick hiding are
  all **no-ops via LabTalk** in this version. Consequence: the Z bottom tick
  label collides with the X far-corner label at the Z base, and it cannot be
  nudged apart by script. Fix by hiding the Z tick NUMBERS
  (`layer.z.label.color=color(white)`) and keeping the bold Z title — read Z
  magnitude from the color scale (surfaces). Only `from/to/inc`, tick length,
  bold, font, and label color respond to LabTalk on OpenGL axes.
- Export once, inspect, then adjust colormap, Z range, viewing angle, and
  label sizes — exactly the 2D "export and inspect" loop.

Verified colormap-surface recipe (Origin 2020). `"Matrix"`/`"Graph1"` below
are example names — in practice read the actual `"name"` from the JSON that
`worksheet_to_matrix` and `create_matrix_plot` return (Origin may rename):

```text
worksheet_to_matrix(data_book="D", data_sheet="Sheet1", x_col=1, y_col=2, z_col=3)
create_matrix_plot(matrix_book="Matrix", plot_type="surface", z_label="Intensity (a.u.)")
colormap(graph_name="Graph1", palette="Viridis", z_min=0, z_max=1)
axis(graph_name="Graph1", op="labels", axis="x", label="X (mm)")
axis(graph_name="Graph1", op="labels", axis="y", label="Y (mm)")
export_graph(graph_name="Graph1", file_path="C:\\fig\\surface.png", sized=True, width=1600)
```

This yields a Z-colored surface with bold X/Y/Z titles, a labeled color
scale, and a clean Z range — the 3D counterpart of `apply_publication_style`.

## Standard Workflow

### 1. Prepare Data

Inspect what is already open before creating anything — this avoids
duplicate workbooks and tells you the exact `data_book`/`data_sheet`
names plotting tools expect:

```text
list_worksheets()                         # open workbooks, sheets, graphs
get_worksheet_data(book_name="Data", sheet_name="Sheet1")
```

Use existing worksheets when available. For direct arrays:

> `create_worksheet`, `create_matrix`, `create_graph`, `create_matrix_plot`,
> `import_data`, and `worksheet_to_matrix` return a JSON string with the
> actual assigned name (Origin may rename on collision) — read `"name"`
> from the result for subsequent calls, not the requested name.

```text
create_worksheet(book_name="Data")
set_worksheet_data(
    book_name="Data",
    sheet_name="Sheet1",
    columns="[[x1,x2,...],[y1,y2,...],[err1,err2,...]]",
    column_names="X,Y,Error"
)
```

For CSV files:

```text
import_data(file_path="C:\\Users\\name\\data.csv")
```

### 2. Build The Plot

For one series:

```text
create_graph(
    graph_name="Fig1",
    data_book="Data",
    data_sheet="Sheet1",
    x_col=1,
    y_col=2,
    plot_type="line+symbol",
    y_error_col=3
)
```

For additional series:

```text
add_plot_to_graph(
    graph_name="Fig1",
    data_book="Data",
    data_sheet="Sheet1",
    x_col=1,
    y_col=4,
    plot_type="line+symbol",
    y_error_col=5
)
```

### 3. Apply Baseline Styling

Use the one-call style tool first:

```text
apply_publication_style(
    graph_name="Fig1",
    x_label="Temperature (K)",
    y_label="Absorbance (a.u.)",
    x_min=280,
    x_max=620,
    y_min=0,
    y_max=2.2,
    legend_entries="Pristine,Annealed",
    legend_position="top-right"
)
```

Then fine-tune only what the exported image shows needs work:

```text
set_plot_style(graph_name="Fig1", plot_index=1, color="blue", line_width=2.0, symbol_size=10)
set_plot_style(graph_name="Fig1", plot_index=2, color="orange", line_width=2.0, symbol_size=10)
axis(graph_name="Fig1", op="range", axis="x", range_min=280, range_max=620)
axis(graph_name="Fig1", op="range", axis="y", range_min=0, range_max=2.2)
set_legend(graph_name="Fig1", visible=True, position="top-right", entries="Pristine,Annealed")
```

`set_plot_style` reference (these are the only accepted values):

- `plot_index` is 1-based in the order series were added; error-bar plots
  are not counted.
- `color`: black, red, green, blue, cyan, magenta, yellow, orange,
  purple, gray.
- `symbol_shape`: 0=auto, 1=square, 2=circle, 3=triangle-up, 4=triangle-down,
  5=diamond, 6=plus, 7=x/cross, 8=asterisk (re-verified live on Origin 2020;
  9-12 render as a dash/vertical-bar/literal glyph, not useful shapes).
- `error_bar_width` / `error_cap_width`: error-bar line/cap width in POINTS
  (LabTalk `-erw`/`-erwc`) — do NOT use `line_width`'s units for these, and
  never style error bars via raw `set -w`/`-ew` (see the gotchas table).
- `line_width`, `symbol_size`, `symbol_shape`, `color`/`rgb`, `open_symbol`,
  `error_bar_width`, `error_cap_width` all default to `None`/`""` = "leave
  this aspect unchanged" — pass only what you want to change; a partial call
  never resets the rest of the curve's style.
- `legend_position` (here and in `apply_publication_style`/`set_legend`):
  top-left, top-right, bottom-left, bottom-right — nothing else is valid.

### 4. Fit When Needed

Use fitting only when it answers the scientific question:

```text
curve_fit(data_book="Data", data_sheet="Sheet1", x_col=1, y_col=2, function="line")
```

For the classic paper presentation (data symbols + fit line), pass the
graph name and the fitted curve is drawn on it as a red 2 pt line with a
short legend entry ("Gauss fit"):

```text
curve_fit(data_book="Data", data_sheet="Sheet1", x_col=1, y_col=2,
          function="gauss", plot_on_graph="Fig1")
```

Call `apply_publication_style` BEFORE `curve_fit(plot_on_graph=...)` —
running it afterwards would restyle the fit curve like a data series
(palette color + symbols).

Report fit parameters and statistics in plain language. Do not hide weak fits behind styling.

### 5. Export And Inspect

```text
export_graph(graph_name="Fig1", file_path="C:\\Users\\name\\figures\\fig1.png")
```

After export, inspect whether labels are readable, legends avoid data, axis ranges are appropriate, and error bars are visible.

## Extended Toolkit (beyond 2D line plots)

The server exposes 45 tools, several of them dispatchers (one tool name,
an `op`/`kind`/`method` argument selecting the action). Reach for these
when a figure needs more than a styled XY plot.

### Surfaces, contours, heatmaps

3D / colormapped figures are built from a **matrix**. Grid scattered XYZ
first, then plot:

```text
worksheet_to_matrix(data_book="D", data_sheet="Sheet1", x_col=1, y_col=2, z_col=3)
create_matrix_plot(matrix_book="Matrix", plot_type="surface")   # or contour, heatmap, image
colormap(graph_name="Graph1", palette="Viridis", z_min=0, z_max=1)   # Viridis, Cividis, Plasma, Inferno, Magma, GrayScale, ...
```

For a quick 2D contour or a 3D scatter straight from XYZ columns, use
`create_graph(..., plot_type="contour", z_col=N)` or `plot_type="3d_scatter"`.

Heatmap tick spacing (`layer.y.inc = <n>` via `run_labtalk`) needs the graph
window active first and a `doc -uw;` refresh after — see the Origin COM Notes
table. If the colorbar overflows the page's right edge, shrink the plot
LAYER (not the colorbar, which isn't directly addressable) — same table.

### Statistical / distribution figures

- `create_graph(..., plot_type="box")` and `plot_type="histogram"` (single Y column).
- `stats(op="column")` (descriptive stats), `stats(op="compare_means")`
  (two-sample t-test), `stats(op="frequency")` (binned counts) all return
  JSON — put the numbers in the caption, don't just draw bars.

### Signal / spectra workflows

`transform(method="smooth"/"differentiate"/"integrate"/"interpolate"/"fft"/"find_peaks")`
writes result columns or returns JSON. Typical Raman/XRD flow: smooth →
find_peaks → `curve_fit(plot_on_graph=...)` → annotate the peaks.

### Multi-panel and axis control

- `axis(graph, op="scale", axis="x"/"y", scale="log10")` for decades-spanning
  data. It auto-rescales the axis to the data (range bounds are ACTUAL
  values, not exponents) so a log switch no longer leaves garbage ticks.
- `axis(graph, op="frame", frame="closed")` draws the top+right border axes.
- `set_tick_labels(graph, axis, format="scientific"|"decimal", bold=, decimal_places=)`
  for tick-label number format. Log axes already render 10^n by default.
- `set_tick_labels(graph, axis, offset_pct=)` tightens/loosens the gap between the
  axis and its tick labels (Origin's default sits farther out than matplotlib's).
  Units are % of the tick-label font size; POSITIVE pulls the labels TOWARD the
  axis (smaller gap), negative pushes them away. Applied perpendicular to each
  axis (x moves vertically, y horizontally), so one value is the gap knob — e.g.
  `set_tick_labels(graph, axis="x", offset_pct=80)` to match a matplotlib-style
  tight bottom-axis gap.
- `set_layer_geometry(graph, left=, top=, width=, height=)` (percent of page)
  when an axis title is clipped or panels must line up.
- `set_graph_font(graph, target="axes", bold=True)` bolds axis titles via
  `\b(...)` markup (there is no `xb.bold` on Origin 2020).
- `add_second_y_axis` / `add_layer` for dual-axis or stacked panels. Coloring
  the right axis to match its data needs a LabTalk recipe beyond what the tool
  sets — see `add_second_y_axis`'s docstring and the Origin COM Notes table.
- `annotate(graph, kind="reference_line", orientation=, value=)` (threshold/
  baseline), `kind="line"` / `kind="arrow"` (callouts), `kind="text"`
  (labels at data coordinates).

### Error bars and plot cleanup

- **n=3 mean±SD error bars.** Put the SD/SE in its own column and either
  create the plot with them, `create_graph(..., y_error_col=N)`, or attach
  them afterward with `set_error_bars(graph, book, sheet, y_col, err_col)` —
  which plots the error column, reassigns it as error bars (`set <err> -o <y>`),
  designates it as an error column, and rebuilds the legend, so no stray curve
  or extra legend entry is left. Or designate a column yourself with
  `manage_columns(book, sheet, op="properties", col=col, designation="yerr")` (or `"xerr"`).
- **Open (hollow) markers** — publication standard — via
  `set_plot_style(graph, plot_index, symbol_shape=2, open_symbol=True)`
  (`open_symbol` maps to LabTalk `set -kf 1`; `-kf 0` is solid).
- **Remove a stray/dead plot** with `remove_plot(graph, plot_index)` — it
  destroys only the indexed plot (COM `DataPlot.Destroy()`), so it is safe
  even when the same dataset is plotted more than once (a bare `delete range`
  or a name-addressed `layer -e` would remove every copy).

### Worksheet prep and IO

`manage_columns(op="formula"/"add"/"delete"/"properties")` (column formulas,
add/delete columns, units/long name/designation), `sort_worksheet`,
`transpose_worksheet`, `import_data` (CSV/text or Excel), `export_worksheet`.
`import_data` suppresses Origin's auto-generated sparkline mini-graph windows
by default for CSV/text imports (`sparklines=False`) — no manual cleanup of
those throwaway windows is needed; pass `sparklines=True` to keep them.

### Reuse and sized export

- `save_graph_template(graph, path)` captures a finished layout as `.otpu`
  to reuse across figures.
- `export_graph(graph, path, sized=True, width=1600)` exports at an exact
  pixel width (vs. the default, which exports at ~1200px wide). Both write
  the file directly via expGraph with no clipboard, so the user's clipboard
  is preserved.

## Origin COM Notes

These COM behaviors were observed while testing on Origin Pro 2020. Other Origin versions may behave differently until verified:

| Issue | Practical workaround |
| --- | --- |
| Bold axis title properties are unreliable | Use Origin text markup such as `\b(Label)` through typed tools |
| `%C` plot shortcuts can fail through COM | Use plot names from `FindGraphLayer().DataPlots`; MCP style tools do this |
| Legend text is awkward through COM | Set worksheet Long Names, rebuild with `legend -r`, then position |
| Legend overlaps a dense plot with no clear corner | Handled automatically: the legend is moved outside the frame (right of the plot, vertically centered) with the plot shrunk — `place_legend_avoiding_data` returns `outside-right`. Origin clamps `legend.left` back inside the frame on the first assignment, so the tool sets it TWICE (with `legend.attach = 1`); the second assignment escapes the clamp |
| `expGraph` needs a directory `path:=` + `filename:=` and `overwrite:=replace` (a full file path or missing args opens a dialog) | `export_graph` (expGraph, ~1200px wide by default, or `sized=True` with `tr1.unit:=2 tr1.width:=` pixels) writes the file directly with no clipboard |
| Graphic-object arrowheads use `arrowEndShape`/`arrowBeginShape` (1=filled, 2=chevron), not `arrowEnd`/`arrowEndType` | `annotate(kind="arrow")` draws the line and sets `arrowEndShape`; the begin/end length/width are `arrowEndLength`/`arrowEndWidth` |
| Save a graph template with `save -t <window> <fullpath.otpu>` (or `-tj` for `.otp`) — both window and full path+extension are required or a dialog opens | `save_graph_template` supplies both, so it never opens a dialog |
| Colormap palettes load via `layer.cmap.load(<name>.pal); layer.cmap.updateScale()` (the `()` matters) | `colormap` wraps this |
| Fit statistics can reset after `nlend` | Read statistics before ending the nonlinear fit session |
| Error-bar plots appear as separate entries in `DataPlots` | MCP styling tools detect them via the column's Y-Error designation and only color-match them — symbol/line commands would redraw error bars as connected lines |
| Error-bar `set -erw`/`-erwc` use POINTS, unlike the data line's `-w` (~200 units/pt) | Set `-erw` (error-bar line width) and `-erwc` (cap/whisker width) in points. Passing a `-w`-scale value (e.g. 550) into `-erw` makes bars explode — that was a units mistake, not an Origin bug. MCP tools set `-erw` in points and scale `-erwc` to the symbol size |
| `set <plot>` silently fails when the graph window is not active | Run `win -a <graph>` first |
| `[Book]Sheet!col(n).type = ...` is silently ignored | Activate the sheet and use `wks.col(n).type` instead |
| A graph loaded from a `.opju` project file can report zero data plots over COM, so per-curve styling/ungrouping used to silently no-op | The core per-curve/axis/frame tools (`set_plot_style`, `ungroup_plots`, `remove_plot`, `axis` range/scale/tick) now activate the page and re-acquire a fresh layer handle before each call; axis-range calls also read the value back and raise if it didn't change. If plots are still zero, the tool raises an actionable error — recreate the graph in-session (`create_graph`/`plotxy`) rather than editing the loaded one. Text/font/legend tools (`set_graph_font`, `set_legend`) still go through plain LabTalk and can silently no-op on a loaded graph — verify those in the exported image |
| A freshly created graph page can silently ignore or partially apply the FIRST styling/read/export command issued right after `create_graph`/`plotxy`/`add_plot_to_graph` (no exception, just no effect) | `create_graph`, `add_plot_to_graph`, and `ungroup_plots` now poll until the new plots enumerate and add a short settle before returning, so callers don't need to add their own delay |
| Combining multiple `-flag`s in ONE `set <ds> ...` command (e.g. `-c` + `-cf`, or `-k` + `-kf` + `-z`) silently corrupts the plot — colors reset to black, or the symbol blanks out | Send exactly ONE flag per `set` call. `set_plot_style` and `apply_publication_style` both do this now; never batch flags yourself via `run_labtalk` |
| `layer.x2.majorTicks` / `layer.y2.majorTicks` set to 0 wipes the NUMBER LABELS on ALL FOUR axes, not just the opposite side | Use `axis(op="tick", axis="top"/"right", tick_direction="none")`, which sets `layer.<ax>.ticks = 0` instead — never write `majorTicks` directly |
| `layer.x/y.*` properties (from/to/inc/thickness/ticks/majorLen) read back real values, but `layer.x2/y2.*` (opposite-axis) reads more often return Origin's missing-property sentinel (a tiny near-zero float, e.g. `-1.23456789e-300`) | `run_labtalk`'s `capture` translates that sentinel to the string `"missing"` instead of leaking the raw float |
| A never-saved project (no on-disk file yet) used to be treated the same as a failed autosave, blocking destructive ops (delete_graph, etc.) under a required autosave policy | Autosave now distinguishes "nothing on disk to protect" (proceeds) from "a real save attempt failed" (blocks when required) |
| Second-Y-axis (layer 2) right-axis styling looks unreachable via `layer.y.*` / `.bold` / `yl.text$` (those target layer 2's HIDDEN LEFT axis + title, not the visible right one) | The visible right axis is layer 2's **`y2`** (positional right). With the graph active and layer 2 selected (`win -a <graph>; page.active=2;`): right axis LINE + ticks via `layer.y2.color=color(...)`, right tick-NUMBER labels via `layer.y2.label.color=color(...)`, then `doc -uw;` — all verified live. `add_second_y_axis(..., color="r,g,b")` emits this. The right-axis TITLE is the separate `YR` text object — `yr.text$="\b(Label)"` (bold via markup, not `YR.bold=1`), `YR.color=color(...)`, `YR.fsize=...` |
| Removing ONE plot when the same dataset is plotted more than once: `layer -e <dataset>` (and `range r; delete r`) address by dataset NAME, so they remove EVERY copy | `remove_plot(plot_index=n)` now uses the COM DataPlot `Destroy()` addressed by index — removes only the indexed plot. (DataPlot has no `Remove`/`Delete`; `Destroy` is the only per-index route on Origin 2020) |
| Colormap/heatmap shows only ~8 discrete bands and looks impossible to make continuous — `rp.colormap.*` writes and `numColors=<n>; setLevels();` (no-arg) are silent no-ops | Raise the level count with `layer.cmap.numMajorLevels=<n>; layer.cmap.setLevels(1); layer.cmap.updateScale();` (the `setLevels(1)` arg + `updateScale()` are the missing steps). `colormap(graph_name, levels=32..64)` emits this and read-back-verifies `layer.cmap.numColors`. NOTE: applying a `palette` DOES recolor the map even when the exported PNG byte-size is unchanged (byte-compare is misleading — inspect pixels) |
| Layer-2 connecting-LINE width via `set <ds> -w <n>` looked non-deterministic ("flaky") | It is deterministic WITH the protocol: `win -a <graph>; page.active=2; set <ds> -w <n>; doc -uw;` + a ~2s settle (5/5 repeats identical live). It SATURATES around `-w 8` (w=8 and w=20 render identically), so to exceed that thickness plot a scaled copy of the data on layer 1 and draw it thick with `set_plot_style(line_width=...)` |
| Bar/column FILL PATTERN (hatch) is NOT scriptable on Origin 2020 | HARD LIMITATION. Every route no-ops (verified live, byte-identical + zero rendered hatch): `set <ds> -pfp/-pfw/-pfc` (the correct documented flags), indexed patterns `-pfpd/-pfpi`, the range tree `rp.pattern.*`, and the plot-object tree `layer.plotN.pattern.*`. Apply hatching in PPT/Illustrator, or use solid fills distinguished by color |
| Plot TRANSPARENCY / alpha is NOT scriptable on Origin 2020 | HARD LIMITATION. `set <ds> -paap/-paal/-paas`, `rr.transparency`, `layer.transparency`, `page.transparency` all return success but no-op (verified live, byte-identical). Use fully-opaque colors, or composite the alpha in PPT. (Unconfirmed future candidate: the `originpro` Python `Plot.transparency(t)` method — not installed in this env to test) |
| `layer.y.inc` (tick spacing) on a heatmap/2D y-axis can appear to no-op | It works — the write needs the graph window ACTIVE first (`win -a <graph>; page.active=<n>;`) and a `doc -uw;` refresh afterward, same as other per-layer LabTalk |
| A colorbar (color scale) on a heatmap/colormap layer can overflow past the page's right edge, and the colorbar object itself is not directly addressable (`ColorScale.left`/`.width` read 0) | The colorbar is ANCHORED to the plot layer — shrink/reposition the LAYER instead: keep the frame's right edge (`layer.left + layer.width`) at roughly ≤72% of the page, e.g. `layer.left=15; layer.width=56; layer.top=13; layer.height=64; doc -uw;` |
| Tick positions are anchored at multiples of `inc` from 0 with no exposed anchor property (`layer.x.anchor` is a no-op) | Workaround: shift the data by an offset and relabel with `layer.x.label.formula$="x+1"` (or similar) so the shifted ticks display the original values |
| Axis label text rejects `\x()` markup (only `\b \i \u \+ \- \g \f` are supported — see `validate_text_escapes`) | Use the actual unicode character (×, Å, μ, °, …) directly in the label string instead of an escape sequence |
| matplotlib nudges an in-corner "0" tick label (e.g. the x-axis 0) sideways so it doesn't collide with the y-axis "0"; Origin has no per-tick-label position/hide property (`nSpecialTicks`/`label.skip`/`label.first` are inconclusive or broken; "Special Ticks" is GUI-only) | Usually moot: Origin places the x-axis "0" BELOW the frame and the y-axis "0" to its LEFT, separated by the corner, so they don't overlap the way matplotlib's in-corner labels do — no nudge is typically needed |

## Pre-Export Checklist

- Axis labels include units and are readable after export
- Tick labels are readable and not overcrowded
- Line widths and symbol sizes survive downscaling
- Colors are distinguishable in colorblind and grayscale contexts
- Error bars are included when available
- Legend or direct labels do not cover data
- Grid, background, and frame match the target venue
- Exported file exists and is large enough to be a real image

## Customize This Skill

This file is not meant to be the one true visual style. Make your own version for:

- Journal presets such as Nature, ACS, IEEE, RSC, or thesis formats
- Lab color palettes and standard marker orders
- Common figure families such as XRD, Raman, IV curves, spectra, kinetics, bar charts, and dose response
- Preferred fonts, export folders, panel labels, and annotation conventions

Keep the workflow generic, but let the final taste be yours.
