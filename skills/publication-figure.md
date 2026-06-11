# Publication Figure

Create paper-grade figures in Origin Pro with the `origin-pro` MCP server. Treat this skill as a generic starting point: copy it, tune it, and keep your own lab or journal presets close to the projects where you use them.

## When To Use

Use this when the user asks for a manuscript figure, publication-quality graph, journal-ready plot, reference-paper style, or polished Origin export.

## Requirements

- Windows with a licensed Origin/OriginPro installation
- `origin-pro` MCP launched by the MCP client through Windows Python
- Prefer typed MCP tools over raw LabTalk
- Use `run_labtalk` for advanced graph and analysis commands; destructive and file-writing commands are blocked
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
  matched to the symbol size. Keep tick labels
  uncrowded (~5 major intervals); the one-call tool handles spacing.
- **The legend must never overlap the data.** `apply_publication_style`,
  `set_legend`, and the fit-curve path auto-place the legend in the
  emptiest corner (the requested `legend_position` is only a preference,
  overridden when data sits there). If every corner is crowded, widen the
  axis range to open a gap rather than letting the legend cover points.
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
  take `apply_color_map`/`set_colormap_levels`. Annotate with
  `add_text_annotation`, `add_line`, and `add_arrow`; save reusable
  layouts with `save_graph_template`.

## 3D and Colormap Design Defaults

Same spirit as the 2D defaults (clean, honest, colorblind-safe, units
everywhere, readable bold labels), adapted to surfaces, contours,
heatmaps, and 3D scatter.

- **The colormap is the 3D palette — choose it as carefully as 2D colors.**
  Use a perceptually-uniform, colorblind- and grayscale-safe map. Good
  built-ins that load by single-word name with `apply_color_map`:
  `Heatmap4ColorBlind` and `BlueGreenYellow` (sequential magnitude),
  `RedWhiteBlue` (diverging / signed data centered on zero), `GrayScale`
  (print-safe). Avoid plain `Rainbow`/`Jet` for quantitative data — they
  invent false boundaries and fail in grayscale. (Origin's
  `Rainbow Isolum`/`Rainbow Balanced` are better rainbows but their names
  contain spaces, so apply them in the GUI for now.)
- **Sequential vs. diverging:** sequential map for one-directional
  magnitude (0→max); diverging map with the midpoint at 0 for signed data
  (deviation, difference, charge).
- **Always show a labeled color scale with units** — it is the legend of a
  colormap plot. Never ship a heatmap/contour without it.
- **Set the Z range honestly** with `set_colormap_levels(z_min, z_max)` to
  the real data range; don't clip features into saturation just to boost
  contrast, and state the range.
- **Label every axis with units**, including Z on 3D plots, in the same
  bold readable font as the 2D figures (`set_axis_labels`; for 3D the Z
  title and tick fonts follow the same sizing).
- **Prefer a 2D contour or heatmap over a 3D surface when exact values
  matter** — a top-down colormap is easier to read off than a tilted
  surface. Use the 3D surface for shape/intuition, the contour for
  quantitative reading; a contour-projected surface gives both.
- **Keep surfaces uncluttered:** a light mesh or a smooth color-mapped
  surface reads better than a dense wireframe; pick one clear viewing
  angle and keep it consistent across a figure set.
- **3D scatter:** make symbols large enough to read at print size (the
  same "bigger symbols" lesson as 2D) and rely on Z color/height, not tiny
  dots, to carry the third dimension.
- **Match the rest of the figure set:** same fonts, same export pixel size
  (`export_graph_sized`), same labeling conventions as the 2D panels so a
  mixed figure looks like one family.
- Export once, inspect, then adjust colormap, Z range, viewing angle, and
  label sizes — exactly the 2D "export and inspect" loop.

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
import_csv_to_worksheet(file_path="C:\\Users\\name\\data.csv")
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
set_axis_range(graph_name="Fig1", x_min=280, x_max=620, y_min=0, y_max=2.2)
set_legend(graph_name="Fig1", visible=True, position="top-right", entries="Pristine,Annealed")
```

`set_plot_style` reference (these are the only accepted values):

- `plot_index` is 1-based in the order series were added; error-bar plots
  are not counted.
- `color`: black, red, green, blue, cyan, magenta, yellow, orange,
  purple, gray.
- `symbol_shape`: 0=auto, 1=square, 2=circle, 3=triangle-up, 4=diamond,
  5=triangle-down, 6=hexagon.
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

The server exposes 56 tools. Reach for these when a figure needs more
than a styled XY plot.

### Surfaces, contours, heatmaps

3D / colormapped figures are built from a **matrix**. Grid scattered XYZ
first, then plot:

```text
worksheet_to_matrix(data_book="D", data_sheet="Sheet1", x_col=1, y_col=2, z_col=3)
create_matrix_plot(matrix_book="Matrix", plot_type="surface")   # or contour, heatmap, image
apply_color_map(graph_name="Graph1", palette="Fire")            # Fire, Rainbow, GrayScale, Maple, ...
set_colormap_levels(graph_name="Graph1", z_min=0, z_max=1)
```

For a quick 2D contour or a 3D scatter straight from XYZ columns, use
`create_graph(..., plot_type="contour", z_col=N)` or `plot_type="3d_scatter"`.

### Statistical / distribution figures

- `create_graph(..., plot_type="box")` and `plot_type="histogram"` (single Y column).
- `column_statistics`, `compare_means` (two-sample t-test), `frequency_count`
  return JSON — put the numbers in the caption, don't just draw bars.

### Signal / spectra workflows

`smooth`, `differentiate`, `integrate`, `interpolate`, `fft`, `find_peaks`
write result columns or return JSON. Typical Raman/XRD flow: smooth →
find_peaks → `curve_fit(plot_on_graph=...)` → annotate the peaks.

### Multi-panel and axis control

- `set_axis_scale(graph, axis, "log10")` for decades-spanning data.
- `add_second_y_axis` / `add_layer` for dual-axis or stacked panels.
- `add_reference_line` (threshold/baseline), `add_line`, `add_arrow` (callouts),
  `add_text_annotation` (labels at data coordinates).

### Worksheet prep and IO

`set_column_formula`, `sort_worksheet`, `set_column_properties` (units/long
name), `add_columns`/`delete_columns`, `transpose_worksheet`, `import_excel`,
`export_worksheet`.

### Reuse and sized export

- `save_graph_template(graph, path)` captures a finished layout as `.otpu`
  to reuse across figures.
- `export_graph_sized(graph, path, width=1600)` exports at an exact pixel
  width (vs. `export_graph`, which follows the Origin page size).

## Origin COM Notes

These COM behaviors were observed while testing on Origin Pro 2020. Other Origin versions may behave differently until verified:

| Issue | Practical workaround |
| --- | --- |
| Bold axis title properties are unreliable | Use Origin text markup such as `\b(Label)` through typed tools |
| `%C` plot shortcuts can fail through COM | Use plot names from `FindGraphLayer().DataPlots`; MCP style tools do this |
| Legend text is awkward through COM | Set worksheet Long Names, rebuild with `legend -r`, then position |
| Legend coordinates use data units | Set axis range before final legend placement; MCP tools keep the legend box inside the frame automatically |
| `expGraph` needs a directory `path:=` + `filename:=` and `overwrite:=replace` (a full file path or missing args opens a dialog) | `export_graph` (clipboard, page size) and `export_graph_sized` (expGraph, `tr1.unit:=2 tr1.width:=` pixels) both handle this correctly |
| Graphic-object arrowheads use `arrowEndShape`/`arrowBeginShape` (1=filled, 2=chevron), not `arrowEnd`/`arrowEndType` | `add_arrow` draws the line and sets `arrowEndShape`; the begin/end length/width are `arrowEndLength`/`arrowEndWidth` |
| Save a graph template with `save -t <window> <fullpath.otpu>` (or `-tj` for `.otp`) — both window and full path+extension are required or a dialog opens | `save_graph_template` supplies both, so it never opens a dialog |
| Colormap palettes load via `layer.cmap.load(<name>.pal); layer.cmap.updateScale()` (the `()` matters) | `apply_color_map` / `set_colormap_levels` wrap this |
| Fit statistics can reset after `nlend` | Read statistics before ending the nonlinear fit session |
| Error-bar plots appear as separate entries in `DataPlots` | MCP styling tools detect them via the column's Y-Error designation and only color-match them — symbol/line commands would redraw error bars as connected lines |
| Error-bar `set -erw`/`-erwc` use POINTS, unlike the data line's `-w` (~200 units/pt) | Set `-erw` (error-bar line width) and `-erwc` (cap/whisker width) in points. Passing a `-w`-scale value (e.g. 550) into `-erw` makes bars explode — that was a units mistake, not an Origin bug. MCP tools set `-erw` in points and scale `-erwc` to the symbol size |
| `set <plot>` silently fails when the graph window is not active | Run `win -a <graph>` first |
| `[Book]Sheet!col(n).type = ...` is silently ignored | Activate the sheet and use `wks.col(n).type` instead |

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
