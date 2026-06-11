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
- Supported plot types: scatter, line, line+symbol, column, bar, area,
  pie, histogram (single Y column), and contour (needs `z_col` — XYZ).
  Box plots and true 3D (scatter/surface, OpenGL) are not yet reliable
  through `create_graph`; flag the limitation instead of shipping a
  broken or empty plot.

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

## Origin COM Notes

These COM behaviors were observed while testing on Origin Pro 2020. Other Origin versions may behave differently until verified:

| Issue | Practical workaround |
| --- | --- |
| Bold axis title properties are unreliable | Use Origin text markup such as `\b(Label)` through typed tools |
| `%C` plot shortcuts can fail through COM | Use plot names from `FindGraphLayer().DataPlots`; MCP style tools do this |
| Legend text is awkward through COM | Set worksheet Long Names, rebuild with `legend -r`, then position |
| Legend coordinates use data units | Set axis range before final legend placement; MCP tools keep the legend box inside the frame automatically |
| `expGraph` may not write files reliably | Use `export_graph`, which copies the rendered page and saves via Pillow |
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
