# Publication Figure

Create publication-quality figures in Origin Pro for journal submission. This skill works together with the `origin-pro` MCP server.

## When to Use

Use when the user asks to:
- Create a figure for a paper, manuscript, or publication
- Make a "publication-quality" or "journal-ready" graph
- Style an existing Origin graph for submission
- Reproduce a figure from a reference paper

## Prerequisites

- Origin Pro must be running on Windows
- The `origin-pro` MCP server must be connected
- Available MCP tools: `create_worksheet`, `set_worksheet_data`, `create_graph`, `add_plot_to_graph`, `set_axis_labels`, `set_axis_range`, `set_plot_style`, `set_graph_font`, `set_legend`, `set_tick_style`, `export_graph`, `run_labtalk`, `curve_fit`

## Before You Start: Ask the User

1. **Data source** — CSV file path? Or provide data directly?
2. **Figure type** — scatter, line+symbol, bar, histogram?
3. **How many datasets** on one graph?
4. **Target journal** — Nature, Science, ACS, or general?
5. **Export path** — where to save the PNG?

If the user doesn't specify, default to: Nature-style, line+symbol, export to desktop or a user-specified directory.

---

## Verified Origin 2020 COM Settings

These settings have been tested and confirmed working via COM automation.

### What WORKS

| LabTalk Command | Effect |
|----------------|--------|
| `xb.text$ = "Label (unit)";` | Set X axis title |
| `yl.text$ = "Label (unit)";` | Set Y axis title |
| `xb.fsize = 24;` | Set X axis title font size |
| `yl.fsize = 24;` | Set Y axis title font size |
| `xb.font$ = "Arial";` | Set X axis title font |
| `layer.x.label.pt = 20;` | Set X tick label font size |
| `layer.y.label.pt = 20;` | Set Y tick label font size |
| `layer.x.from = 0; layer.x.to = 100;` | Set X axis range |
| `layer.y.from = 0; layer.y.to = 50;` | Set Y axis range |
| `layer.x.opposite = 1;` | Show top axis (closed frame) |
| `layer.y.opposite = 1;` | Show right axis (closed frame) |
| `layer.x.thickness = 2;` | Set X axis line thickness |
| `layer.y.thickness = 2;` | Set Y axis line thickness |
| `layer.x.grid = 0;` | Remove X gridlines |
| `layer.y.grid = 0;` | Remove Y gridlines |
| `layer.x.minorGrid = 0;` | Remove X minor gridlines |
| `legend.fsize = 18;` | Set legend font size |
| `legend.font$ = "Arial";` | Set legend font |
| `legend.x = 85; legend.y = 15;` | Move legend position |
| `layer.x.label.bold = 1;` | Bold tick labels |
| `range r1 = [Graph]Layer1!1; r1.color = 4;` | Set plot color (range method) |

### What DOES NOT WORK via COM (and workarounds)

| Command | Issue | Workaround |
|---------|-------|------------|
| `xb.bold = 1;` | Property doesn't exist | Use `\b(text)` markup: `xb.text$ = "\\b(Label)";` |
| `set %C1 -c 4;` | `%C` notation fails via COM | Use actual plot name: `set Data_B -c 4;` (get names via `FindGraphLayer().DataPlots`) |
| `set %C1 -k 2;` | Same `%C` issue | Use: `set Data_B -k 2; set Data_B -kf 1;` |
| `layer.x.minor` via `o.Execute()` | May fail depending on context | Use `FindGraphLayer().Execute()` instead |
| `layer.x.majorLen` via `o.Execute()` | Same context issue | Use `FindGraphLayer().Execute()` instead |
| `expGraph` (LabTalk) | Does not produce files via COM | Use `CopyPage` + Pillow clipboard |
| `nlr.r2` after `nlend` | Returns 0 | Must read BEFORE `nlend` |

### Reliable execution method

For graph layer properties (`layer.x.*`, `layer.y.*`), use `FindGraphLayer().Execute()` — this is the most reliable method. The MCP tool `apply_publication_style` uses this internally.

For plot styling (`set PLOTNAME -c/-k/-w/-z`), use the actual plot name from `FindGraphLayer().DataPlots`. The MCP tool `set_plot_style` handles this automatically.

### Export Method

Our MCP uses **CopyPage + clipboard** (not LabTalk expGraph). This means:
- Export captures what Origin renders on screen
- Font sizes need to be **visually large** in the Origin window
- Exported resolution depends on Origin's display, not a DPI setting

---

## Font Size Guide

Since export is clipboard-based, use these sizes for readable output:

| Element | Size (pt) | Notes |
|---------|-----------|-------|
| Axis titles | **24** | Clear and readable |
| Tick labels | **20** | Number labels on axes |
| Legend text | **18** | Dataset names |
| Panel label (a, b, c) | **22 bold** | If multi-panel |
| Annotation text | **16-18** | In-graph notes |

Font: **Arial** everywhere. No exceptions.

---

## Standard Workflow

### 1. Clean Start

```
new_project()
```

### 2. Load Data

**From user-provided arrays:**
```
create_worksheet(book_name="Data")
set_worksheet_data(
    book_name="Data", sheet_name="Sheet1",
    columns="[[x1,x2,...],[y1,y2,...],[err1,err2,...]]",
    column_names="X,Y,Error"
)
```

**From CSV file:**
```
import_csv_to_worksheet(file_path="C:\\Users\\path\\data.csv")
```

### 3. Create Graph

```
create_graph(
    graph_name="Fig1",
    data_book="Data", data_sheet="Sheet1",
    x_col=1, y_col=2,
    plot_type="line+symbol",
    y_error_col=3
)
```

For multiple datasets, add more:
```
add_plot_to_graph(
    graph_name="Fig1",
    data_book="Data", data_sheet="Sheet1",
    x_col=1, y_col=4,
    plot_type="line+symbol",
    y_error_col=5
)
```

### 4. Style: Apply ALL at once (RECOMMENDED)

**Use `apply_publication_style` — one call does everything:**

```
apply_publication_style(
    graph_name="Fig1",
    x_label="Temperature (K)",
    y_label="Absorbance (a.u.)",
    x_min=300, x_max=600,
    y_min=0, y_max=2.5,
    legend_entries="Pristine,Annealed",
    legend_position="top-right"
)
```

This single call applies:
- Bold axis titles (28pt Arial)
- Bold tick labels (22pt)
- Colorblind-safe colors (auto: blue, red, green, orange...)
- 1.5pt line width, line connection on
- Distinct filled symbols (auto: ■ square, ● circle, ▲ triangle, ◆ diamond)
- Symbol size 10 (appropriate for publication)
- Inward ticks with minor ticks
- Closed frame (4 sides), 2pt thickness
- No grid lines
- Legend with 20pt Arial

### 4b. Style: Manual Typography (if fine-tuning needed)

Use `\b(...)` markup for **bold axis titles**. Use `FindGraphLayer().Execute()` for tick properties.

```
run_labtalk('win -a Fig1;')
run_labtalk('xb.text$ = "\\b(Temperature (K))"; xb.fsize = 28; xb.font$ = "Arial";')
run_labtalk('yl.text$ = "\\b(Absorbance (a.u.))"; yl.fsize = 28; yl.font$ = "Arial";')
run_labtalk('layer.x.label.pt = 22; layer.y.label.pt = 22;')
run_labtalk('layer.x.label.bold = 1; layer.y.label.bold = 1;')
```

Always include units in parentheses: `"Pressure (MPa)"`, `"Time (s)"`, `"Wavelength (nm)"`.

### 5. Style: Colors & Lines

Apply colorblind-safe colors in this order:

| Dataset | Color | Origin name |
|---------|-------|-------------|
| 1st | Blue | `"blue"` |
| 2nd | Red | `"red"` |
| 3rd | Green | `"green"` |
| 4th | Orange | `"orange"` |
| 5th | Purple | `"purple"` |
| 6th | Cyan | `"cyan"` |

**Recommended color combinations by number of datasets:**

| Datasets | Combination | Notes |
|----------|-------------|-------|
| 2 | blue + red | High contrast, classic |
| 2 | blue + orange | Colorblind-safe best pair |
| 3 | blue + red + green | Standard trio |
| 3 | blue + orange + purple | Colorblind-safe trio |
| 4 | blue + red + green + orange | Clear distinction |
| 5+ | blue + red + green + orange + purple + cyan | Full palette |

**NEVER** use red + green as the only two colors (colorblind users can't distinguish).

```
set_plot_style(graph_name="Fig1", plot_index=1, color="blue", line_width=2.5, symbol_size=12)
set_plot_style(graph_name="Fig1", plot_index=2, color="red", line_width=2.5, symbol_size=12)
```

### 6. Style: Axes & Frame

```
# Tight axis range — look at the data and choose wisely
set_axis_range(graph_name="Fig1", x_min=0, x_max=300, y_min=0, y_max=4.5)

# Ticks: inward, minor ticks visible
# Note: tick length and minor tick properties are read-only in Origin 2020 COM
# Only tick direction can be reliably set. Minor ticks are shown by default.
set_tick_style(graph_name="Fig1", tick_direction="in")

# Closed frame (4 sides) + thick axis lines
run_labtalk("layer.x.opposite = 1; layer.y.opposite = 1;")
run_labtalk("layer.x.thickness = 2; layer.y.thickness = 2;")

# REMOVE all grid lines (critical!)
run_labtalk("layer.x.grid = 0; layer.y.grid = 0;")
run_labtalk("layer.x.minorGrid = 0; layer.y.minorGrid = 0;")
```

### 7. Style: Legend

```
set_legend(graph_name="Fig1", visible=True, position="top-left", entries="Sample A,Sample B")
run_labtalk('legend.fsize = 18; legend.font$ = "Arial";')
```

Choose legend position to avoid overlapping data:
- **top-left**: if data rises from left to right
- **top-right**: if data falls from left to right
- **bottom-right**: if data is high on the left

### 8. Export

```
export_graph(graph_name="Fig1", file_path="C:\\Users\\yourname\\figures\\fig1.png")
```

### 9. Show Result & Iterate

After export, tell the user:
- File path
- What was applied
- Ask: **"조절할 부분이 있나요? (글자 크기, 색상, 축 범위 등)"**

---

## Common Figure Recipes

### Recipe A: Scatter with Linear Fit

```
create_graph(..., plot_type="scatter")
curve_fit(data_book="Data", data_sheet="Sheet1", x_col=1, y_col=2, function="line")
# Show R² value to user from fit result
# Add fit line annotation via LabTalk if needed
```

### Recipe B: Multi-dataset Comparison

```
create_graph(..., plot_type="line+symbol", y_error_col=3)
add_plot_to_graph(..., y_col=4, y_error_col=5)
add_plot_to_graph(..., y_col=6, y_error_col=7)
# Use blue, red, green for 3 datasets
```

### Recipe C: Bar Chart

```
create_graph(..., plot_type="column")
set_plot_style(..., color="blue")
```

### Recipe D: Before/After Comparison (2 panels)

Create two separate graphs, export both:
```
create_graph(graph_name="Fig1a", ...)
create_graph(graph_name="Fig1b", ...)
# Apply same styling to both
export_graph(graph_name="Fig1a", file_path="...\\fig1a.png")
export_graph(graph_name="Fig1b", file_path="...\\fig1b.png")
# Tell user to combine in PowerPoint/Illustrator with (a), (b) labels
```

---

## Pre-Export Checklist

Before exporting, verify every item:

- [ ] **Font**: Arial everywhere, no default fonts remaining
- [ ] **Axis titles**: 24 pt bold, with units in parentheses — use `\b(...)` markup
- [ ] **Tick labels**: 20 pt bold — use `layer.x.label.bold = 1;`
- [ ] **Legend**: 18 pt, positioned away from data, no border
- [ ] **Lines**: 2-2.5 pt width
- [ ] **Symbols**: size 8-10, distinct shapes per dataset (■●▲◆), filled
- [ ] **Colors**: blue/red/green/orange (colorblind-safe order)
- [ ] **Ticks**: inward direction, minor ticks visible
- [ ] **Grid lines**: ALL removed
- [ ] **Frame**: closed (4 sides), thickness 2 pt
- [ ] **Axis range**: tight, no excessive white space
- [ ] **Background**: white, no shadows, no 3D effects
- [ ] **Error bars**: included if user has error data
