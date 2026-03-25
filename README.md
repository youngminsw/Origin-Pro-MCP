# Origin Pro MCP Server

An MCP (Model Context Protocol) server that enables AI assistants like Claude to control **OriginLab Origin Pro** via COM automation. All operations are reflected in Origin's GUI in real-time — you watch as the AI creates worksheets, plots graphs, and styles figures.

## What Can It Do?

- **Worksheet Management** — Create workbooks, read/write data, import CSV files
- **Graph Creation** — Scatter, line, line+symbol, bar, histogram, box, contour, pie, bubble plots
- **Plot Styling** — Colors, symbols, line width, publication-ready formatting in one call
- **Curve Fitting** — Linear, polynomial, exponential, Gaussian, Lorentz, Voigt, and more
- **Project Management** — New/save/load projects, export all graphs
- **LabTalk Scripting** — Direct LabTalk execution for advanced operations

## Quick Start

### 1. Prerequisites

- **Windows** with **Origin Pro 2020+** installed and running
- **Python 3.10+** (Windows Python, not WSL)
- **pip packages**: `pip install -r requirements.txt`

### 2. Install

```bash
git clone https://github.com/youngminsw/Origin-Pro-MCP.git
cd origin-mcp-server
pip install -r requirements.txt
```

### 3. Configure Claude Code

Add to your Claude Code MCP settings (`~/.claude/settings.json` or project settings):

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "python",
      "args": ["C:\\path\\to\\origin-mcp-server\\server.py"],
      "cwd": "C:\\path\\to\\origin-mcp-server"
    }
  }
}
```

> **Note**: Use your Windows Python path. If Claude Code runs in WSL, use the full Windows path like `C:\\Users\\yourname\\origin-mcp-server\\server.py`.

### 4. Start Origin Pro

Open Origin Pro before starting the server. The server attaches to the running instance via `Origin.ApplicationSI`.

### 5. Use It

Just ask Claude to work with Origin:

```
"Create a scatter plot from this data: x=[1,2,3,4,5], y=[2.1,4.0,5.9,8.1,10.0]"
"Apply publication styling to Fig1 with axis labels Temperature (K) and Absorbance (a.u.)"
"Fit a Gaussian to the data in Book1"
"Export all graphs to C:\Users\me\figures\"
```

## Architecture

```
Claude Code (WSL or Windows)
    |  stdio (MCP protocol)
    v
MCP Server (Windows Python + win32com)
    |  COM automation
    v
Origin Pro (GUI visible in real-time)
```

## Available Tools (23 total)

### Project Management
| Tool | Description |
|------|-------------|
| `new_project` | Create new empty Origin project |
| `save_project` | Save project to .opju file |
| `load_project` | Open existing .opj/.opju file |

### Data
| Tool | Description |
|------|-------------|
| `create_worksheet` | Create new workbook |
| `set_worksheet_data` | Write column data (JSON arrays) |
| `get_worksheet_data` | Read worksheet data as JSON |
| `import_csv_to_worksheet` | Import CSV/text file |
| `list_worksheets` | List all open workbooks and graphs |

### Graphing
| Tool | Description |
|------|-------------|
| `create_graph` | Create graph (scatter, line, line+symbol, bar, etc.) |
| `add_plot_to_graph` | Add another dataset to existing graph |
| `set_axis_labels` | Set X/Y axis labels and title |
| `set_axis_range` | Set axis min/max values |
| `export_graph` | Export graph to PNG/JPG image |
| `export_all_graphs` | Export every graph in the project |

### Styling
| Tool | Description |
|------|-------------|
| `apply_publication_style` | **One-call publication formatting** (recommended) |
| `set_plot_style` | Set color, symbol, line width per plot |
| `set_graph_font` | Set font family and size |
| `set_legend` | Configure legend text and position |
| `set_tick_style` | Set tick direction and length |

### Analysis
| Tool | Description |
|------|-------------|
| `curve_fit` | Curve fitting with R², SSR statistics |
| `list_fitting_functions` | Show available fit functions |

### Advanced
| Tool | Description |
|------|-------------|
| `run_labtalk` | Execute any LabTalk script directly |
| `get_labtalk_variable` | Read a LabTalk variable value |

## Example: Publication-Quality Figure

```python
# 1. Start fresh
new_project()

# 2. Create data
create_worksheet(book_name="Data")
set_worksheet_data(
    book_name="Data", sheet_name="Sheet1",
    columns="[[300,350,400,450,500,550,600],[0.12,0.35,0.89,1.62,1.95,1.92,1.25],[0.08,0.22,0.62,1.35,1.88,1.90,1.12]]",
    column_names="Temperature,Pristine,Annealed"
)

# 3. Create line+symbol graph
create_graph(graph_name="Fig1", data_book="Data", data_sheet="Sheet1",
             x_col=1, y_col=2, plot_type="line+symbol")
add_plot_to_graph(graph_name="Fig1", data_book="Data", data_sheet="Sheet1",
                  x_col=1, y_col=3, plot_type="line+symbol")

# 4. One call does everything: colors, fonts, ticks, frame, legend
apply_publication_style(
    graph_name="Fig1",
    x_label="Temperature (K)",
    y_label="Absorbance (a.u.)",
    x_min=280, x_max=620, y_min=0, y_max=2.2,
    legend_entries="Pristine,Annealed",
    legend_position="top-right"
)

# 5. Export
export_graph(graph_name="Fig1", file_path="C:\\Users\\me\\fig1.png")
```

Result: A publication-ready figure with bold Arial labels, colorblind-safe colors (blue circles + red triangles), filled symbols, solid lines, inward ticks, closed frame, and positioned legend.

## Claude Code Skill: Publication Figure

This repo includes a **skill file** (`skills/publication-figure.md`) that teaches Claude how to create journal-quality figures step by step. To use it:

1. Copy `skills/publication-figure.md` to your Claude Code skills directory
2. When you ask Claude to "make a publication figure", it will automatically follow the skill's workflow:
   - Ask about data source, figure type, target journal
   - Use colorblind-safe color palette (blue → red → green → orange → purple → cyan)
   - Apply proper typography (Arial, bold, correct sizes)
   - Follow a pre-export checklist

The skill also documents **Origin 2020 COM quirks** — what works, what doesn't, and tested workarounds. This is invaluable if you need to customize beyond `apply_publication_style`.

### Customizing the Skill for Your Style

The included skill is a starting template. You should **customize it to match your lab's or journal's requirements**:

- **Font**: Change from Arial to your journal's preferred font (e.g., Helvetica, Times New Roman)
- **Font sizes**: Adjust axis title/tick label/legend sizes to match your journal's figure guidelines
- **Color palette**: Replace the default colorblind-safe palette with your group's standard colors
- **Default export path**: Set to your working directory
- **Figure recipes**: Add templates for your common figure types (XRD patterns, IV curves, etc.)
- **Journal presets**: Add specific formatting rules for your target journals (Nature, ACS, RSC, etc.)

Copy `skills/publication-figure.md` to your project and edit freely — it's meant to be a starting point, not a rigid template.

### Key Origin 2020 COM Quirks (documented in skill)

| Issue | Workaround |
|-------|-----------|
| Bold axis titles (`xb.bold`) doesn't exist | Use `\b(text)` markup in `xb.text$` |
| `legend.text$` doesn't support multiline via COM | Set column Long Names, then `legend -r` |
| `legend.x/y` uses data coordinates, not % | Calculate from `layer.x.from/to` |
| `%C` notation fails via COM | Use actual plot names from `DataPlots` |
| `expGraph` doesn't produce files via COM | Use `CopyPage` + Pillow clipboard |
| `nlr.r2` returns 0 after `nlend` | Read statistics BEFORE `nlend` |
| Plot styling commands can conflict | Add 0.2s delay between `set` commands |

## Supported Plot Types

| Type | Description |
|------|-------------|
| `scatter` | Scatter plot (symbols only) |
| `line` | Line plot (no symbols) |
| `line+symbol` | Line with symbols (recommended for publications) |
| `column` | Vertical bar chart |
| `bar` | Horizontal bar chart |
| `area` | Area plot |
| `histogram` | Histogram |
| `box` | Box plot |
| `contour` | Contour plot |
| `pie` | Pie chart |
| `bubble` | Bubble chart |

## Supported Fitting Functions

| Category | Functions |
|----------|----------|
| Linear | `line` |
| Polynomial | `poly2`, `poly3`, `poly4`, `poly5` |
| Exponential | `exp1`, `exp2`, `expgrow1`, `expdecay1` |
| Peak | `gauss`, `lorentz`, `voigt` |
| Growth/Sigmoidal | `boltzmann`, `hill`, `logistic`, `lognormal` |
| Other | `power`, `sine` |

## Color Palette

Default colorblind-safe order used by `apply_publication_style`:

| Order | Color | Best for |
|-------|-------|----------|
| 1st | Blue | Primary dataset |
| 2nd | Red | Comparison dataset |
| 3rd | Green | Third dataset |
| 4th | Orange | Fourth dataset |
| 5th | Purple | Fifth dataset |
| 6th | Cyan | Sixth dataset |

> **Tip**: Never use red + green as the only two colors — colorblind users cannot distinguish them. Use blue + red or blue + orange instead.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Origin.ApplicationSI" error | Make sure Origin Pro is running before starting the server |
| Tools timeout | Origin may be showing a dialog — check the Origin window |
| Export returns empty | Increase the clipboard wait time; check if Origin window is minimized |
| Legend missing after styling | Legend uses data coordinates — verify axis range is set before positioning |
| Symbols appear hollow | Do NOT use `set -d` flag (it's for dash patterns, not fill) |

## License

MIT
