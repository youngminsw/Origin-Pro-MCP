# Origin Pro MCP Server

An MCP (Model Context Protocol) server that enables AI assistants like Claude to control **OriginLab Origin Pro** via COM automation. All operations are reflected in Origin's GUI in real-time — you watch as the AI creates worksheets, plots graphs, and styles figures.

## What Can It Do?

- **Worksheet Management** — Create workbooks, read/write data, import CSV files
- **Graph Creation** — Scatter, line, line+symbol, bar, histogram, box, contour, pie, bubble plots
- **Plot Styling** — Colors, symbols, line width, publication-ready formatting in one call
- **Curve Fitting** — Linear, polynomial, exponential, Gaussian, Lorentz, Voigt, and more
- **Project Management** — New/save/load projects, export all graphs
- **LabTalk Scripting** — Direct LabTalk execution with destructive/file-overwrite commands blocked

This MCP server is intentionally **Windows-runtime-only**. The AI agent or MCP client can run from Windows or WSL, but the MCP server process that talks to Origin must be launched with Windows Python and `pywin32`. WSL/Linux can edit the project and run non-COM unit tests, but cannot directly control Origin COM.

## Quick Start

### 1. Prerequisites

- **Windows** with a licensed Origin/OriginPro installation that exposes the Automation Server
- **Python 3.10+** (Windows Python, not WSL)

Tested environment: **Origin Pro 2020**. Other Origin/OriginPro versions may work if they expose compatible COM Automation Server and LabTalk behavior, but they are not verified by this project yet.

### 2. Install & Configure

**Option A: uvx (recommended — zero setup)**

No manual install needed. The MCP client launches the server for you. Just add this to your Claude Code MCP settings when Claude Code is running on Windows:

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "uvx",
      "args": ["--quiet", "origin-pro-mcp"]
    }
  }
}
```

`uvx` automatically downloads and runs the server in an isolated environment. Nothing else to install. The `--quiet` flag keeps first-run dependency messages out of your MCP client logs.

If Claude Code or another MCP client is running inside WSL, launch the same Windows server by calling Windows `uvx.exe` directly:

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "uvx.exe",
      "args": ["--quiet", "origin-pro-mcp"]
    }
  }
}
```

For a local checkout before publishing/installing, point Windows `uvx` at the Windows path of the repo:

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "uvx.exe",
      "args": ["--quiet", "--refresh", "--from", "D:\\04.Agent OS\\Origin-Pro-MCP", "origin-pro-mcp"]
    }
  }
}
```

Keep the command and args as separate JSON array entries. That avoids quoting problems when a Windows path contains spaces. If WSL cannot find `uvx.exe`, set `command` to the full WSL path for the Windows executable, for example `/mnt/c/Users/YOU/.local/bin/uvx.exe`.

**Option B: pip install from PyPI**

```bash
pip install origin-pro-mcp
```

Then configure Claude Code:

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "origin-pro-mcp"
    }
  }
}
```

**Option C: Clone and run directly**

```bash
git clone https://github.com/youngminsw/Origin-Pro-MCP.git
cd Origin-Pro-MCP
pip install -e .
```

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "origin-pro-mcp"
    }
  }
}
```

> **Note**: If Claude Code runs in WSL, make sure the `uvx` or `python` command points to your **Windows** Python, not WSL Python. Origin COM only works from Windows.

### 3. Origin Startup

You do not need to start Python or Origin manually. The MCP client starts `origin-pro-mcp`, and the server connects to Origin through `Origin.ApplicationSI`. If Origin is already open, the server uses it; otherwise COM launches Origin and the server makes it visible.

### 4. Use It

Just ask Claude to work with Origin:

```
"Create a scatter plot from this data: x=[1,2,3,4,5], y=[2.1,4.0,5.9,8.1,10.0]"
"Apply publication styling to Fig1 with axis labels Temperature (K) and Absorbance (a.u.)"
"Fit a Gaussian to the data in Book1"
"Export all graphs to C:\Users\me\figures\"
```

File paths can be Windows style (`C:\Users\me\fig.png`) or WSL style (`/mnt/c/Users/me/fig.png`) — the server converts WSL paths automatically, so agents running in WSL can pass their native paths.

### Agent Location vs Server Runtime

The agent does not have to run on Windows. These setups are valid:

- Windows agent -> Windows `origin-pro-mcp` server -> Origin Pro
- WSL agent -> Windows `origin-pro-mcp` server -> Origin Pro

The unsupported setup is WSL/Linux `origin-pro-mcp` server -> Origin Pro, because COM is a Windows API.

## Version Support

This project is currently verified only with **Origin Pro 2020**. The implementation uses Origin's COM Automation Server and LabTalk, which exist across multiple Origin releases, so other versions may work. Treat them as unverified until someone runs the test suite and a real graph/export smoke test on that version.

## Direct LabTalk Safety

The `run_labtalk` tool is available by default for styling, analysis, graph tweaks, and other advanced Origin operations. It blocks common destructive or file-writing LabTalk commands such as project reset, delete, save/open, file dialogs, external script execution, and graph export. Use the typed tools for saving, loading, importing, and exporting.

This is an accident-prevention guard, not a security sandbox for untrusted code.

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
| `list_worksheets` | List open workbooks (with sheets) and graphs |

### Graphing
| Tool | Description |
|------|-------------|
| `create_graph` | Create graph (scatter, line, line+symbol, bar, etc.) |
| `add_plot_to_graph` | Add another dataset to existing graph |
| `set_axis_labels` | Set X/Y axis labels and title |
| `set_axis_range` | Set axis min/max values |
| `export_graph` | Export graph to PNG/JPG/TIF/BMP image |
| `export_all_graphs` | Export every graph in the project |

> Export uses Origin's clipboard copy (the only export route that works
> over COM), so the Windows clipboard contents are replaced during export
> and the image size follows the Origin page setup.

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
| `curve_fit` | Curve fitting: parameters ± std errors, R², SSR, reduced χ²; optional `plot_on_graph` draws the fit curve on a graph |
| `list_fitting_functions` | Show available fit functions |

### Advanced
| Tool | Description |
|------|-------------|
| `run_labtalk` | Execute LabTalk with destructive/file-writing commands blocked |
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

The skill also documents COM quirks observed while testing on **Origin Pro 2020** — what works, what doesn't, and tested workarounds. This is invaluable if you need to customize beyond `apply_publication_style`.

### Customizing the Skill for Your Style

The included skill is a generic starting template for paper-grade Origin figures. You should **copy and customize it** to match your lab's habits, target journals, and visual taste:

- **Font**: Change from Arial to your journal's preferred font (e.g., Helvetica, Times New Roman)
- **Font sizes**: Adjust axis title/tick label/legend sizes to match your journal's figure guidelines
- **Color palette**: Replace the default colorblind-safe palette with your group's standard colors
- **Default export path**: Set to your working directory
- **Figure recipes**: Add templates for your common figure types (XRD patterns, IV curves, etc.)
- **Journal presets**: Add specific formatting rules for your target journals (Nature, ACS, RSC, etc.)

Copy `skills/publication-figure.md` to your project and edit freely — it's meant to be a starting point, not a rigid template.

### Key Origin Pro 2020 COM Quirks (documented in skill)

| Issue | Workaround |
|-------|-----------|
| Bold axis titles (`xb.bold`) doesn't exist | Use `\b(text)` markup in `xb.text$` |
| `legend.text$` doesn't support multiline via COM | Set column Long Names, then `legend -r` |
| `legend.x/y` uses data coordinates, not % | Calculate from `layer.x.from/to` |
| `%C` notation fails via COM | Use actual plot names from `DataPlots` |
| `expGraph` doesn't produce files via COM | Use `CopyPage` + Pillow clipboard |
| `nlr.r2` returns 0 after `nlend` | Read statistics BEFORE `nlend` |
| Plot styling commands can conflict | Add 0.2s delay between `set` commands |
| `[Book]Sheet!col(n).type = ...` silently ignored | Activate the sheet, then use `wks.col(n).type` |
| `set <plot>` fails when the graph isn't active | Run `win -a <graph>` before `set` commands |
| Typed LabTalk locals (`int x = ...`) unreadable later | Use untyped assignment to read values back via COM |

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

`curve_fit` returns the fitted parameter values with standard errors plus
R², SSR, reduced χ², and DoF. Exception: `power` fits and draws the curve,
but Origin 2020 does not expose its parameter values over COM, so only the
statistics are returned. Use `list_fitting_functions` to see the parameter
names for each function.

## Color Palette

`apply_publication_style` uses a muted pastel palette (no pure primaries —
easier on the eyes, survives grayscale printing, colorblind-distinguishable):

| Order | Color | RGB |
|-------|-------|-----|
| 1st | Soft steel blue | (93, 143, 179) |
| 2nd | Muted rose | (204, 102, 119) |
| 3rd | Muted teal | (68, 170, 153) |
| 4th | Soft amber | (221, 170, 102) |
| 5th | Soft purple | (153, 136, 187) |
| 6th | Gray cyan | (119, 170, 187) |

Error bars automatically match their data series color. The fit curve drawn
by `curve_fit(plot_on_graph=...)` uses a muted brick red (170, 68, 80).

> **Tip**: Never use red + green as the only two colors — colorblind users cannot distinguish them.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Could not connect to Origin via COM" | Check that Origin/OriginPro is installed and licensed; if it is, run Origin once as administrator to re-register the Automation Server |
| Tools timeout | Origin may be showing a dialog — check the Origin window |
| Export fails with a clipboard error | The server already polls the clipboard for up to 5 s; check that the Origin window is not minimized and no other app holds the clipboard |
| "Window 'X' not found" errors | The error lists every open workbook/graph — use one of those names (Origin may have renamed the window if the name was taken) |
| Legend missing after styling | Legend uses data coordinates — verify axis range is set before positioning |
| Symbols appear hollow | Do NOT use `set -d` flag (it's for dash patterns, not fill) |

## License

MIT
