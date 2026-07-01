# Origin Pro MCP Server

An MCP (Model Context Protocol) server that enables AI assistants like Claude to control **OriginLab Origin Pro** via COM automation. All operations are reflected in Origin's GUI in real-time — you watch as the AI creates worksheets, plots graphs, and styles figures.

## What Can It Do?

- **Worksheet Management** — Create workbooks, read/write data, import CSV/Excel, export CSV, column formulas, sort, transpose
- **Matrices & 3D** — Matrix books, XYZ gridding, 3D surface/scatter, contour, heatmap, image plots
- **Graph Creation** — Scatter, line, line+symbol, column, bar, area, pie, histogram, contour plots
- **Graph Layers & Axes** — Log scales, dual Y axis, panels, reference lines, text annotations
- **Plot Styling** — Colors, symbols, line width, publication-ready formatting in one call
- **Analysis** — Curve fitting, FFT, smoothing, integration, differentiation, interpolation, peak finding
- **Statistics** — Descriptive stats, two-sample t-test, frequency counts
- **Project Management** — New/save/load projects, export all graphs (with pixel-size control)
- **LabTalk Scripting** — Direct LabTalk execution with destructive/file-overwrite commands blocked

This MCP server is intentionally **Windows-runtime-only**. The AI agent or MCP client can run from Windows or WSL, but the MCP server process that talks to Origin must be launched with Windows Python and `pywin32`. WSL/Linux can edit the project and run non-COM unit tests, but cannot directly control Origin COM.

## Quick Start

### Agentic Setup (read this if you are an AI agent configuring this server)

Copy-paste rules to self-configure the `origin-pro` MCP with zero human steps:

1. **Environment requirement.** This server only works on **Windows** with a licensed **Origin/OriginPro** install and **Windows Python with `pywin32`**. COM automation cannot run from WSL or Linux Python. If the agent runs in WSL, it must launch the **Windows** executable (`uvx.exe` / Windows `origin-pro-mcp`), never WSL Python.
2. **Preferred install path.** Use `uvx` (zero setup) or `pip install origin-pro-mcp`. Both install to a Windows-local location and "just work":
   ```json
   { "mcpServers": { "origin-pro": { "command": "uvx", "args": ["--quiet", "origin-pro-mcp"] } } }
   ```
   From WSL, call `uvx.exe` instead of `uvx` (see Option A below).
3. **Critical UNC rule.** If you point at a **local checkout** from WSL, the repo MUST live under a **Windows drive** (e.g. `C:\Users\YOU\Origin-Pro-MCP`). **Never** reference it through a `\\wsl.localhost\...` UNC path: the background daemon (`python -m origin_pro_mcp.daemon`) cannot resolve a `-m` module or a working directory over UNC, so it fails silently (Origin never opens). `uvx`/`pip install` avoid this entirely.
4. **Discover skills first.** After connecting, call the **`list_skills`** tool to see bundled skills (e.g. `publication-figure`) and when to use each, then **`get_skill(name)`** to load the full step-by-step instructions before producing figures/plots/analysis. No file copying is needed — skills ship inside the package and are exposed by the server.

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

> **WSL users — the package must live on a Windows-local path, not a `\\wsl.localhost\...` (UNC) path.** The server runs a background daemon (`python -m origin_pro_mcp.daemon`); Windows cannot resolve a `-m` module or use a working directory over a UNC path, so launching it from a WSL-filesystem checkout fails silently (Origin never opens and a console window flashes). `uvx`/`pip install` already install to a Windows-local location, so they are unaffected. If you point at a local checkout, **clone it under a Windows drive** (e.g. `C:\Users\YOU\Origin-Pro-MCP`) and reference that path — do not point the MCP config at a repo inside the WSL filesystem.

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

You do not need to start Python or Origin manually. The MCP client starts `origin-pro-mcp`, which launches an isolated Origin instance for the session. By default the Origin window is **visible**, so you watch the agent create worksheets and plot graphs in real time.

**Display mode** — set the `ORIGIN_PRO_MCP_VISIBLE` environment variable in your MCP config:

| Value | Mode | Use for |
| :--- | :--- | :--- |
| `1` (default) | **Visible** — Origin window shown | Watching the agent work interactively |
| `0` | **Invisible** — Origin runs hidden | Headless/batch runs, many concurrent agents, no window pop-ups |

```json
{ "mcpServers": { "origin-pro": {
  "command": "origin-pro-mcp",
  "env": { "ORIGIN_PRO_MCP_VISIBLE": "0" }
} } }
```

#### Reliability & recovery (advanced env vars)

The background daemon runs one isolated Origin instance per session. A few
opt-in environment variables harden it against a wedged Origin COM call (a
synchronous operation that never returns) and against destructive mistakes.
All default **off** — behavior is unchanged unless you set them.

| Variable | Default | Effect |
| :--- | :--- | :--- |
| `ORIGIN_PRO_MCP_DISPATCH_TIMEOUT` | `off` | Seconds to bound each tool dispatch. If an Origin operation wedges past this budget, the daemon force-terminates *that session's* Origin process, frees the pool slot, and returns an actionable error — the daemon itself keeps serving other sessions. Set e.g. `120`. `off`/`0` disables. |
| `ORIGIN_PRO_MCP_AUTOSAVE` | `off` | Set `on` to snapshot a recoverable project copy **before** a destructive op (delete graph/plot, column deletion, project load/new, overwriting a populated sheet, or a `confirm`ed destructive `run_labtalk`). Origin's `Save(path)` rebinds the project identity, so autosave writes a timestamped backup and then re-saves your project to restore its original path — meaning autosave also re-persists your open project. Opt-in for that reason. |
| `ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED` | `1` | When autosave is on and a required snapshot fails, the destructive op is **not** run and an error is returned. Set `0` to proceed without a backup. |
| `ORIGIN_PRO_MCP_AUTOSAVE_RETENTION` | `3` | How many autosave copies to keep per project (oldest pruned). |
| `ORIGIN_PRO_MCP_AUTOSAVE_DIR` | project dir | Directory for autosave copies (defaults alongside the saved project, or the daemon's working dir for an unsaved project). Files are named `<project>.autosave-<timestamp>.opju`. |

Per-call override: `run_labtalk(script, confirm=True, timeout=120)` bounds that
one call even when `ORIGIN_PRO_MCP_DISPATCH_TIMEOUT` is off.

**Rollback:** unset any of these (or set the timeout to `off`) to return to the
prior behavior — no code change or redeploy required.

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

## Direct CLI (no MCP client)

Every MCP tool is also a plain command, so the repo alone can drive
Origin — no MCP client, no per-task scripts. Same Windows-runtime rule
applies (Windows Python + pywin32 + Origin).

```bash
# List all tools with their arguments
python -m origin_pro_mcp.cli list

# Call any tool: simple flags...
python -m origin_pro_mcp.cli list_worksheets
python -m origin_pro_mcp.cli apply_publication_style --graph_name Fig1

# ...or --json for values with spaces/paths (recommended for scripting)
python -m origin_pro_mcp.cli apply_publication_style --json '{"graph_name": "Fig1", "x_label": "Temperature (K)"}'
python -m origin_pro_mcp.cli export_graph --json '{"graph_name": "Fig1", "file_path": "/mnt/c/Users/me/fig1.png"}'
```

After `pip install` (or via `uvx`) the same is available as the
`origin-pro-cli` command, e.g. `origin-pro-cli list_worksheets`.

Running a WSL agent? Invoke Windows Python with the package on the path,
for example from the repo's `src/` directory:

```bash
cd src && /mnt/c/.../python.exe -m origin_pro_mcp.cli list_worksheets
```

The CLI reflects over the same tool registry as the MCP server, so it
always exposes exactly the tools listed below.

## Available Tools (56 total)

### Project Management
| Tool | Description |
|------|-------------|
| `new_project` | Create new empty Origin project |
| `save_graph_template` | Save a graph as a reusable .otpu/.otp template |
| `save_project` | Save project to .opju file |
| `load_project` | Open existing .opj/.opju file |

### Worksheet Data
| Tool | Description |
|------|-------------|
| `create_worksheet` | Create new workbook |
| `set_worksheet_data` | Write column data (JSON arrays) |
| `get_worksheet_data` | Read worksheet data as JSON |
| `import_csv_to_worksheet` | Import CSV/text file |
| `import_excel` | Import an .xls/.xlsx file |
| `export_worksheet` | Export a worksheet to CSV/text |
| `list_worksheets` | List open workbooks, graphs, and matrices |
| `set_column_formula` | Fill a column from a formula of other columns |
| `set_column_properties` | Set long name, units, comment, designation |
| `set_column_designation` | Set a column's plot role (x/y/z/yerr/xerr/label) by name |
| `sort_worksheet` | Sort rows by a column (asc/desc) |
| `add_columns` / `delete_columns` | Add or remove columns |
| `transpose_worksheet` | Transpose rows and columns |

### Matrix
| Tool | Description |
|------|-------------|
| `create_matrix` | Create a matrix book |
| `set_matrix_data` / `get_matrix_data` | Write / read a 2D grid |
| `worksheet_to_matrix` | Grid scattered XYZ into a matrix (xyz2mat) |
| `create_matrix_plot` | Surface (3D), contour, heatmap, or image from a matrix (with a data-linked color scale) |

### Graphing
| Tool | Description |
|------|-------------|
| `create_graph` | Create graph (scatter, line, line+symbol, column, bar, area, pie, histogram, box, contour, 3d_scatter) |
| `add_plot_to_graph` | Add another dataset to an existing graph |
| `delete_graph` | Delete a graph window |
| `remove_plot` | Remove one data plot from a graph (uses `layer -d`, actually deletes) |
| `set_error_bars` | Attach Y/X error bars to an existing plot from an error column (no duplicate) |
| `set_layer_geometry` | Set a layer's panel position/size (left/top/width/height) |
| `add_second_y_axis` | Add a right-Y layer and plot on it |
| `add_layer` | Add a panel/axis layer (right-y, top-x, inset) |
| `set_axis_labels` | Set X/Y axis labels and title |
| `set_axis_range` | Set axis min/max values |
| `set_axis_scale` | Linear / log10 / ln / log2 scale (auto-rescales range to data) |
| `add_reference_line` | Horizontal/vertical line at a value |
| `add_line` | Straight line between two data points |
| `add_arrow` | Single/double-headed arrow between two points |
| `apply_color_map` | Apply a colormap; bundles viridis/cividis/plasma/inferno/magma (colorblind-safe) plus Origin built-ins |
| `set_colormap_levels` | Set the Z color-scale range |
| `add_text_annotation` | Place a text label at data coordinates |
| `export_graph` | Export via clipboard (page size) |
| `export_graph_sized` | Export at a chosen pixel width/height |
| `export_all_graphs` | Export every graph in the project |

> `export_graph` uses Origin's clipboard copy (size follows the page
> setup). `export_graph_sized` uses `expGraph` for direct pixel control.

### Styling
| Tool | Description |
|------|-------------|
| `apply_publication_style` | **One-call publication formatting** (recommended) |
| `set_plot_style` | Set color, line width, symbol shape/size, and open/solid marker |
| `set_graph_font` | Set font family, size, and optional bold |
| `set_legend` | Configure legend text and position |
| `set_tick_style` | Set tick direction and length |
| `set_tick_labels` | Tick-label numeric format (decimal/scientific/engineering), bold, decimal places |

### Analysis
| Tool | Description |
|------|-------------|
| `curve_fit` | Curve fitting: parameters ± std errors, R², SSR, reduced χ²; optional `plot_on_graph` |
| `list_fitting_functions` | Show available fit functions |
| `integrate` | Area under the curve |
| `differentiate` | Derivative dY/dX into a new column |
| `smooth` | Savitzky-Golay / adjacent / binomial smoothing |
| `interpolate` | Resample onto evenly spaced X (linear/spline/bspline/akima) |
| `fft` | Forward FFT + dominant frequency |
| `find_peaks` | Peak positions and heights |

### Statistics
| Tool | Description |
|------|-------------|
| `column_statistics` | mean, sd, se, variance, median, min, max, sum, n |
| `compare_means` | Two-sample t-test (t, df, p, means) |
| `frequency_count` | Binned histogram counts |

### Advanced
| Tool | Description |
|------|-------------|
| `run_labtalk` | Execute LabTalk with destructive/file-writing commands blocked; optional `capture` reads variables back |
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

This server ships a **bundled skill** (`src/origin_pro_mcp/skills/publication-figure.md`) that teaches Claude how to create journal-quality figures step by step. **No manual copying is needed** — the skill is packaged inside the wheel and exposed over MCP, so any connecting agent discovers it automatically:

1. Call the **`list_skills`** tool — it returns each skill's name, title, and when to use it (e.g. `publication-figure`).
2. Call **`get_skill("publication-figure")`** to load the full markdown instructions.
3. Each skill is also available as an MCP **resource** at `skill://<name>` (e.g. `skill://publication-figure`) for clients that browse resources.

When you ask Claude to "make a publication figure", it can autonomously pull this skill and follow its workflow:
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

Pull the skill with `get_skill("publication-figure")` (or copy `src/origin_pro_mcp/skills/publication-figure.md`) into your project and edit freely — it's meant to be a starting point, not a rigid template.

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

| Type | Description | Data |
|------|-------------|------|
| `scatter` | Scatter plot (symbols only) | X, Y |
| `line` | Line plot (no symbols) | X, Y |
| `line+symbol` | Line with symbols (recommended for publications) | X, Y |
| `column` | Vertical bar chart | X, Y |
| `bar` | Horizontal bar chart | X, Y |
| `area` | Area plot | X, Y |
| `pie` | Pie chart | X, Y |
| `histogram` | Histogram | single Y column |
| `box` | Box chart | single Y column |
| `contour` | Filled contour map | X, Y, Z (pass `z_col`) |
| `3d_scatter` | 3D scatter (OpenGL) | X, Y, Z (pass `z_col`) |

> Matrix plots (surface, contour, heatmap, image) come from a matrix via
> `create_matrix_plot`.

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
