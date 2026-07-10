# Origin Pro MCP Server

An MCP (Model Context Protocol) server that enables AI assistants like Claude to control **OriginLab Origin Pro** via COM automation. All operations are reflected in Origin's GUI in real-time — you watch as the AI creates worksheets, plots graphs, and styles figures.

## What Can It Do?

- **Worksheet Management** — Create workbooks, read/write data, import CSV/Excel, export CSV, column formulas, sort, transpose
- **Matrices & 3D** — Matrix books, XYZ gridding, 3D surface/scatter, contour, heatmap, image plots
- **Graph Creation** — Scatter, line, line+symbol, column, bar, area, pie, histogram, contour plots
- **Graph Layers & Axes** — Log scales, dual Y axis, panels, reference lines, text annotations
- **Plot Styling** — Colors, symbols, line width, publication-ready formatting in one call
- **Analysis** — Curve fitting, FFT, smoothing, integration, differentiation, interpolation, peak finding
- **Statistics** — Descriptive stats, two-sample t-test, frequency counts (via the `stats` and `transform` tools)
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

The background daemon runs one isolated Origin instance per session. These
environment variables harden it against a wedged Origin COM call (a synchronous
operation that never returns) and against destructive mistakes.
Some default **off**; the wedge/data-safety ones now default **on** because they
are safe (see each row). Set any to `off` to opt out.

| Variable | Default | Effect |
| :--- | :--- | :--- |
| `ORIGIN_PRO_MCP_DISPATCH_TIMEOUT` | `90` | Soft budget (seconds) for each tool dispatch. If Origin doesn't respond within it, a persistent per-session watchdog is polled for the modal dialog that is blocking it: when one is found, the error **names its exact title** and says whether it was already auto-dismissed (retry the call) or is waiting for you to close it by hand; if none is found it falls back to the generic "most likely a modal dialog" message. NOTHING is killed at this stage. Set `off`/`0` to disable. |
| `ORIGIN_PRO_MCP_DISPATCH_KILL_GRACE` | `90` | Grace (seconds) AFTER the soft warning before Origin is **force-killed** as a last resort (so a wedged session can never permanently hold a pool slot). Total time to a hard reset = timeout + grace (default 180s). Set `off`/`0` for no warning phase (legacy: force-kill straight at the soft budget). |
| `ORIGIN_PRO_MCP_DIALOG_AUTODISMISS` | `on` | Every session runs a persistent watchdog that polls (~2s) for modal dialogs owned by its Origin process and records their titles. By **default** it also auto-dismisses (closes) each one it finds, so a startup or mid-session dialog no longer freezes the session. Set to `0`/`off`/`false`/`no` to keep detection and reporting (dispatch-timeout errors still name the dialog) without the daemon closing it — you then close it by hand in the Origin window. |
| `ORIGIN_PRO_MCP_AUTOSAVE` | `on` | Save the project **in place** (its own file, same name — like pressing Save) **before** a destructive op (delete graph/plot, column deletion, project load/new, overwriting a populated sheet, or a `confirm`ed destructive `run_labtalk`), so a bad edit is recoverable by reloading. It never writes a differently-named copy, and never overwrites a real file with an empty/blanked project (N5-safe). Only saves a project that already has a file on disk. Set `off` to disable autosave entirely. |
| `ORIGIN_PRO_MCP_AUTOSAVE_INTERVAL` | `300` | Also save the project **in place** every N seconds (proactive autosave), not just before destructive ops. Applies to agent-isolated sessions with a saved project; the Origin you `ATTACH` to is left to you. `off`/`0` disables periodic autosave (preflight still runs). |
| `ORIGIN_PRO_MCP_AUTOSAVE_REQUIRED` | `1` | When autosave is on and a required *preflight* in-place save fails, the destructive op is **not** run and an error is returned. Set `0` to proceed without saving. |
| `ORIGIN_PRO_MCP_REAP_CLOSE` | `off` | Session lifecycle: by **default** a session ending gracefully (idle / client disconnect) is **detached** — the session's worker thread stops but *your Origin window is left exactly as it was (original save path and unsaved edits intact) so you keep the project*. Set `1` to restore the old save-a-recovery-copy-and-close behavior. (A *wedged* session's Origin is still force-killed — the only way to free a worker stuck in a synchronous COM call.) |
| `ORIGIN_PRO_MCP_SWEEP_ORPHANS` | `off` | By **default** a restarting daemon does **not** kill leftover Origin windows (so a restart never destroys a project you kept open). Set `1` to have startup reclaim leftover Origins (orphan cleanup, at the cost of closing kept windows). |
| `ORIGIN_PRO_MCP_NO_SPAWN` | `off` | Set `1` to stop the shim from auto-respawning the daemon. Use it to shut the daemon down from the process manager and keep it stopped — tool calls then return a clear "daemon not running" error instead of relaunching it. |
| `ORIGIN_PRO_MCP_ATTACH` | `off` | Set `1` so this session **attaches to the Origin you already have open** (the shared `ApplicationSI` instance) instead of spawning a fresh isolated one — the agent then works on your currently-open project. Only **one** session can attach (a second falls back to an isolated instance); other agents keep their own isolated Origins. The attached instance is never force-killed by the daemon (it's yours). |

Per-call override: `run_labtalk(script, confirm=True, timeout=120)` bounds that
one call even when a longer/shorter budget than `ORIGIN_PRO_MCP_DISPATCH_TIMEOUT`
is needed (and works even when the timeout is `off`).

**Rollback:** unset any of these (or set the timeout to `off`) to return to the
prior behavior — no code change or redeploy required.

#### Session lifecycle & restarts

Each MCP client process gets its own daemon session (its own isolated Origin).
Because sessions and the daemon can restart independently, the daemon keeps a
small ledger sidecar (`sessions.json`, next to the private lockfile) recording
each session's last Origin PID and project path. When a **new** session starts,
the daemon reads that ledger once and, on the first successful tool response,
piggybacks a short one-time **`[origin-mcp]` notice** telling the agent what
happened — so it continues the work instead of silently rebuilding into an empty
window. What you may see:

- **Your MCP client restarted (new session).** Your previous Origin window was
  **detached, not closed** (see `ORIGIN_PRO_MCP_REAP_CLOSE`): the notice says it
  is still open with your project and to save/close it in the GUI before
  reloading, or just work in the fresh instance.
- **The daemon restarted and your old Origin is gone.** The notice says your
  project is not loaded and to reopen it with `load_project`.
- **Ghost windows.** Leftover Origins from earlier sessions are preserved by
  default, so they can accumulate. The notice summarizes how many are still open;
  close them in the GUI once saved, or set `ORIGIN_PRO_MCP_SWEEP_ORPHANS=1` so a
  restarting daemon reclaims them.
- **Attach (`ORIGIN_PRO_MCP_ATTACH=1`).** If you got the user's open Origin, the
  notice reminds you that autosave and force-recovery are **disabled** there —
  save explicitly and avoid destructive ops. If another session already holds the
  single attach slot, the notice says you got an isolated Origin instead.

Separately, `load_project` appends a one-line **collision warning** to its result
when the ledger shows another live Origin still holding the same project file
(saving from both would clobber it — close the other first). The notice and the
warning are advisory strings only; they never block a call, and a
missing/corrupt ledger is treated as empty.

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

### LabTalk Gotchas (Origin 2020, styling-report fixes)

- **One flag per `set` call.** `set <ds> -c color(255,0,0) -cf color(255,0,0);`
  (combining flags in ONE command) silently wipes the plot to black; the same
  applies to `-k`/`-kf`/`-z` combined (can blank the symbol). Send each flag
  as its own `set <ds> -flag val;` call.
- **Never write `layer.x2.majorTicks` / `layer.y2.majorTicks`.** Setting it
  to 0 wipes the number labels on ALL FOUR axes, not just the opposite side.
  Use `layer.<ax>.ticks = 0` (or `axis(op="tick", axis="top"/"right",
  tick_direction="none")`) to remove tick marks instead.
- **Units differ between line width and error-bar width.** `set -w` is
  ~200 units per point (500 = 2.5pt); error bars use `-erw <points>` /
  `-erwc <cap width>` directly in points — do not reuse the `-w` scale for
  error bars, and never style them via a bare `set -w`/`-ew`.
  `set_plot_style(error_bar_width=, error_cap_width=)` handles the units.
- **The active window matters.** `layer.*`, `col()`, and `%C` all target
  whatever window is currently active — pass `window=<name>` to `run_labtalk`
  to activate it first, or prefer the typed tools (`graph_name`/`book_name`
  params never depend on activation state).
- **A freshly created page needs a moment before its FIRST styling/read/
  export command** — `create_graph`/`add_plot_to_graph`/`ungroup_plots`
  handle this internally now; a raw `run_labtalk` sequence right after
  `CreatePage`/`plotxy` may still need its own settle.
- **Symbol shape `-k` codes** (Origin 2020, re-verified live): 1=square,
  2=circle, 3=triangle-up, 4=triangle-down, 5=diamond, 6=plus, 7=x/cross,
  8=asterisk. Codes 9-12 render as a dash/vertical-bar/literal glyph, not
  useful marker shapes.

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

## Available Tools (45 total)

Several tools are **dispatchers**: one tool name with an `op`/`kind`/`method`
argument that selects the specific action, so a handful of tools cover what
used to be many single-purpose ones.

### Project Management
| Tool | Description |
|------|-------------|
| `new_project` | Create new empty Origin project |
| `save_project` | Save project to .opju file |
| `load_project` | Open existing .opj/.opju file |
| `export_all_graphs` | Export every graph in the project to image files |
| `save_graph_template` | Save a graph as a reusable .otpu/.otp template |

> **Note**: `create_worksheet`, `create_matrix`, `create_graph`,
> `create_matrix_plot`, `import_data`, and `worksheet_to_matrix` return a
> JSON string (not a sentence) with the actual assigned name — Origin may
> rename on collision, so read `"name"` from the result rather than assuming
> the requested name was used.

### Worksheet Data
| Tool | Description |
|------|-------------|
| `create_worksheet` | Create new workbook |
| `set_worksheet_data` | Write column data (JSON arrays) |
| `get_worksheet_data` | Read worksheet data as JSON (empty cells → null) |
| `import_data` | Import a CSV/text or Excel file (`format="auto"/"csv"/"excel"`); CSV/text import suppresses Origin's auto-generated sparkline mini-graph windows by default (`sparklines=False`) |
| `export_worksheet` | Export a worksheet to CSV/text |
| `list_worksheets` | List open workbooks, graphs, and matrices |
| `manage_columns` | Add, delete, or edit columns — `op="add"/"delete"/"properties"/"formula"` |
| `sort_worksheet` | Sort rows by a column (asc/desc) |
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
| `remove_plot` | Remove one data plot from a graph (uses `layer -e` + `layer -ie`, actually deletes; `layer -d` would delete the whole layer) |
| `set_error_bars` | Attach Y/X error bars to an existing plot from an error column (no duplicate) |
| `set_layer_geometry` | Set a layer's panel position/size (left/top/width/height) |
| `add_second_y_axis` | Add a right-Y layer and plot on it |
| `add_layer` | Add a panel/axis layer (right-y, top-x, inset, independent) |
| `axis` | Configure axes — `op="labels"/"range"/"scale"/"tick"/"frame"` |
| `annotate` | Add an annotation — `kind="reference_line"/"text"/"line"/"arrow"` |
| `colormap` | Apply a palette and/or set the Z color-scale range on a colormapped graph |
| `export_graph` | Export to an image file; `sized=True` for an exact pixel width/height (default ~1200px wide) |
| `ungroup_plots` | Break a plot group so each curve can be colored independently |

### Styling
| Tool | Description |
|------|-------------|
| `apply_publication_style` | **One-call publication formatting** (recommended) |
| `set_plot_style` | Set color, line width, symbol shape/size, and open/solid marker |
| `set_graph_font` | Set font family, size, and optional bold |
| `set_legend` | Configure legend text and position |
| `set_tick_labels` | Tick-label numeric format (decimal/scientific/engineering), bold, decimal places |

### Analysis
| Tool | Description |
|------|-------------|
| `curve_fit` | Curve fitting: parameters ± std errors, R², SSR, reduced χ²; optional `plot_on_graph` |
| `list_fitting_functions` | Show available fit functions |
| `transform` | Numerical transform on an XY curve — `method="integrate"/"differentiate"/"smooth"/"interpolate"/"fft"/"find_peaks"` |
| `stats` | Statistics on worksheet columns — `op="column"/"compare_means"/"frequency"` |

### Advanced
| Tool | Description |
|------|-------------|
| `run_labtalk` | Execute LabTalk with destructive/file-writing commands blocked; optional `capture` reads variables back. If a 2+ statement script fails as a whole, it is automatically retried statement-by-statement and each statement's OK/FAILED status is reported (partial application is possible — this is intentional) |
| `get_labtalk_variable` | Read a LabTalk variable value |

### Skills
| Tool | Description |
|------|-------------|
| `list_skills` | List bundled skills (name, title, when to use) |
| `get_skill` | Load a skill's full markdown instructions by name |

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
   - Use colorblind-safe color palette (steel blue → rose → teal → amber → purple → gray cyan; see the Color Palette table below)
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
| `expGraph` needs `path:=`/`filename:=`/`overwrite:=replace` (a full path or missing args opens a dialog) | `export_graph` writes the file directly via `expGraph`, no clipboard involved |
| `nlr.r2` returns 0 after `nlend` | Read statistics BEFORE `nlend` |
| Combining multiple `-flag`s in ONE `set` command corrupts the plot (color reset to black, or symbol blanked) | Send one `set <ds> -flag val;` call per flag, not a delay — see "LabTalk Gotchas" above |
| `layer.x2.majorTicks`/`layer.y2.majorTicks` wipes ALL axes' number labels | Use `layer.<ax>.ticks = 0` instead |
| `[Book]Sheet!col(n).type = ...` silently ignored | Activate the sheet, then use `wks.col(n).type` |
| `set <plot>` fails when the graph isn't active | Run `win -a <graph>` before `set` commands |
| Typed LabTalk locals (`int x = ...`) unreadable later | Use untyped assignment to read values back via COM |
| A graph loaded from a `.opju` can report zero data plots over COM (per-curve styling/ungrouping silently no-ops) | The core per-curve/axis/frame tools (`set_plot_style`, `ungroup_plots`, `remove_plot`, `axis` range/scale/tick) now activate the page and re-acquire a fresh layer handle before each call; if the layer is still empty, the tool raises an actionable error instead of returning fake success — recreate the graph in-session if you hit it. Text/font/legend tools (`set_graph_font`, `set_legend`) still go through plain LabTalk and can silently no-op on a loaded graph — verify those visually |

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
R², SSR, reduced χ², and DoF. Exceptions: `power` fits and draws the curve,
but Origin 2020 does not expose its parameter values over COM, so only the
statistics are returned; and a `line` fit *without* `plot_on_graph` uses the
fast `fitlr` path, which returns intercept/slope values and R² only (no
standard errors, SSR, χ², or DoF — pass `plot_on_graph` to get them). Use
`list_fitting_functions` to see the parameter names for each function.

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
| Tools timeout / "Origin has not responded" | A persistent watchdog polls for modal dialogs on your session's Origin process; by default it auto-dismisses them, and a timeout error names the dialog's exact title and whether it was closed for you. If auto-dismiss is off (`ORIGIN_PRO_MCP_DIALOG_AUTODISMISS=0`), switch to the Origin window and close/confirm the named dialog yourself — the operation then finishes on its own |
| "No editable data plots found" on `set_plot_style`/`ungroup_plots`/axis ops, or "did not take effect" on an axis-range call | The graph was loaded from a `.opju` project file in a state Origin freezes over COM (its layer reports zero data plots even after activating the window). This used to silently no-op; it is now a raised error instead of a fake success. Recreate the graph in-session (`create_graph`/`plotxy`) or reopen the project fresh — no manual export-and-compare-pixels needed to detect it anymore |
| "Export failed: ... was not created" | `export_graph` writes the file directly via `expGraph` (no clipboard); this means Origin (Windows) could not write to that path — use a Windows path (`C:\...`) or `/mnt/<drive>/...` instead of a WSL/Linux path |
| "... is a Linux/WSL path that Origin ... cannot access" | Any file-path argument (import/export/save/load) given as a bare POSIX path (e.g. `/tmp/x.csv`) is now rejected upfront with this error rather than silently resolving against the wrong drive — use `/mnt/<drive>/...` (auto-translated) or a native Windows path |
| "Window 'X' not found" errors | The error lists every open workbook/graph — use one of those names (Origin may have renamed the window if the name was taken) |
| Legend missing after styling | Legend uses data coordinates — verify axis range is set before positioning |
| Symbols appear hollow | Do NOT use `set -d` flag (it's for dash patterns, not fill) |

## License

MIT
