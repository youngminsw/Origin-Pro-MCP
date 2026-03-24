# Origin Pro MCP Server

An MCP (Model Context Protocol) server that enables AI assistants to control OriginLab Origin Pro via COM automation.

## Features

- **Worksheet Management** — Create workbooks, read/write data, import CSV files
- **Graph Creation** — Scatter, line, line+symbol, bar, histogram, box, contour, pie, bubble plots
- **Plot Styling** — Colors, symbols, line width, publication-ready formatting in one call
- **Curve Fitting** — Linear, polynomial, exponential, Gaussian, Lorentz, Voigt, and more
- **Project Management** — New/save/load projects, export all graphs
- **LabTalk Scripting** — Direct LabTalk execution for advanced operations

## Requirements

- Windows with Origin Pro 2020+ installed
- Python 3.10+ (Windows, not WSL)
- Origin must be running before starting the server

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Claude Code (recommended)

Add to your Claude Code MCP settings:

```json
{
  "mcpServers": {
    "origin-pro": {
      "command": "python",
      "args": ["C:\\Users\\<you>\\origin-mcp-server\\server.py"],
      "cwd": "C:\\Users\\<you>\\origin-mcp-server"
    }
  }
}
```

### Manual

```bash
python server.py
```

The server communicates via stdio using the MCP protocol.

## Tools

| Tool | Description |
|------|-------------|
| `new_project` | Create new Origin project |
| `save_project` | Save project to file |
| `load_project` | Open .opj/.opju file |
| `create_worksheet` | Create new workbook |
| `set_worksheet_data` | Write data to worksheet |
| `get_worksheet_data` | Read data from worksheet |
| `import_csv_to_worksheet` | Import CSV/text file |
| `list_worksheets` | List all open workbooks |
| `create_graph` | Create graph from data |
| `add_plot_to_graph` | Add data series to graph |
| `set_axis_labels` | Set axis labels and title |
| `set_axis_range` | Set axis min/max |
| `export_graph` | Export graph to image |
| `export_all_graphs` | Export all graphs |
| `set_plot_style` | Set line/symbol style per plot |
| `apply_publication_style` | One-call publication formatting |
| `set_graph_font` | Set font for graph elements |
| `set_legend` | Configure legend text/position |
| `set_tick_style` | Set tick direction/length |
| `curve_fit` | Curve fitting with statistics |
| `list_fitting_functions` | List available fit functions |
| `run_labtalk` | Execute raw LabTalk script |
| `get_labtalk_variable` | Read LabTalk variable value |

## Architecture

```
Claude Code (WSL/Windows)
    |  stdio (MCP protocol)
    v
MCP Server (Windows Python)
    |  COM automation (win32com)
    v
Origin Pro (GUI visible)
```

The server uses `Origin.ApplicationSI` to attach to the running Origin instance. All operations are reflected in the Origin GUI in real-time.

## Notes

- Origin 2020 COM has specific quirks handled by this server:
  - Legend text uses column Long Names + `legend -r` (direct `legend.text$` doesn't support multiline)
  - Legend positioning uses data coordinates via `execute_labtalk` (not `gl.Execute`)
  - Plot styling commands need small delays between calls to avoid rendering bugs
  - Graph export uses clipboard-based approach (CopyPage + Pillow) for reliability

## License

MIT
