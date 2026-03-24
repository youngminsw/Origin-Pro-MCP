from app import mcp
from origin_connection import get_origin, execute_labtalk

@mcp.tool()
def new_project() -> str:
    """Create a new empty Origin project (closes current without saving)."""
    o = get_origin()
    o.NewProject()
    return "New project created"

@mcp.tool()
def save_project(file_path: str = "") -> str:
    """Save the current Origin project.

    Args:
        file_path: Full Windows path to save (e.g., C:\\Users\\data\\experiment.opju).
                   If empty, saves to current location.

    Returns:
        Save confirmation with path
    """
    o = get_origin()
    if file_path:
        o.Save(file_path)
        return f"Project saved to: {file_path}"
    else:
        o.Save("")
        return "Project saved"

@mcp.tool()
def load_project(file_path: str) -> str:
    """Open an Origin project file.

    Args:
        file_path: Full Windows path to .opj or .opju file

    Returns:
        Success message
    """
    o = get_origin()
    o.Load(file_path)
    return f"Loaded project: {file_path}"

@mcp.tool()
def export_all_graphs(
    output_dir: str,
    format: str = "png",
    dpi: int = 300,
    width: int = 800,
    height: int = 600
) -> str:
    """Export all graphs in the current project to image files.

    Args:
        output_dir: Windows directory path for output files
        format: Image format (png, jpg, tif, emf, eps, pdf, svg)
        dpi: Resolution (default 300)
        width: Width in pixels
        height: Height in pixels

    Returns:
        Confirmation message
    """
    escaped_dir = output_dir.replace("\\", "\\\\")

    script = (
        'doc -ef G {'
        '  string __gname$ = page.name$;'
        '  win -a %(__gname$);'
        f'  string __outpath$ = "{escaped_dir}\\\\" + __gname$ + ".{format}";'
        f'  expGraph type:={format} path:=__outpath$ '
        f'  tr1.Width.nVal:={width} tr1.Height.nVal:={height} tr1.Resolution.nVal:={dpi};'
        '};'
    )
    execute_labtalk(script)
    return f"Exported all graphs to: {output_dir}"
