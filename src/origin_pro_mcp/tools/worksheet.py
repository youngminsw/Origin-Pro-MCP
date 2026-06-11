import json
import os

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_origin,
    graph_names,
    require_worksheet,
    workbook_names,
)
from ..labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_path,
    labtalk_string,
    windows_path,
)


@mcp.tool()
def create_worksheet(book_name: str, sheet_name: str = "Sheet1") -> str:
    """Create a new workbook with a worksheet in Origin.

    Args:
        book_name: Name for the new workbook
        sheet_name: Name for the sheet (default: Sheet1)

    Returns:
        Created workbook name (may differ from book_name if it was taken)
    """
    o = get_origin()
    safe_book_name = labtalk_name(book_name, "book_name")
    safe_sheet_name = labtalk_name(sheet_name, "sheet_name")
    name = o.CreatePage(2, safe_book_name, "origin")
    if safe_sheet_name != "Sheet1":
        execute_labtalk(f'page.active$ = "Sheet1"; wks.name$ = {labtalk_string(safe_sheet_name, "sheet_name")};')
    return f"Created workbook: [{name}]{safe_sheet_name}"

@mcp.tool()
def set_worksheet_data(
    book_name: str,
    sheet_name: str,
    columns: str,
    column_names: str = ""
) -> str:
    """Write data to an Origin worksheet.

    Args:
        book_name: Workbook name (e.g., "Book1")
        sheet_name: Sheet name (e.g., "Sheet1")
        columns: JSON array of arrays, each inner array is a column of data.
                 Example: [[1,2,3],[4,5,6]] for 2 columns with 3 rows
        column_names: Optional comma-separated column names (e.g., "X,Y,Error")

    Returns:
        Success message
    """
    o = get_origin()
    safe_book_name = labtalk_name(book_name, "book_name")
    safe_sheet_name = labtalk_name(sheet_name, "sheet_name")
    target = require_worksheet(safe_book_name, safe_sheet_name)

    try:
        cols = json.loads(columns)
    except json.JSONDecodeError as exc:
        msg = (
            "columns must be a JSON array of arrays, e.g. [[1,2,3],[4,5,6]]. "
            f"Parse error: {exc}"
        )
        raise ValueError(msg) from exc
    if isinstance(cols, list) and cols and all(
        isinstance(x, (int, float)) and not isinstance(x, bool) for x in cols
    ):
        # Forgive a flat array: treat [1,2,3] as one column
        cols = [cols]
    if (
        not isinstance(cols, list)
        or not cols
        or not all(isinstance(c, list) and c for c in cols)
    ):
        msg = "columns must be a non-empty JSON array of non-empty arrays, e.g. [[1,2,3],[4,5,6]]."
        raise ValueError(msg)

    for i, col_data in enumerate(cols):
        try:
            float_data = [float(x) for x in col_data]
        except (TypeError, ValueError) as exc:
            msg = f"Column {i + 1} contains non-numeric values; only numbers are supported."
            raise ValueError(msg) from exc
        if not o.PutWorksheet(target, float_data, 0, i):
            msg = f"Origin rejected the data for column {i + 1} of {target}."
            raise ValueError(msg)

    if column_names:
        names = [n.strip() for n in column_names.split(",")]
        activate_window(safe_book_name, "book_name")
        execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet_name, "sheet_name")};')
        for i, name in enumerate(names):
            execute_labtalk(f"wks.col{i+1}.lname$ = {labtalk_string(name, 'column_names')};")

    n_rows = max(len(c) for c in cols)
    return f"Set {len(cols)} columns x {n_rows} rows in {target}"

@mcp.tool()
def get_worksheet_data(book_name: str, sheet_name: str) -> str:
    """Read data from an Origin worksheet.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name

    Returns:
        JSON object with column data
    """
    o = get_origin()
    safe_book_name = labtalk_name(book_name, "book_name")
    safe_sheet_name = labtalk_name(sheet_name, "sheet_name")
    target = f"[{safe_book_name}]{safe_sheet_name}"
    data = o.GetWorksheet(target)
    # On failure GetWorksheet returns None or an HRESULT int, never a sequence
    if not isinstance(data, (list, tuple)):
        books = ", ".join(workbook_names()) or "(none)"
        return json.dumps(
            {"error": f"Worksheet {target} not found. Open workbooks: {books}."}
        )

    if len(data) == 0:
        return json.dumps({"columns": []})

    num_cols = len(data[0])
    columns = []
    for c in range(num_cols):
        columns.append([row[c] for row in data])

    return json.dumps({"columns": columns})

@mcp.tool()
def import_csv_to_worksheet(
    file_path: str,
    book_name: str = "",
    delimiter: str = ","
) -> str:
    """Import a CSV/text file into an Origin worksheet.

    Args:
        file_path: Path to the file (Windows or WSL style, e.g.
                   C:\\Users\\data.csv or /mnt/c/Users/data.csv)
        book_name: Optional workbook name. Auto-generated if empty.
        delimiter: Column delimiter (default: comma)

    Returns:
        Name of the created workbook
    """
    o = get_origin()
    path = windows_path(file_path, "file_path")
    if not os.path.isfile(path):
        msg = f"File not found: {path}"
        raise ValueError(msg)

    if book_name:
        safe_book_name = labtalk_name(book_name, "book_name")
        o.CreatePage(2, safe_book_name, "origin")
        execute_labtalk(f"win -a {safe_book_name};")

    if delimiter == ",":
        ok = execute_labtalk(f"impasc fname:={labtalk_path(path, 'file_path')} options.FileStruct.Delimiter:=1;")
    elif delimiter == "\t":
        ok = execute_labtalk(f"impasc fname:={labtalk_path(path, 'file_path')} options.FileStruct.Delimiter:=0;")
    else:
        safe_delimiter = labtalk_choice(delimiter, {";", "|", " "}, "delimiter")
        ok = execute_labtalk(
            f"impasc fname:={labtalk_path(path, 'file_path')} "
            f"options.FileStruct.CustomDelimiter:={labtalk_string(safe_delimiter, 'delimiter')};"
        )
    if not ok:
        msg = f"Origin could not import {path} — check the file format and delimiter."
        raise ValueError(msg)

    active_book = o.LTStr("page.name$")
    return f"Imported to workbook: {active_book}"

@mcp.tool()
def list_worksheets() -> str:
    """List all open workbooks (with their sheets) and graph windows.

    Returns:
        JSON object: {"workbooks": [{"name", "sheets"}], "graphs": [names]}
    """
    o = get_origin()
    workbooks = []
    pages = o.WorksheetPages
    for i in range(pages.Count):
        page = pages.Item(i)
        layers = page.Layers
        sheets = [layers.Item(j).Name for j in range(layers.Count)]
        workbooks.append({"name": page.Name, "sheets": sheets})

    return json.dumps({"workbooks": workbooks, "graphs": graph_names()})
