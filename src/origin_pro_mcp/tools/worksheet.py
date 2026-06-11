import csv
import json
import os

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_origin,
    get_lt_str,
    graph_names,
    matrix_names,
    require_worksheet,
    workbook_names,
)
from ..labtalk_safe import (
    labtalk_choice,
    labtalk_name,
    labtalk_formula,
    positive_column,
    positive_int,
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
    """List all open workbooks (with their sheets), graphs, and matrices.

    Returns:
        JSON object: {"workbooks": [{"name", "sheets"}], "graphs": [names],
        "matrices": [names]}
    """
    o = get_origin()
    workbooks = []
    pages = o.WorksheetPages
    for i in range(pages.Count):
        page = pages.Item(i)
        layers = page.Layers
        sheets = [layers.Item(j).Name for j in range(layers.Count)]
        workbooks.append({"name": page.Name, "sheets": sheets})

    return json.dumps(
        {"workbooks": workbooks, "graphs": graph_names(), "matrices": matrix_names()}
    )


# LabTalk column designation codes for set_column_properties.
_DESIGNATIONS = {"none": 0, "x": 4, "y": 1, "z": 6, "yerr": 3, "xerr": 5, "label": 5}


@mcp.tool()
def set_column_formula(book_name: str, sheet_name: str, col: int, formula: str) -> str:
    """Fill a worksheet column from a formula of other columns.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name
        col: Target column (1-based); created if it does not exist
        formula: LabTalk expression, e.g. "col(1)^2", "col(2)*100/col(3)"

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    safe_col = positive_column(col, "col")
    safe_formula = labtalk_formula(formula, "formula")
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_book, "book_name")
    execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet, "sheet_name")};')
    # Grow the sheet if the target column does not exist yet.
    while get_origin().LTVar("wks.ncols") < safe_col:
        execute_labtalk("wks.addCol();")
    if not execute_labtalk(f"col({safe_col}) = {safe_formula};"):
        msg = f"Origin could not evaluate formula '{safe_formula}' into column {safe_col}."
        raise ValueError(msg)
    return f"Set column {safe_col} of [{safe_book}]{safe_sheet} = {safe_formula}"


@mcp.tool()
def sort_worksheet(book_name: str, sheet_name: str, col: int, descending: bool = False) -> str:
    """Sort all rows of a worksheet by one column.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name
        col: Column to sort by (1-based)
        descending: Sort high-to-low when True

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    safe_col = positive_column(col, "col")
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_book, "book_name")
    execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet, "sheet_name")};')
    order = 1 if descending else 0
    if not execute_labtalk(f"wsort bycol:={safe_col} descending:={order};"):
        msg = f"Origin could not sort [{safe_book}]{safe_sheet} by column {safe_col}."
        raise ValueError(msg)
    direction = "descending" if descending else "ascending"
    return f"Sorted [{safe_book}]{safe_sheet} by column {safe_col} ({direction})"


@mcp.tool()
def add_columns(book_name: str, sheet_name: str, count: int = 1) -> str:
    """Append empty columns to a worksheet.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name
        count: Number of columns to add (default 1)

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    safe_count = positive_int(count, "count")
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_book, "book_name")
    execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet, "sheet_name")};')
    for _ in range(safe_count):
        execute_labtalk("wks.addCol();")
    return f"Added {safe_count} column(s) to [{safe_book}]{safe_sheet}"


@mcp.tool()
def delete_columns(book_name: str, sheet_name: str, col: int, count: int = 1) -> str:
    """Delete one or more columns starting at a position.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name
        col: First column to delete (1-based)
        count: How many consecutive columns to delete (default 1)

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    safe_col = positive_column(col, "col")
    safe_count = positive_int(count, "count")
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_book, "book_name")
    execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet, "sheet_name")};')
    o = get_origin()
    deleted = 0
    for _ in range(safe_count):
        if o.LTVar("wks.ncols") < safe_col:
            break
        # After deleting col(N), the next column shifts down into slot N.
        execute_labtalk(f"delete col({safe_col});")
        deleted += 1
    return f"Deleted {deleted} column(s) from [{safe_book}]{safe_sheet}"


@mcp.tool()
def set_column_properties(
    book_name: str,
    sheet_name: str,
    col: int,
    long_name: str = "",
    units: str = "",
    comment: str = "",
    designation: str = ""
) -> str:
    """Set a column's long name, units, comment, and/or designation.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name
        col: Column (1-based)
        long_name: Long name (label row)
        units: Units (label row)
        comment: Comment (label row)
        designation: One of none, x, y, z, yerr, xerr, label (blank = leave)

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    safe_col = positive_column(col, "col")
    require_worksheet(safe_book, safe_sheet)
    activate_window(safe_book, "book_name")
    execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet, "sheet_name")};')
    changed = []
    if long_name:
        execute_labtalk(f'wks.col{safe_col}.lname$ = {labtalk_string(long_name, "long_name")};')
        changed.append("long_name")
    if units:
        execute_labtalk(f'wks.col{safe_col}.unit$ = {labtalk_string(units, "units")};')
        changed.append("units")
    if comment:
        execute_labtalk(f'wks.col{safe_col}.comment$ = {labtalk_string(comment, "comment")};')
        changed.append("comment")
    if designation:
        safe_des = labtalk_choice(designation.lower(), _DESIGNATIONS, "designation")
        execute_labtalk(f"wks.col{safe_col}.type = {_DESIGNATIONS[safe_des]};")
        changed.append("designation")
    if not changed:
        msg = "Provide at least one of long_name, units, comment, or designation."
        raise ValueError(msg)
    return f"Updated column {safe_col} of [{safe_book}]{safe_sheet}: {', '.join(changed)}"


@mcp.tool()
def transpose_worksheet(book_name: str, sheet_name: str, output_book: str = "") -> str:
    """Transpose a worksheet (rows become columns).

    Args:
        book_name: Source workbook name
        sheet_name: Source sheet name
        output_book: Optional new workbook name for the result; if empty,
                     the source sheet is transposed in place

    Returns:
        Success message
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    src = require_worksheet(safe_book, safe_sheet)
    if output_book:
        safe_out = labtalk_name(output_book, "output_book")
        name = get_origin().CreatePage(2, safe_out, "origin")
        dest = f"[{name}]Sheet1"
    else:
        dest = src
        name = safe_book
    if not execute_labtalk(f"wtranspose iw:={src}! ow:={dest}!;"):
        msg = f"Origin could not transpose {src}."
        raise ValueError(msg)
    return f"Transposed {src} into [{name}]{'Sheet1' if output_book else safe_sheet}"


_EXPORT_DELIMITERS = {",": ",", "tab": "\t", "\t": "\t", ";": ";", "|": "|", " ": " "}


@mcp.tool()
def export_worksheet(
    book_name: str,
    sheet_name: str,
    file_path: str,
    delimiter: str = ","
) -> str:
    """Export a worksheet to a CSV/text file (with column headers).

    Args:
        book_name: Workbook name
        sheet_name: Sheet name
        file_path: Output path (Windows or WSL style)
        delimiter: ",", "tab", ";", "|", or " "

    Returns:
        Path and row/column counts written
    """
    safe_book = labtalk_name(book_name, "book_name")
    safe_sheet = labtalk_name(sheet_name, "sheet_name")
    target = require_worksheet(safe_book, safe_sheet)
    sep = _EXPORT_DELIMITERS.get(delimiter)
    if sep is None:
        msg = "delimiter must be one of: ',', 'tab', ';', '|', ' '."
        raise ValueError(msg)
    path = windows_path(file_path, "file_path")
    o = get_origin()
    data = o.GetWorksheet(target)
    if not isinstance(data, (list, tuple)):
        msg = f"Could not read {target}."
        raise ValueError(msg)
    # Header from column long names (fall back to col index).
    activate_window(safe_book, "book_name")
    execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet, "sheet_name")};')
    ncols = int(o.LTVar("wks.ncols"))
    header = []
    for i in range(1, ncols + 1):
        lname = get_lt_str(f"wks.col{i}.lname$")
        header.append(lname or f"Col{i}")
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    def _cell(v):
        if isinstance(v, (int, float)) and abs(v) >= 1e100:
            return ""  # Origin missing-value sentinel
        return v

    rows_written = 0
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, delimiter=sep)
        if header:
            writer.writerow(header)
        for row in data:
            writer.writerow([_cell(v) for v in row])
            rows_written += 1
    if not os.path.exists(path):
        msg = f"Export failed: {path} was not created."
        raise ValueError(msg)
    return f"Exported {target} ({rows_written} rows x {ncols} cols) to {path}"


@mcp.tool()
def import_excel(file_path: str, book_name: str = "") -> str:
    """Import an Excel (.xls/.xlsx) file into a new workbook.

    Args:
        file_path: Path to the Excel file (Windows or WSL style)
        book_name: Optional name for the resulting workbook

    Returns:
        Name of the workbook the data was imported into
    """
    o = get_origin()
    path = windows_path(file_path, "file_path")
    if not os.path.isfile(path):
        msg = f"File not found: {path}"
        raise ValueError(msg)
    if not execute_labtalk(f"impExcel fname:={labtalk_path(path, 'file_path')};"):
        msg = f"Origin could not import the Excel file: {path}"
        raise ValueError(msg)
    active = o.LTStr("page.name$")
    if book_name:
        safe_book = labtalk_name(book_name, "book_name")
        if active != safe_book and execute_labtalk(f"win -r {active} {safe_book};"):
            active = safe_book
    return f"Imported {os.path.basename(path)} into workbook: {active}"
