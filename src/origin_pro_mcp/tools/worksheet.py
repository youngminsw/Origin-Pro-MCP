import json
from ..app import mcp
from ..origin_connection import get_origin, execute_labtalk
from ..labtalk_safe import labtalk_choice, labtalk_name, labtalk_path, labtalk_string

@mcp.tool()
def create_worksheet(book_name: str, sheet_name: str = "Sheet1") -> str:
    """Create a new workbook with a worksheet in Origin.

    Args:
        book_name: Name for the new workbook
        sheet_name: Name for the sheet (default: Sheet1)

    Returns:
        Created workbook name
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
    target = f"[{safe_book_name}]{safe_sheet_name}"
    cols = json.loads(columns)

    for i, col_data in enumerate(cols):
        float_data = [float(x) for x in col_data]
        o.PutWorksheet(target, float_data, 0, i)

    if column_names:
        names = [n.strip() for n in column_names.split(",")]
        execute_labtalk(f"win -a {safe_book_name};")
        for i, name in enumerate(names):
            execute_labtalk(f"wks.col{i+1}.lname$ = {labtalk_string(name, 'column_names')};")

    return f"Set {len(cols)} columns x {len(cols[0])} rows in {target}"

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
    if data is None:
        return json.dumps({"error": f"Worksheet {target} not found"})

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
        file_path: Full Windows path to the file (e.g., C:\\Users\\data.csv)
        book_name: Optional workbook name. Auto-generated if empty.
        delimiter: Column delimiter (default: comma)

    Returns:
        Name of the created workbook
    """
    o = get_origin()
    if book_name:
        safe_book_name = labtalk_name(book_name, "book_name")
        o.CreatePage(2, safe_book_name, "origin")
        execute_labtalk(f"win -a {safe_book_name};")

    if delimiter == ",":
        execute_labtalk(f"impasc fname:={labtalk_path(file_path, 'file_path')} options.FileStruct.Delimiter:=1;")
    elif delimiter == "\t":
        execute_labtalk(f"impasc fname:={labtalk_path(file_path, 'file_path')} options.FileStruct.Delimiter:=0;")
    else:
        safe_delimiter = labtalk_choice(delimiter, {";", "|", " "}, "delimiter")
        execute_labtalk(
            f"impasc fname:={labtalk_path(file_path, 'file_path')} "
            f"options.FileStruct.CustomDelimiter:={labtalk_string(safe_delimiter, 'delimiter')};"
        )

    active_book = o.LTStr("page.name$")
    return f"Imported to workbook: {active_book}"

@mcp.tool()
def list_worksheets() -> str:
    """List all open workbooks and worksheets in Origin.

    Returns:
        JSON list of workbooks with their sheets
    """
    o = get_origin()
    result = []

    execute_labtalk('int __n = doc.pages();')
    n_pages = int(o.LTVar("__n"))

    for i in range(1, n_pages + 1):
        execute_labtalk(f'string __pname$ = doc.page{i}.name$;')
        execute_labtalk(f'int __ptype = doc.page{i}.type;')
        pname = o.LTStr("__pname$")
        ptype = int(o.LTVar("__ptype"))

        if ptype == 2:
            result.append({"book": pname, "type": "workbook"})
        elif ptype == 3:
            result.append({"book": pname, "type": "graph"})

    return json.dumps(result)
