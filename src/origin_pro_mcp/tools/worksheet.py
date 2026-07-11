import csv
import glob
import json
import math
import os
import re

from ..app import mcp
from ..origin_connection import (
    activate_window,
    execute_labtalk,
    get_origin,
    get_lt_str,
    graph_names,
    require_worksheet,
    safe_page_names,
    sheet_names,
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
    labtalk_text,
    windows_path,
)


@mcp.tool()
def create_worksheet(book_name: str, sheet_name: str = "Sheet1") -> str:
    """Create a worksheet in Origin: a new workbook, or a new sheet added to
    an EXISTING workbook of the given name (never a second, auto-renamed
    workbook — Origin's ``CreatePage`` on a taken name silently uniquifies
    it, which used to leave callers with two books instead of one).

    Args:
        book_name: Workbook name. If a workbook by this name is already
            open, the sheet is added to it instead of creating a new book.
        sheet_name: Name for the sheet (default: Sheet1)

    Returns:
        JSON object: {"name": <actual workbook name>, "requested_name":
        <book_name>, "renamed": <bool, true if Origin uniquified the name;
        always false when added_to_existing_book is true>, "sheet": <sheet
        name>, "added_to_existing_book": <bool>}

    Raises:
        ValueError: if book_name already has a sheet named sheet_name.
    """
    o = get_origin()
    safe_book_name = labtalk_name(book_name, "book_name")
    safe_sheet_name = labtalk_name(sheet_name, "sheet_name")

    if safe_book_name in workbook_names():
        existing_page = _find_workbook_page(o, safe_book_name)
        if safe_sheet_name in _sheet_names(o, safe_book_name, existing_page):
            msg = (
                f"Workbook '{safe_book_name}' already has a sheet named "
                f"'{safe_sheet_name}'. Use list_worksheets to see existing sheets."
            )
            raise ValueError(msg)
        activate_window(safe_book_name, "book_name")
        if not execute_labtalk(
            f'newsheet name:={labtalk_string(safe_sheet_name, "sheet_name")} cols:=2;'
        ):
            msg = f"Origin could not add sheet '{safe_sheet_name}' to workbook '{safe_book_name}'."
            raise ValueError(msg)
        return json.dumps({
            "name": safe_book_name,
            "requested_name": safe_book_name,
            "renamed": False,
            "sheet": safe_sheet_name,
            "added_to_existing_book": True,
        })

    name = o.CreatePage(2, safe_book_name, "origin")
    if safe_sheet_name != "Sheet1":
        execute_labtalk(f'page.active$ = "Sheet1"; wks.name$ = {labtalk_string(safe_sheet_name, "sheet_name")};')
    return json.dumps({
        "name": name,
        "requested_name": safe_book_name,
        "renamed": name != safe_book_name,
        "sheet": safe_sheet_name,
        "added_to_existing_book": False,
    })

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
                 Example: [[1,2,3],[4,5,6]] for 2 columns with 3 rows. A cell
                 may be JSON null (or NaN) to write an Origin missing value —
                 rendered as a gap in the polyline, not interpolated through.
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

    # (column_index, row_index) pairs, 0-based, to be written as Origin
    # missing values via LabTalk after the bulk numeric write below.
    missing_cells: list[tuple[int, int]] = []
    for i, col_data in enumerate(cols):
        float_data = []
        for r, x in enumerate(col_data):
            if x is None or (isinstance(x, float) and math.isnan(x)):
                missing_cells.append((i, r))
                float_data.append(0.0)  # placeholder, overwritten below
                continue
            try:
                float_data.append(float(x))
            except (TypeError, ValueError) as exc:
                msg = f"Column {i + 1} contains non-numeric values; only numbers are supported."
                raise ValueError(msg) from exc
        if not o.PutWorksheet(target, float_data, 0, i):
            msg = f"Origin rejected the data for column {i + 1} of {target}."
            raise ValueError(msg)

    if missing_cells:
        # The bulk PutWorksheet above already wrote a numeric 0.0 placeholder
        # into every null/NaN cell. Each must be overwritten with Origin's real
        # missing-value sentinel (col(c)[r]=0/0). If activation or ANY per-cell
        # write fails, that cell silently keeps the 0.0 — a wrong number wearing
        # a success message, the worst class of bug — so check every return and
        # raise, naming exactly which cells still hold the placeholder.
        def _cells(pairs):
            return ", ".join(f"col {c + 1} row {r + 1}" for c, r in pairs)

        activate_window(safe_book_name, "book_name")
        if not execute_labtalk(
            f'page.active$ = {labtalk_string(safe_sheet_name, "sheet_name")};'
        ):
            msg = (
                f"Could not activate sheet '{safe_sheet_name}' of "
                f"'{safe_book_name}' to write missing values. The bulk write left "
                f"a numeric 0.0 placeholder in these cells that should be blank: "
                f"{_cells(missing_cells)}. Treat that data as wrong."
            )
            raise ValueError(msg)
        failed = []
        for col_idx, row_idx in missing_cells:
            if not execute_labtalk(f"col({col_idx + 1})[{row_idx + 1}]=0/0;"):
                failed.append((col_idx, row_idx))
        if failed:
            msg = (
                f"Origin rejected the missing-value write for these cells of "
                f"{target}: {_cells(failed)}. They still hold the numeric 0.0 "
                f"placeholder from the bulk write instead of a blank/missing "
                f"value — treat that data as wrong."
            )
            raise ValueError(msg)

    if column_names:
        names = [n.strip() for n in column_names.split(",")]
        activate_window(safe_book_name, "book_name")
        execute_labtalk(f'page.active$ = {labtalk_string(safe_sheet_name, "sheet_name")};')
        for i, name in enumerate(names):
            execute_labtalk(f"wks.col{i+1}.lname$ = {labtalk_text(name, 'column_names')};")

    n_rows = max(len(c) for c in cols)
    return f"Set {len(cols)} columns x {n_rows} rows in {target}"

@mcp.tool()
def get_worksheet_data(book_name: str, sheet_name: str) -> str:
    """Read data from an Origin worksheet.

    Args:
        book_name: Workbook name
        sheet_name: Sheet name

    Returns:
        JSON object with column data; empty cells (including a shorter
        ragged column padded by Origin) are returned as null.

    Raises:
        ValueError: if the worksheet is not found.
    """
    o = get_origin()
    safe_book_name = labtalk_name(book_name, "book_name")
    safe_sheet_name = labtalk_name(sheet_name, "sheet_name")
    target = f"[{safe_book_name}]{safe_sheet_name}"
    data = o.GetWorksheet(target)
    # On failure GetWorksheet returns None or an HRESULT int, never a sequence
    if not isinstance(data, (list, tuple)):
        books = ", ".join(workbook_names()) or "(none)"
        msg = f"Worksheet {target} not found. Open workbooks: {books}."
        raise ValueError(msg)

    if len(data) == 0:
        return json.dumps({"columns": []})

    def _cell(v):
        # Origin's missing-value sentinel (~1.2e308), same convention as
        # get_matrix_data / export_worksheet.
        if isinstance(v, (int, float)) and abs(v) >= 1e100:
            return None
        return v

    num_cols = len(data[0])
    columns = []
    for c in range(num_cols):
        columns.append([_cell(row[c]) for row in data])

    return json.dumps({"columns": columns})

def _import_csv_to_worksheet_impl(
    file_path: str,
    book_name: str = "",
    delimiter: str = ",",
    sparklines: bool = False,
) -> str:
    """Import a CSV/text file into an Origin worksheet.

    Args:
        file_path: Path to the file (Windows or WSL style, e.g.
                   C:\\Users\\data.csv or /mnt/c/Users/data.csv)
        book_name: Optional workbook name. Auto-generated if empty.
        delimiter: Column delimiter (default: comma)
        sparklines: Whether to allow Origin's auto-generated sparkline mini-
            graph windows for this import (default False = suppressed).
            ASCII/CSV import can otherwise spawn a dozen-plus throwaway graph
            windows per import, bloating the project. Suppression is
            attempted at the source (an ImpASC option) and, regardless of
            whether that took effect, any graph window that appears as a
            side effect of this specific import call is deleted afterward.

    Returns:
        JSON object: {"name": <actual workbook name>, "requested_name":
        <book_name if given, else null>, "renamed": <bool, true if the
        actual name differs from the requested name>, "file": <file_path>,
        "sparklines_suppressed": <bool, true when suppression actually worked —
        the options.Sparklines:=0 import option ran AND left no sparkline graph
        windows behind; cannot be true while sparklines_deleted > 0>,
        "sparklines_deleted": <int, sparkline graph windows the post-import
        cleanup still had to remove>}
    """
    o = get_origin()
    path = windows_path(file_path, "file_path")
    if not os.path.isfile(path):
        msg = f"File not found: {path}"
        raise ValueError(msg)

    requested_name = None
    if book_name:
        requested_name = labtalk_name(book_name, "book_name")
        # CreatePage UNIQUIFIES a taken name (e.g. "Data" -> "Data1"), so its
        # RETURN is the actual new book. The old code discarded it and did
        # `win -a <requested>`, which activates the pre-existing book of that
        # name — impasc (which imports into the ACTIVE window) then landed the
        # data in the WRONG book. Activate the actual new book instead.
        actual_book = o.CreatePage(2, requested_name, "origin")
        activate_window(actual_book, "book_name")

    if delimiter == ",":
        delim_clause = "options.FileStruct.Delimiter:=1"
    elif delimiter == "\t":
        delim_clause = "options.FileStruct.Delimiter:=0"
    else:
        safe_delimiter = labtalk_choice(delimiter, {";", "|", " "}, "delimiter")
        delim_clause = (
            f"options.FileStruct.CustomDelimiter:={labtalk_string(safe_delimiter, 'delimiter')}"
        )

    graphs_before = set(graph_names()) if not sparklines else set()
    sparklines_option_ok = False
    if sparklines:
        ok = execute_labtalk(f"impasc fname:={labtalk_path(path, 'file_path')} {delim_clause};")
    else:
        # `options.Sparklines:=0` suppresses the per-column sparkline graph
        # windows an ASCII import otherwise spawns (live-verified: a 12-column
        # CSV spawns 12 graph windows unsuppressed, 0 with this key). The
        # earlier `options.Miscellaneous.Sparklines:=0` was the WRONG key —
        # Origin rejected it (Execute returned False), so suppression never ran
        # at the source and only the cleanup below saved the project. If the
        # key is ever rejected again we retry without it and the cleanup is the
        # backstop.
        ok = execute_labtalk(
            f"impasc fname:={labtalk_path(path, 'file_path')} {delim_clause} "
            "options.Sparklines:=0;"
        )
        if ok:
            sparklines_option_ok = True
        else:
            ok = execute_labtalk(f"impasc fname:={labtalk_path(path, 'file_path')} {delim_clause};")
    if not ok:
        msg = f"Origin could not import {path} — check the file format and delimiter."
        raise ValueError(msg)

    sparklines_deleted = 0
    if not sparklines:
        # Defensive net: delete only graph windows that appeared as a result
        # of THIS import call (never a pre-existing/user window). Explicit
        # per-name `win -cd <name>;` commands, not a `win -cd %()` loop — the
        # latter is unreliable inside LabTalk loops.
        new_graphs = [g for g in graph_names() if g not in graphs_before]
        for g in new_graphs:
            execute_labtalk(f"win -cd {g};")
        sparklines_deleted = len(new_graphs)

    # Report ACTUAL suppression, so the two telemetry fields can't contradict
    # each other (usability F12): suppression "succeeded" only when the source
    # option ran AND left no sparkline graph windows for the cleanup to remove.
    sparklines_suppressed = sparklines_option_ok and sparklines_deleted == 0

    active_book = o.LTStr("page.name$")
    return json.dumps({
        "name": active_book,
        "requested_name": requested_name,
        "renamed": bool(requested_name) and active_book != requested_name,
        "file": path,
        "sparklines_suppressed": sparklines_suppressed,
        "sparklines_deleted": sparklines_deleted,
    })

def _find_workbook_page(o, book_name: str):
    """The WorksheetPages COM item for book_name, or None if not open."""
    pages = o.WorksheetPages
    try:
        count = pages.Count
    except Exception:
        count = 0
    for i in range(count):
        try:
            page = pages.Item(i)
            if page.Name == book_name:
                return page
        except Exception:
            continue
    return None


def _sheet_names(o, book_name: str, page=None) -> list:
    """Sheet names of a workbook: the shared crash-safe LabTalk enumeration
    (``origin_connection.sheet_names``), with an isolated per-sheet COM fallback
    used only when LabTalk yields nothing (test fakes / odd COM builds)."""
    names = sheet_names(book_name)
    if names:
        return names
    if page is None:
        return []
    out: list = []
    try:
        layers = page.Layers
        for j in range(layers.Count):
            try:
                out.append(layers.Item(j).Name)
            except Exception:
                continue
    except Exception:
        pass
    return out


@mcp.tool()
def list_worksheets() -> str:
    """List all open workbooks (with their sheets), graphs, and matrices.

    Returns:
        JSON object: {"workbooks": [{"name", "sheets"}], "graphs": [names],
        "matrices": [names]}
    """
    o = get_origin()
    # Save the active window: enumeration activates each workbook to read its
    # sheets via LabTalk, so restore the user's active window afterward.
    saved_active = ""
    try:
        o.Execute("string _opm_act$=%H;")
        saved_active = o.LTStr("_opm_act$") or ""
    except Exception:
        saved_active = ""
    try:
        workbooks = []
        pages = o.WorksheetPages
        try:
            count = pages.Count
        except Exception:
            count = 0
        for i in range(count):
            try:
                page = pages.Item(i)
                name = page.Name
            except Exception:
                continue  # skip a workbook whose name can't be read
            workbooks.append({"name": name, "sheets": _sheet_names(o, name, page)})
        return json.dumps({
            "workbooks": workbooks,
            "graphs": safe_page_names(o.GraphPages),
            "matrices": safe_page_names(o.MatrixPages),
        })
    finally:
        if saved_active:
            try:
                o.Execute(f"win -a {saved_active};")
            except Exception:
                pass


# LabTalk `wks.col.type` designation codes (verified against OriginLab docs):
# 1=Y, 2=Disregard, 3=Y Error, 4=X, 5=Label, 6=Z, 7=X Error. (0 clears the
# designation to "none".) Earlier `xerr` was wrongly 5 — the same code as
# `label` — so an X-error column was silently designated as a label.
_DESIGNATIONS = {
    "none": 0, "disregard": 2, "y": 1, "yerr": 3, "x": 4, "label": 5,
    "z": 6, "xerr": 7,
}


def _set_column_formula_impl(book_name: str, sheet_name: str, col: int, formula: str) -> str:
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


def _add_columns_impl(book_name: str, sheet_name: str, count: int = 1) -> str:
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
        if not execute_labtalk("wks.addCol();"):
            msg = f"Origin could not add a column to [{safe_book}]{safe_sheet}."
            raise ValueError(msg)
    return f"Added {safe_count} column(s) to [{safe_book}]{safe_sheet}"


def _delete_columns_impl(book_name: str, sheet_name: str, col: int, count: int = 1) -> str:
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


def _set_column_properties_impl(
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
    ref = f"[{safe_book}]{safe_sheet}"
    if long_name:
        if not execute_labtalk(f'wks.col{safe_col}.lname$ = {labtalk_text(long_name, "long_name")};'):
            msg = f"Origin could not set the long name of column {safe_col} of {ref}."
            raise ValueError(msg)
        changed.append("long_name")
    if units:
        if not execute_labtalk(f'wks.col{safe_col}.unit$ = {labtalk_text(units, "units")};'):
            msg = f"Origin could not set the units of column {safe_col} of {ref}."
            raise ValueError(msg)
        changed.append("units")
    if comment:
        if not execute_labtalk(f'wks.col{safe_col}.comment$ = {labtalk_text(comment, "comment")};'):
            msg = f"Origin could not set the comment of column {safe_col} of {ref}."
            raise ValueError(msg)
        changed.append("comment")
    if designation:
        safe_des = labtalk_choice(designation.lower(), _DESIGNATIONS, "designation")
        if not execute_labtalk(f"wks.col{safe_col}.type = {_DESIGNATIONS[safe_des]};"):
            msg = f"Origin could not set the designation of column {safe_col} of {ref}."
            raise ValueError(msg)
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


def _import_excel_impl(file_path: str, book_name: str = "") -> str:
    """Import an Excel (.xls/.xlsx) file into a new workbook.

    Args:
        file_path: Path to the Excel file (Windows or WSL style)
        book_name: Optional name for the resulting workbook

    Returns:
        JSON object: {"name": <actual workbook name>, "requested_name":
        <book_name if given, else null>, "renamed": <bool, true if the
        actual name differs from the requested name>, "file": <file_path>}
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
    requested_name = None
    if book_name:
        requested_name = labtalk_name(book_name, "book_name")
        if active != requested_name and execute_labtalk(f"win -r {active} {requested_name};"):
            active = requested_name
    return json.dumps({
        "name": active,
        "requested_name": requested_name,
        "renamed": bool(requested_name) and active != requested_name,
        "file": path,
    })


# --- Consolidated dispatchers (Phase 2) ---------------------------------------

_MANAGE_COLUMNS_OPS = {"add", "delete", "properties", "formula"}


@mcp.tool()
def manage_columns(
    book_name: str,
    sheet_name: str,
    op: str,
    col: int | None = None,
    count: int = 1,
    long_name: str | None = None,
    units: str | None = None,
    comment: str | None = None,
    designation: str | None = None,
    formula: str | None = None,
) -> str:
    """Add, delete, or edit worksheet columns.

    Args:
        book_name, sheet_name: Target workbook and sheet.
        op: Which column operation to run:
            - "add": append `count` empty columns (default 1).
            - "delete": delete `count` columns starting at `col` (col required).
            - "properties": set a column's metadata (col required). Uses
              `long_name`, `units`, `comment`, and/or `designation`
              (none/x/y/z/yerr/xerr/label); at least one is required.
            - "formula": fill `col` from `formula` (col and formula required),
              e.g. "col(1)^2".
        col: Target column (1-based; required for delete/properties/formula).
        count: Column count (op="add" or "delete"; default 1).
        long_name, units, comment, designation: Column metadata
            (op="properties").
        formula: LabTalk column expression (op="formula").

    Returns:
        Success message for the selected operation.
    """
    safe_op = labtalk_choice(op.lower(), _MANAGE_COLUMNS_OPS, "op")
    if safe_op == "add":
        return _add_columns_impl(book_name, sheet_name, count)
    if safe_op == "delete":
        if col is None:
            msg = "manage_columns op 'delete' requires col."
            raise ValueError(msg)
        return _delete_columns_impl(book_name, sheet_name, col, count)
    if safe_op == "properties":
        if col is None:
            msg = "manage_columns op 'properties' requires col."
            raise ValueError(msg)
        return _set_column_properties_impl(
            book_name, sheet_name, col,
            long_name=long_name or "",
            units=units or "",
            comment=comment or "",
            designation=designation or "",
        )
    # formula
    if col is None or formula is None:
        msg = "manage_columns op 'formula' requires col and formula."
        raise ValueError(msg)
    return _set_column_formula_impl(book_name, sheet_name, col, formula)


_IMPORT_FORMATS = {"auto", "csv", "excel"}
_EXCEL_EXTENSIONS = (".xls", ".xlsx", ".xlsm")
# Data files a folder/glob import will pick up.
_BATCH_DATA_EXTENSIONS = (".csv", ".txt", ".dat", ".xls", ".xlsx", ".xlsm")
# Import at most this many files per batch call (time/context guard).
_BATCH_CAP = 20


def _is_batch_target(file_path: str) -> bool:
    """True when file_path names a DIRECTORY or a glob pattern (so import_data
    should batch-import many files rather than one)."""
    if any(ch in file_path for ch in "*?[]"):
        return True
    try:
        win = windows_path(file_path, "file_path")
    except ValueError:
        return False
    return os.path.isdir(win)


def _book_name_from_stem(stem: str) -> str:
    """A valid Origin book name derived from a file stem: non-identifier chars
    become underscores, and a leading non-letter/underscore is prefixed."""
    safe = re.sub(r"[^A-Za-z0-9_]", "_", stem) or "Book"
    if not (safe[0].isalpha() or safe[0] == "_"):
        safe = "B_" + safe
    return safe[:60]


def _import_batch_impl(
    file_path: str, format: str, delimiter: str, sparklines: bool
) -> str:
    """Import every data file in a directory / matching a glob, each into its
    own book named from the file stem. Per-file partial failures are reported,
    not fatal. Caps at _BATCH_CAP files."""
    win = windows_path(file_path, "file_path")
    if os.path.isdir(win):
        matches = glob.glob(os.path.join(win, "*"))
    else:
        matches = glob.glob(win)
    files = sorted(
        f for f in matches
        if os.path.isfile(f)
        and os.path.splitext(f)[1].lower() in _BATCH_DATA_EXTENSIONS
    )
    if not files:
        msg = (
            f"No data files (csv/txt/dat/xls/xlsx) found for {win}. "
            "Pass a folder or a glob like '.../*.csv'."
        )
        raise ValueError(msg)
    matched = len(files)
    capped = matched > _BATCH_CAP
    files = files[:_BATCH_CAP]
    results = []
    for f in files:
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            single = import_data(
                f, format=format, book_name=_book_name_from_stem(stem),
                delimiter=delimiter, sparklines=sparklines,
            )
            results.append(
                {"file": f, "name": json.loads(single).get("name"), "ok": True}
            )
        except (ValueError, RuntimeError) as exc:
            results.append({"file": f, "error": str(exc), "ok": False})
    payload = {
        "batch": True,
        "matched": matched,
        "imported": sum(1 for r in results if r["ok"]),
        "results": results,
    }
    if capped:
        payload["note"] = (
            f"{matched} files matched; imported the first {_BATCH_CAP} "
            "(cap to guard time/context). Narrow the pattern for the rest."
        )
    return json.dumps(payload)


@mcp.tool()
def import_data(
    file_path: str,
    format: str = "auto",
    book_name: str = "",
    delimiter: str = ",",
    sparklines: bool = False,
) -> str:
    """Import a data file into an Origin worksheet.

    Args:
        file_path: Path to the file (Windows or WSL style). May ALSO be a
            DIRECTORY or a glob pattern (e.g. C:\\data or C:\\data\\*.csv) to
            batch-import every data file (csv/txt/dat/xls/xlsx) it matches,
            each into its own book named from the file stem.
        format: "auto" (detect by extension; default), "csv" (text/CSV), or
            "excel" (.xls/.xlsx). "auto" treats .xls/.xlsx/.xlsm as Excel and
            everything else as CSV/text.
        book_name: Optional workbook name for the result. Ignored in batch mode
            (each file is named from its stem).
        delimiter: Column delimiter for CSV/text (default comma; ignored for
            Excel).
        sparklines: CSV/text import only. Whether to allow Origin's auto-
            generated sparkline mini-graph windows (default False =
            suppressed — an unsuppressed CSV import can spawn a dozen-plus
            throwaway graph windows, bloating the project). Applied per file in
            batch mode.

    Returns:
        Single file: JSON object {"name": <actual workbook name>,
        "requested_name": <book_name if given, else null>, "renamed": <bool,
        true if Origin gave the workbook a different name>, "file":
        <file_path>}; CSV/text also adds {"sparklines_suppressed": <bool>,
        "sparklines_deleted": <int>}.
        Batch (directory/glob): JSON object {"batch": true, "matched": <int
        files matched>, "imported": <int succeeded>, "results": [{"file",
        "name"/"error", "ok"}, ...]} — partial failures are reported per file,
        not fatal. If more than 20 files match, only the first 20 are imported
        and a "note" says so.
    """
    if _is_batch_target(file_path):
        return _import_batch_impl(file_path, format, delimiter, sparklines)
    safe_format = labtalk_choice(format.lower(), _IMPORT_FORMATS, "format")
    if safe_format == "auto":
        ext = os.path.splitext(file_path)[1].lower()
        safe_format = "excel" if ext in _EXCEL_EXTENSIONS else "csv"
    if safe_format == "excel":
        return _import_excel_impl(file_path, book_name)
    return _import_csv_to_worksheet_impl(file_path, book_name, delimiter, sparklines)
