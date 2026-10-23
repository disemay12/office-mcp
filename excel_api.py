"""Excel automation exposed as MCP tools."""

from __future__ import annotations

import builtins
import os
from typing import Any

import pywintypes

from com import (
    OfficeError,
    cm_to_points,
    default_output,
    file_info,
    hex_to_bgr,
    is_index,
    normalize_matrix,
    prepare_output,
    resolve_path,
    run,
)

CHART_TYPES = {
    "column": 51,  # xlColumnClustered
    "bar": 57,  # xlBarClustered
    "line": 4,  # xlLine
    "line_markers": 65,  # xlLineMarkers
    "pie": 5,  # xlPie
    "doughnut": -4120,  # xlDoughnut
    "area": 1,  # xlArea
    "scatter": -4169,  # xlXYScatter
    "radar": -4151,  # xlRadar
}

ALIGNMENTS = {
    "left": -4131,  # xlHAlignLeft
    "center": -4108,  # xlHAlignCenter
    "right": -4152,  # xlHAlignRight
    "justify": -4130,
}


def _wb(app: Any, identifier: str | None) -> Any:
    books = app.Workbooks
    if books.Count == 0:
        raise OfficeError("No Excel workbook is open. Use excel_create or excel_open first.")
    if identifier is None or str(identifier).strip() == "":
        try:
            return app.ActiveWorkbook
        except pywintypes.com_error:
            return books(1)

    ident = str(identifier).strip()
    if is_index(ident):
        index = int(ident)
        if not 1 <= index <= books.Count:
            raise OfficeError(f"Workbook index {index} is out of range (1..{books.Count}).")
        return books(index)

    wanted_path = os.path.normcase(os.path.abspath(ident)) if os.path.isabs(ident) else None
    lowered = ident.lower()
    for book in books:
        try:
            if wanted_path and os.path.normcase(os.path.abspath(book.FullName)) == wanted_path:
                return book
            if book.Name.lower() == lowered:
                return book
        except pywintypes.com_error:
            continue
    names = []
    for book in books:
        try:
            names.append(book.Name)
        except pywintypes.com_error:
            pass
    hint = ""
    if os.path.isfile(ident):
        hint = " The file exists but is not open — call excel_open first."
    raise OfficeError(
        f"Workbook '{identifier}' is not open.{hint} Open workbooks: {', '.join(names) or 'none'}."
    )


def _ws(app: Any, workbook: str | None, sheet: str | None) -> Any:
    book = _wb(app, workbook)
    if sheet is None or str(sheet).strip() == "":
        return book.ActiveSheet

    ident = str(sheet).strip()
    sheets = book.Worksheets
    if is_index(ident):
        index = int(ident)
        if not 1 <= index <= sheets.Count:
            raise OfficeError(f"Sheet index {index} is out of range (1..{sheets.Count}).")
        return sheets(index)
    for candidate in sheets:
        if candidate.Name.lower() == ident.lower():
            return candidate
    names = [candidate.Name for candidate in sheets]
    raise OfficeError(f"Sheet '{sheet}' not found. Sheets: {', '.join(names)}.")


def _range(ws: Any, address: str | None) -> Any:
    if address is None or str(address).strip() == "":
        return ws.UsedRange
    try:
        return ws.Range(str(address).strip())
    except pywintypes.com_error as exc:
        raise OfficeError(f"Invalid range '{address}': {exc}") from exc


def _matrix(value: Any) -> list[list[Any]]:
    if isinstance(value, tuple):
        value = list(value)
    if isinstance(value, list):
        if not value:
            return []
        if not any(isinstance(item, (list, tuple)) for item in value):
            return [list(value)]
        return [list(row) if isinstance(row, (list, tuple)) else [row] for row in value]
    return [[value]]


def _column_letter(number: int) -> str:
    letters = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _address(target: Any) -> str:
    """A1 address without '$'."""
    try:
        return str(target.Address).replace("$", "")
    except pywintypes.com_error:
        # Fallback for Excel builds that reject the parameterized getter.
        row = int(target.Row)
        column = int(target.Column)
        rows = int(target.Rows.Count)
        columns = int(target.Columns.Count)
        if rows == 1 and columns == 1:
            return f"{_column_letter(column)}{row}"
        return (
            f"{_column_letter(column)}{row}:"
            f"{_column_letter(column + columns - 1)}{row + rows - 1}"
        )


def _apply_number_format(target: Any, number_format: str) -> None:
    """Set a number format, tolerating localized Excel format codes.

    Excel installed with a Russian (or other) locale rejects en-US codes such as
    '#,##0.00' through .NumberFormat. Try the given text, then a translated
    variant, then the local property.
    """
    candidates = [number_format]
    translated = number_format.replace(",", "\x00").replace(".", ",").replace("\x00", " ")
    if translated != number_format:
        candidates.append(translated)
    swapped = number_format.replace(".", "\x00").replace(",", ".").replace("\x00", ",")
    if swapped not in candidates:
        candidates.append(swapped)

    for candidate in candidates:
        for attribute in ("NumberFormat", "NumberFormatLocal"):
            try:
                setattr(target, attribute, candidate)
                return
            except pywintypes.com_error:
                continue
    raise OfficeError(
        f"Cannot apply number format '{number_format}'. This Excel uses localized format "
        "codes; pass the local variant, e.g. '# ##0,00' for Russian Excel."
    )


def _cell_text(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


# -- workbooks ----------------------------------------------------------


def excel_list_workbooks() -> dict[str, Any]:
    """List open Excel workbooks with sheet names, saved state and used ranges."""

    def op(app: Any) -> dict[str, Any]:
        items = []
        active = None
        try:
            active = app.ActiveWorkbook.Name
        except pywintypes.com_error:
            pass
        for book in app.Workbooks:
            sheets = []
            for sheet in book.Worksheets:
                entry = {"name": sheet.Name}
                try:
                    entry["used_range"] = _address(sheet.UsedRange)
                except pywintypes.com_error:
                    entry["used_range"] = None
                sheets.append(entry)
            items.append(
                {
                    "name": book.Name,
                    "path": book.FullName if book.Path else None,
                    "saved": bool(book.Saved),
                    "active": book.Name == active,
                    "sheets": sheets,
                }
            )
        return {"count": len(items), "workbooks": items}

    return run("excel", op, create=False)


def excel_create(path: str | None = None, visible: bool = True) -> dict[str, Any]:
    """Create a new Excel workbook, optionally saved at `path` (.xlsx/.xlsm/.xls).

    Without `path` the workbook stays unsaved; excel_save picks an automatic file.
    """

    def op(app: Any) -> dict[str, Any]:
        target = prepare_output(path) if path else None
        if target and os.path.exists(target):
            raise OfficeError(f"File already exists: {target}. Use excel_open or choose another path.")
        book = app.Workbooks.Add()
        if target:
            book.SaveAs(target)
        return {
            "name": book.Name,
            "path": book.FullName if book.Path else None,
            "sheets": [sheet.Name for sheet in book.Worksheets],
            "saved": bool(book.Saved),
        }

    return run("excel", op, visible=visible)


def excel_open(path: str, visible: bool = True, read_only: bool = False) -> dict[str, Any]:
    """Open an Excel workbook (or focus it if already open) and return its sheets."""
    source = resolve_path(path, must_exist=True)

    def op(app: Any) -> dict[str, Any]:
        for book in app.Workbooks:
            try:
                if os.path.normcase(book.FullName) == os.path.normcase(source):
                    book.Activate()
                    return {
                        "name": book.Name,
                        "path": book.FullName,
                        "sheets": [sheet.Name for sheet in book.Worksheets],
                        "message": "Workbook was already open",
                    }
            except pywintypes.com_error:
                continue
        book = app.Workbooks.Open(source, 0, bool(read_only))
        book.Activate()
        return {
            "name": book.Name,
            "path": book.FullName,
            "sheets": [sheet.Name for sheet in book.Worksheets],
            "message": "Workbook opened",
        }

    return run("excel", op, visible=visible)


def excel_list_sheets(workbook: str | None = None) -> dict[str, Any]:
    """List worksheets in a workbook with used range, row/column counts and visibility."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        items = []
        for index, sheet in enumerate(book.Worksheets, start=1):
            used = sheet.UsedRange
            items.append(
                {
                    "index": index,
                    "name": sheet.Name,
                    "visible": int(sheet.Visible) == -1,
                    "used_range": _address(used),
                    "rows": int(used.Rows.Count),
                    "columns": int(used.Columns.Count),
                }
            )
        return {"workbook": book.Name, "count": len(items), "sheets": items}

    return run("excel", op, create=False)


def excel_read_range(
    range: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    formatted: bool = False,
    include_formulas: bool = False,
    max_rows: int = 200,
    max_cols: int = 50,
) -> dict[str, Any]:
    """Read cells from an Excel worksheet as a 2-D array.

    `range` is an A1 address ('A1:D20'); omit it to read the used range. Empty
    cells come back as null. `formatted=True` returns displayed text (good for
    dates and currency), otherwise raw numbers. Large ranges are truncated to
    `max_rows` x `max_cols` and flagged with truncated=true.
    """

    row_cap = max(1, min(int(max_rows), 1000))
    col_cap = max(1, min(int(max_cols), 200))

    def op(app: Any) -> dict[str, Any]:
        ws = _ws(app, workbook, sheet)
        target = _range(ws, range)
        total_rows = int(target.Rows.Count)
        total_cols = int(target.Columns.Count)
        rows = min(total_rows, row_cap)
        cols = min(total_cols, col_cap)
        if rows != total_rows or cols != total_cols:
            target = ws.Range(
                target.Cells(1, 1), target.Cells(rows, cols)
            )
        if formatted:
            values = [
                [target.Cells(row_index, col_index).Text for col_index in builtins.range(1, cols + 1)]
                for row_index in builtins.range(1, rows + 1)
            ]
        else:
            values = _matrix(target.Value2)
        values = [[_cell_text(cell) for cell in row] for row in values]
        payload: dict[str, Any] = {
            "workbook": ws.Parent.Name,
            "sheet": ws.Name,
            "address": _address(target),
            "rows": rows,
            "columns": cols,
            "truncated": rows != total_rows or cols != total_cols,
            "values": values,
        }
        if include_formulas:
            formulas = _matrix(target.Formula)
            payload["formulas"] = [
                [None if cell is None else str(cell) for cell in row] for row in formulas
            ]
        return payload

    return run("excel", op, create=False)


def excel_write_range(
    values: list[list[Any]] | list[Any],
    range: str = "A1",
    sheet: str | None = None,
    workbook: str | None = None,
    clear_first: bool = False,
    autofit: bool = False,
) -> dict[str, Any]:
    """Write a 2-D array into a worksheet starting at the top-left of `range`.

    Strings beginning with '=' become formulas; null clears the cell. A single
    row may be passed as a flat list. Use `clear_first=True` to wipe the target
    rectangle before writing.
    """

    matrix = normalize_matrix(values)
    if not matrix:
        raise OfficeError("values is empty")
    nrows, ncols = len(matrix), len(matrix[0])

    def op(app: Any) -> dict[str, Any]:
        ws = _ws(app, workbook, sheet)
        anchor = _range(ws, range).Cells(1, 1)
        target = ws.Range(anchor, ws.Cells(anchor.Row + nrows - 1, anchor.Column + ncols - 1))
        if clear_first:
            target.ClearContents()

        # Range.Formula accepts numbers, text, booleans, dates, formulas and
        # None (clears the cell) in one assignment. That replaces a COM call
        # per formula and per blank cell.
        bulk: list[list[Any]] = []
        formulas = 0
        for row in matrix:
            out_row = []
            for value in row:
                if isinstance(value, str) and value.startswith("="):
                    formulas += 1
                out_row.append(value)
            bulk.append(out_row)

        try:
            target.Formula = tuple(tuple(row) for row in bulk)
        except pywintypes.com_error as exc:
            raise OfficeError(f"Cannot write values: {exc}") from exc

        if autofit:
            try:
                ws.UsedRange.Columns.AutoFit()
            except pywintypes.com_error:
                pass
        return {
            "workbook": ws.Parent.Name,
            "sheet": ws.Name,
            "address": _address(target),
            "rows": nrows,
            "columns": ncols,
            "formulas": formulas,
        }

    return run("excel", op, create=False)


def excel_format_range(
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    number_format: str | None = None,
    bg_color: str | None = None,
    font_color: str | None = None,
    align: str | None = None,
    wrap: bool | None = None,
    borders: str | None = None,
    column_width: float | None = None,
    row_height: float | None = None,
) -> dict[str, Any]:
    """Format a cell range: fonts, colours ('#RRGGBB'), number format, alignment, borders, sizes.

    `borders` accepts 'all', 'outline' or 'none'. `align` accepts left/center/right/justify.
    """

    def op(app: Any) -> dict[str, Any]:
        ws = _ws(app, workbook, sheet)
        target = _range(ws, range)
        if bold is not None:
            target.Font.Bold = bool(bold)
        if italic is not None:
            target.Font.Italic = bool(italic)
        if font_size is not None:
            target.Font.Size = float(font_size)
        if font_color is not None:
            target.Font.Color = hex_to_bgr(font_color)
        if bg_color is not None:
            target.Interior.Color = hex_to_bgr(bg_color)
        if number_format is not None:
            _apply_number_format(target, number_format)
        if align is not None:
            key = str(align).strip().lower()
            if key not in ALIGNMENTS:
                raise OfficeError("align must be left, center, right or justify")
            target.HorizontalAlignment = ALIGNMENTS[key]
        if wrap is not None:
            target.WrapText = bool(wrap)
        if borders is not None:
            key = str(borders).strip().lower()
            if key == "all":
                target.Borders.LineStyle = 1  # xlContinuous
            elif key == "outline":
                target.BorderAround(1)
            elif key == "none":
                target.Borders.LineStyle = -4142  # xlNone
            else:
                raise OfficeError("borders must be all, outline or none")
        if column_width is not None:
            target.EntireColumn.ColumnWidth = float(column_width)
        if row_height is not None:
            target.EntireRow.RowHeight = float(row_height)
        return {"workbook": ws.Parent.Name, "sheet": ws.Name, "address": _address(target)}

    return run("excel", op, create=False)


def excel_add_sheet(
    name: str,
    workbook: str | None = None,
    activate: bool = True,
) -> dict[str, Any]:
    """Add a worksheet at the end of a workbook and optionally activate it."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        for sheet in book.Worksheets:
            if sheet.Name.lower() == name.lower():
                raise OfficeError(f"Sheet '{name}' already exists.")
        last = book.Worksheets(book.Worksheets.Count)
        created = book.Worksheets.Add(None, last)
        created.Name = name
        if activate:
            created.Activate()
        return {"workbook": book.Name, "sheet": created.Name, "sheets": [s.Name for s in book.Worksheets]}

    return run("excel", op, create=False)


def excel_delete_sheet(name: str, workbook: str | None = None) -> dict[str, Any]:
    """Delete a worksheet by name. The last remaining sheet cannot be deleted."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        ws = _ws(app, book.Name, name)
        if book.Worksheets.Count <= 1:
            raise OfficeError("Cannot delete the only worksheet in a workbook.")
        sheet_name = ws.Name
        ws.Delete()
        return {"workbook": book.Name, "deleted": sheet_name, "sheets": [s.Name for s in book.Worksheets]}

    return run("excel", op, create=False)


def excel_rename_sheet(
    old_name: str,
    new_name: str,
    workbook: str | None = None,
) -> dict[str, Any]:
    """Rename a worksheet."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        ws = _ws(app, book.Name, old_name)
        for sheet in book.Worksheets:
            if sheet.Name.lower() == new_name.lower():
                raise OfficeError(f"Sheet '{new_name}' already exists.")
        ws.Name = new_name
        return {"workbook": book.Name, "sheet": ws.Name}

    return run("excel", op, create=False)


def excel_add_chart(
    range: str,
    chart_type: str = "column",
    title: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    anchor: str = "A1",
    width_cm: float = 16,
    height_cm: float = 8,
) -> dict[str, Any]:
    """Insert a chart for a data range. Types: column, bar, line, line_markers, pie, doughnut, area, scatter, radar.

    The chart is placed at `anchor` on the same sheet; its top-left corner uses
    the anchor cell position.
    """

    key = str(chart_type).strip().lower()
    if key not in CHART_TYPES:
        raise OfficeError(f"Unknown chart_type '{chart_type}'. Use: {', '.join(CHART_TYPES)}")

    def op(app: Any) -> dict[str, Any]:
        ws = _ws(app, workbook, sheet)
        source = _range(ws, range)
        corner = _range(ws, anchor).Cells(1, 1)
        shape = ws.Shapes.AddChart2(
            -1,
            CHART_TYPES[key],
            corner.Left,
            corner.Top,
            cm_to_points(width_cm),
            cm_to_points(height_cm),
        )
        shape.Chart.SetSourceData(source)
        if title:
            shape.Chart.HasTitle = True
            shape.Chart.ChartTitle.Text = title
        return {
            "workbook": ws.Parent.Name,
            "sheet": ws.Name,
            "chart": shape.Name,
            "source": _address(source),
            "chart_type": key,
        }

    return run("excel", op, create=False)


def excel_add_table(
    range: str,
    name: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    style: str = "TableStyleMedium2",
) -> dict[str, Any]:
    """Turn a range into an Excel table (ListObject) with a header row and banded style.

    A single-cell `range` expands to the surrounding data block.
    """

    def op(app: Any) -> dict[str, Any]:
        ws = _ws(app, workbook, sheet)
        target = _range(ws, range)
        if int(target.Cells.Count) == 1:
            target = target.CurrentRegion
        table = ws.ListObjects.Add(1, target, None, 1)  # xlSrcRange, xlYes (has headers)
        if name:
            table.Name = name
        if style:
            try:
                table.TableStyle = style
            except pywintypes.com_error as exc:
                raise OfficeError(f"Unknown table style '{style}': {exc}") from exc
        return {
            "workbook": ws.Parent.Name,
            "sheet": ws.Name,
            "table": table.Name,
            "address": _address(target),
            "rows": int(target.Rows.Count),
        }

    return run("excel", op, create=False)


def excel_autofit(
    sheet: str | None = None,
    workbook: str | None = None,
) -> dict[str, Any]:
    """Auto-fit column widths and row heights on the used range of a worksheet."""

    def op(app: Any) -> dict[str, Any]:
        ws = _ws(app, workbook, sheet)
        used = ws.UsedRange
        try:
            used.Columns.AutoFit()
        except pywintypes.com_error:
            pass
        try:
            used.Rows.AutoFit()
        except pywintypes.com_error:
            pass
        return {"workbook": ws.Parent.Name, "sheet": ws.Name, "address": _address(used)}

    return run("excel", op, create=False)


def excel_recalculate(workbook: str | None = None) -> dict[str, Any]:
    """Force a full recalculation of all formulas in the workbook."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        app.CalculateFullRebuild()
        return {"workbook": book.Name, "recalculated": True}

    return run("excel", op, create=False)


def excel_save(workbook: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Save a workbook in place or to `path`; unsaved workbooks get an automatic .xlsx path."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        target = prepare_output(path) if path else None
        if target:
            book.SaveAs(target)
        elif not book.Path:
            target = default_output("excel", "xlsx")
            book.SaveAs(target)
        else:
            book.Save()
        return {
            "workbook": book.Name,
            "path": book.FullName,
            "saved": bool(book.Saved),
            "file": file_info(book.FullName),
        }

    return run("excel", op, create=False)


def excel_export_pdf(workbook: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Export the active sheet of a workbook to PDF (saving the workbook first if unsaved)."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        if not book.Path:
            book.SaveAs(default_output("excel", "xlsx"))
        target = prepare_output(path) if path else default_output("excel", "pdf")
        book.ExportAsFixedFormat(0, target)  # xlTypePDF
        return {
            "workbook": book.Name,
            "pdf": target,
            "file": file_info(target),
        }

    return run("excel", op, create=False)


def excel_close(workbook: str, save: bool = False) -> dict[str, Any]:
    """Close a workbook; `save=True` saves it first (unsaved files get an automatic path)."""

    def op(app: Any) -> dict[str, Any]:
        book = _wb(app, workbook)
        name = book.Name
        if save and not book.Path:
            book.SaveAs(default_output("excel", "xlsx"))
        book.Close(bool(save))
        return {"closed": name, "workbooks_left": int(app.Workbooks.Count)}

    return run("excel", op, create=False)
