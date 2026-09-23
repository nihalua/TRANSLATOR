"""Write the step -> number mapping to a new Excel file."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string

from .pdf_numbers import StepResult

NOT_FOUND_FILL = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")


def _cell_value(numbers: list[str], separator: str):
    if len(numbers) == 1 and numbers[0].isdigit() and not numbers[0].startswith("0"):
        return int(numbers[0])
    return f" {separator} ".join(numbers) if numbers else None


def write_mapping(
    output_path: str,
    results: list[StepResult],
    *,
    sheet_name: str = "Mapping",
    step_column: str = "A",
    number_column: str = "B",
    pages_column: str | None = None,
    start_row: int = 1,
    step_header: str | None = "Step",
    number_header: str | None = "Number",
    pages_header: str = "PDF 2 pages",
    separator: str = "OR",
    not_found_text: str = "",
) -> None:
    """Create `output_path` with one row per step.

    Headers are written on `start_row` (unless both headers are None) and the data
    starts on the row below. Steps without a number are highlighted in yellow.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    step_col = column_index_from_string(step_column.upper())
    number_col = column_index_from_string(number_column.upper())
    pages_col = column_index_from_string(pages_column.upper()) if pages_column else None

    row = start_row
    if step_header is not None or number_header is not None:
        headers = [(step_col, step_header), (number_col, number_header)]
        if pages_col:
            headers.append((pages_col, pages_header))
        for col, text in headers:
            if text is not None:
                ws.cell(row=row, column=col, value=text).font = Font(bold=True)
        row += 1

    for result in results:
        ws.cell(row=row, column=step_col, value=result.step)
        value = _cell_value(result.numbers, separator)
        number_cell = ws.cell(row=row, column=number_col, value=value)
        if value is None:
            if not_found_text:
                number_cell.value = not_found_text
            number_cell.fill = NOT_FOUND_FILL
        if pages_col and result.pages:
            ws.cell(row=row, column=pages_col, value=", ".join(map(str, result.pages)))
        row += 1

    for col, width in ((step_col, 15), (number_col, 40), (pages_col, 15)):
        if col:
            ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width

    wb.save(output_path)
