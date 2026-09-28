"""Write the step -> number mapping to a new Excel file."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string

from .pdf_numbers import StepResult

HEADER_FILL = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
NOT_FOUND_FILL = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
_THIN = Side(style="thin", color="000000")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
CENTER = Alignment(horizontal="center", vertical="center")

DEFAULT_STEP_HEADERS = ("Step", "Step", "Step OPC UA Number")
DEFAULT_TRANSITION_HEADERS = ("Transition", "Transition", "Transition OPC UA Number")
OTHER_HEADERS = ("Other", "Other", "Other OPC UA Number")


def _cell_value(numbers: list[str], separator: str):
    if len(numbers) == 1 and numbers[0].isdigit() and not numbers[0].startswith("0"):
        return int(numbers[0])
    return f" {separator} ".join(numbers) if numbers else None


def step_digits(step: str) -> int | None:
    """The number inside a step name: "T10" -> 10, "S02" -> 2, "T00" -> 0."""
    m = re.search(r"\d+", step)
    return int(m.group()) if m else None


def natural_key(step: str):
    """Alphabetical order that compares the digits as numbers: S2 < S02b < S10."""
    return [int(part) if part.isdigit() else part.upper() for part in re.split(r"(\d+)", step)]


@dataclass
class Sheet:
    name: str
    results: list[StepResult] = field(default_factory=list)
    error: str | None = None  # shown on the sheet when the sequence could not be read


def write_mapping(output_path: str, results: list[StepResult], *, sheet_name: str = "Mapping",
                  checks: list[str] | None = None, **layout) -> None:
    """Create `output_path` with a single sheet (see write_workbook for `layout`)."""
    write_workbook(output_path, [Sheet(sheet_name, results)], checks=checks, **layout)


def _blocks(results: list[StepResult], group: bool, step_headers, transition_headers):
    """Split the steps into tables: steps (S...) first, then transitions (T...), then any
    other prefix. Without `group` everything is one table with the step headers."""
    if not group:
        return [(step_headers, results)]
    steps, transitions, other = [], [], []
    for r in results:
        prefix = re.match(r"[A-Za-z]*", r.step).group().upper()
        (steps if prefix == "S" else transitions if prefix == "T" else other).append(r)
    blocks = [(step_headers, steps), (transition_headers, transitions), (OTHER_HEADERS, other)]
    return [(h, rs) for h, rs in blocks if rs] or [(step_headers, [])]


def write_workbook(
    output_path: str,
    sheets: list[Sheet],
    *,
    checks: list[str] | None = None,
    step_column: str = "A",
    step_number_column: str | None = "B",
    number_column: str = "C",
    pages_column: str | None = None,
    start_row: int = 1,
    step_header: str | None = DEFAULT_STEP_HEADERS[0],
    step_number_header: str | None = DEFAULT_STEP_HEADERS[1],
    number_header: str | None = DEFAULT_STEP_HEADERS[2],
    transition_header: str | None = DEFAULT_TRANSITION_HEADERS[0],
    transition_number_header: str | None = DEFAULT_TRANSITION_HEADERS[1],
    transition_opc_header: str | None = DEFAULT_TRANSITION_HEADERS[2],
    pages_header: str = "PDF 2 pages",
    separator: str = "OR",
    not_found_text: str = "",
    sort: bool = True,
    group: bool = True,
) -> None:
    """Create `output_path` with one sheet per sequence.

    With `group` (default) each sheet holds a table of steps (S...) and, below it after
    a blank row, a table of transitions (T...), each with its own yellow header row.
    `step_number_column` receives the digits of the name (T10 -> 10). The first header
    row is `start_row`. Steps without a number get a grey cell.
    `checks` (points for a human to verify) are listed on a separate "Check" sheet.
    """
    wb = Workbook()
    wb.remove(wb.active)

    step_col = column_index_from_string(step_column.upper())
    step_no_col = column_index_from_string(step_number_column.upper()) if step_number_column else None
    number_col = column_index_from_string(number_column.upper())
    pages_col = column_index_from_string(pages_column.upper()) if pages_column else None
    columns = [c for c in (step_col, step_no_col, number_col, pages_col) if c]
    with_headers = step_header is not None or number_header is not None
    step_headers = (step_header, step_number_header, number_header)
    transition_headers = (transition_header, transition_number_header, transition_opc_header)

    def styled(ws, row, col, value=None, fill=None, bold=False):
        cell = ws.cell(row=row, column=col, value=value)
        cell.border, cell.alignment = BORDER, CENTER
        if fill:
            cell.fill = fill
        if bold:
            cell.font = Font(bold=True)
        return cell

    for sheet in sheets:
        ws = wb.create_sheet(sheet.name[:31])
        row = start_row

        if sheet.error:
            cell = ws.cell(row=row, column=step_col, value=f"ERROR: {sheet.error}")
            cell.font = Font(bold=True, color="FF0000")
            ws.sheet_properties.tabColor = "FF0000"
            continue

        results = sorted(sheet.results, key=lambda r: natural_key(r.step)) if sort else sheet.results
        for i, (headers, block) in enumerate(_blocks(results, group, step_headers,
                                                      transition_headers)):
            if i:
                row += 1  # blank row between the tables
            if with_headers:
                texts = {step_col: headers[0], step_no_col: headers[1], number_col: headers[2],
                         pages_col: pages_header}
                for col in columns:
                    styled(ws, row, col, texts[col], HEADER_FILL, bold=True)
                row += 1
            for result in block:
                styled(ws, row, step_col, result.step)
                if step_no_col:
                    styled(ws, row, step_no_col, step_digits(result.step))
                value = _cell_value(result.numbers, separator)
                if value is None:
                    styled(ws, row, number_col, not_found_text or None, NOT_FOUND_FILL)
                else:
                    styled(ws, row, number_col, value)
                if pages_col:
                    styled(ws, row, pages_col,
                           ", ".join(map(str, result.pages)) if result.pages else None)
                row += 1

        widths = {step_col: 14, step_no_col: 12, number_col: 30, pages_col: 15}
        for col in columns:
            ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = widths[col]

    if checks:
        check = wb.create_sheet("Check")
        check.cell(row=1, column=1, value="Please verify").font = Font(bold=True)
        for i, text in enumerate(checks, start=2):
            check.cell(row=i, column=1, value=text)
        check.column_dimensions["A"].width = 110
        check.sheet_properties.tabColor = "FF0000"

    wb.save(output_path)
