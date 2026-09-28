"""Read the Excel file to check and write the check report."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string

from .pdf_check import ERROR, INFO, WARNING, SequenceReport


@dataclass
class ExcelSheet:
    name: str
    rows: list[tuple[str, object, list[str]]] = field(default_factory=list)  # step, step no., numbers
    error: str | None = None


def split_numbers(value) -> list[str]:
    if value is None or str(value).strip() == "":
        return []
    return [p.strip() for p in re.split(r"\s+OR\s+", str(value).strip(), flags=re.I) if p.strip()]


def read_sheet(path: str, sheet: str, step_col="A", step_no_col="B", number_col="C",
               header_row=1) -> ExcelSheet:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            return ExcelSheet(sheet, error=f"sheet '{sheet}' is missing in the Excel file")
        ws = wb[sheet]
        sc = column_index_from_string(step_col)
        bc = column_index_from_string(step_no_col) if step_no_col else None
        nc = column_index_from_string(number_col)
        out = ExcelSheet(sheet)
        for row in ws.iter_rows(min_row=header_row, values_only=True):
            def cell(c):
                return row[c - 1] if c and c - 1 < len(row) else None
            step = cell(sc)
            if step is None or str(step).strip() == "":
                continue  # blank row between the steps and transitions tables
            step = str(step).strip()
            if step.upper().startswith("ERROR"):
                out.error = step
                continue
            if not re.search(r"\d", step):
                continue  # a table header ("Step", "Transition", ...); a real step name
                # without digits would still be reported as missing by the PDF 1 check
            out.rows.append((step, cell(bc), split_numbers(cell(nc))))
        return out
    finally:
        wb.close()


_FILL = {
    ERROR: PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE"),
    WARNING: PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C"),
    INFO: None,
    "OK": PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE"),
}


def write_report(path: str, reports: list[SequenceReport], sources: dict[str, str]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    bold = Font(bold=True)
    r = 1
    for key, value in sources.items():
        ws.cell(row=r, column=1, value=key).font = bold
        ws.cell(row=r, column=2, value=value)
        r += 1
    r += 1
    headers = ["Sequence", "Steps in Excel", "Steps confirmed in PDF 1", "Numbers in Excel",
               "Numbers confirmed in PDF 2", "Errors", "Warnings", "Result"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=r, column=c, value=h).font = bold
    for rep in reports:
        r += 1
        ok = rep.errors == 0 and rep.warnings == 0
        result = "OK" if ok else ("ERRORS" if rep.errors else "CHECK WARNINGS")
        values = [rep.sequence, rep.excel_steps, rep.steps_confirmed, rep.numbers_claimed,
                  rep.numbers_confirmed, rep.errors, rep.warnings, result]
        for c, v in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=8).fill = _FILL["OK" if ok else (ERROR if rep.errors else WARNING)]
    for c, width in enumerate([14, 15, 24, 17, 26, 8, 10, 16], start=1):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = width
    ws.column_dimensions["B"].width = max(15, max((len(str(v)) for v in sources.values()),
                                                  default=15))

    issues = wb.create_sheet("Issues")
    for c, h in enumerate(["Severity", "Sequence", "Step", "Page", "Message"], start=1):
        issues.cell(row=1, column=c, value=h).font = bold
    order = {ERROR: 0, WARNING: 1, INFO: 2}
    all_issues = sorted((i for rep in reports for i in rep.issues),
                        key=lambda i: (order[i.severity], i.sequence))
    for r, issue in enumerate(all_issues, start=2):
        for c, v in enumerate([issue.severity, issue.sequence, issue.step, issue.page,
                               issue.message], start=1):
            cell = issues.cell(row=r, column=c, value=v)
            if _FILL[issue.severity]:
                cell.fill = _FILL[issue.severity]
    for col, width in zip("ABCDE", [10, 14, 10, 10, 120]):
        issues.column_dimensions[col].width = width
    if not all_issues:
        issues.cell(row=2, column=5, value="No issues found.")
    wb.save(path)
