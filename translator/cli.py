"""Command line entry point.

Example:
    python -m translator --pdf1 spec.pdf --section 3.1 --pdf2 steps.pdf --output mapping.xlsx
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass

from .excel_writer import Sheet, write_workbook
from .pdf_numbers import DEFAULT_NUMBER_LABEL, find_step_numbers
from .pdf_steps import DEFAULT_STEP_PATTERN, extract_step_names, parse_page_range


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="translator",
        description="Map step names (table in PDF 1) to step numbers (PDF 2) in a new Excel file.",
    )
    p.add_argument("--gui", action="store_true", help="open the graphical interface")
    p.add_argument("--pdf1", help="PDF containing the table with the step names")
    p.add_argument("--section", help='section number of the table in PDF 1, e.g. "3.1"')
    p.add_argument("--pages", help='PDF page numbers of the table in PDF 1, e.g. "3-7" '
                                   "(as shown by the PDF viewer)")
    p.add_argument("--pdf2", help="PDF containing the step properties with the 'number:' field")
    p.add_argument("--pdf2-pages", help='PDF page numbers to search in PDF 2, e.g. "100-250" '
                                        "(default: all pages)")
    p.add_argument("--sequence", action="append", metavar="SECTION:PAGES1:PAGES2",
                   help='one sequence per option, e.g. --sequence "3.1:3-7:100-250" '
                        '--sequence "3.2:8-9:251-300"; each becomes a sheet '
                        '"Sequence 01", "Sequence 02", ... (parts may be left empty)')
    p.add_argument("--output", "-o", help="Excel file to create, e.g. mapping.xlsx")

    g = p.add_argument_group("PDF options")
    g.add_argument("--step-pattern", default=DEFAULT_STEP_PATTERN,
                   help="regex a first-column cell must match to count as a step name "
                        "(default: %(default)s)")
    g.add_argument("--table-strategy", choices=["auto", "position", "lines", "text"],
                   default="auto",
                   help="how the step names in PDF 1 are read: 'position' (leftmost column of "
                        "step-like words), table by ruling 'lines', table by 'text' alignment, "
                        "or 'auto' (try them in that order) (default: %(default)s)")
    g.add_argument("--number-label", default=DEFAULT_NUMBER_LABEL,
                   help="label of the number field in PDF 2 (default: %(default)s)")
    g.add_argument("--name-label", default=None,
                   help='label of the step name field in PDF 2, e.g. "name:" (optional, '
                        "makes matching stricter)")

    x = p.add_argument_group("Excel options")
    x.add_argument("--sheet", default="Mapping", help="sheet name (default: %(default)s)")
    x.add_argument("--step-column", default="A", help="column for step names (default: A)")
    x.add_argument("--step-number-column", default="B",
                   help='column for the number inside the step name, T10 -> 10 (default: B; '
                        '"" to leave it out)')
    x.add_argument("--number-column", default="C", help="column for the numbers from PDF 2 "
                                                        "(default: C)")
    x.add_argument("--pages-column", default=None,
                   help="optional column listing the PDF 2 pages each number came from")
    x.add_argument("--start-row", type=int, default=1,
                   help="row of the header (data starts below it) (default: 1)")
    x.add_argument("--step-header", default="Step", help="header text for the step column")
    x.add_argument("--step-number-header", default="Step no.",
                   help="header text for the step number column")
    x.add_argument("--number-header", default="Number", help="header text for the number column")
    x.add_argument("--no-sort", action="store_true",
                   help="keep the order of PDF 1 instead of sorting the steps alphabetically")
    x.add_argument("--no-header", action="store_true", help="do not write a header row")
    x.add_argument("--separator", default="OR",
                   help="word placed between multiple numbers (default: %(default)s)")
    x.add_argument("--not-found", default="",
                   help="text written when no number is found (cell is highlighted anyway)")
    return p


@dataclass
class Job:
    """One sequence: one table in PDF 1 -> one sheet in the Excel file."""
    sheet: str
    section: str | None = None
    pages: str | None = None       # PDF 1 pages, e.g. "3-7"
    pdf2_pages: str | None = None  # PDF 2 pages, e.g. "100-250"


def parse_sequence(text: str, index: int) -> Job:
    """--sequence "3.1:3-7:100-250" (section:PDF 1 pages:PDF 2 pages; parts may be empty)."""
    parts = [p.strip() or None for p in (text.split(":") + ["", ""])[:3]]
    if not parts[0] and not parts[1]:
        raise ValueError(f"--sequence {text!r}: give a section and/or PDF 1 pages")
    return Job(f"Sequence {index:02d}", *parts)


def _run_job(job: Job, args, log, checks: list[str]) -> Sheet:
    where = " and ".join(filter(None, [
        f"section {job.section}" if job.section else "",
        f"pages {job.pages}" if job.pages else "",
    ]))
    log(f"[{job.sheet}] Reading step names from {where} of PDF 1 ...")
    warnings: list[str] = []
    steps = extract_step_names(
        args.pdf1, job.section, args.step_pattern, args.table_strategy,
        pages=parse_page_range(job.pages) if job.pages else None, log=log, warnings=warnings,
    )
    log(f"  {len(steps)} step(s) found: {', '.join(steps)}")

    in_pages = f" (pages {job.pdf2_pages})" if job.pdf2_pages else ""
    log(f"[{job.sheet}] Searching PDF 2{in_pages} for '{args.number_label}' values ...")
    by_step = find_step_numbers(
        args.pdf2, steps, args.number_label, args.name_label,
        pages=parse_page_range(job.pdf2_pages) if job.pdf2_pages else None,
    )
    results = [by_step[s] for s in steps]
    for r in results:
        if r.numbers:
            variants = ", ".join(dict.fromkeys(m.variant for m in r.matches))
            log(f"  {r.step}: {f' {args.separator} '.join(r.numbers)}  (as {variants})")
    missing = [r.step for r in results if not r.numbers]
    if missing:
        log(f"  WARNING: no number found for {len(missing)} step(s): {', '.join(missing)}")
        warnings.append(f"No number found in PDF 2 for: {', '.join(missing)}")

    for w in warnings:
        log(f"  CHECK: {w}")
    checks.extend(f"{job.sheet}: {w}" for w in warnings)
    return Sheet(job.sheet, results)


def run(args: argparse.Namespace, log=print, jobs: list[Job] | None = None) -> int:
    """Process every job (sequence) and write one Excel file with a sheet per job.

    Returns the number of sequences that failed (0 = all fine)."""
    if jobs is None:
        if args.sequence:
            jobs = [parse_sequence(text, i) for i, text in enumerate(args.sequence, start=1)]
        else:
            jobs = [Job(args.sheet, args.section, args.pages, args.pdf2_pages)]
    if not jobs:
        raise ValueError("Nothing to do: no sequence was filled in")

    checks: list[str] = []
    sheets: list[Sheet] = []
    failed = 0
    for job in jobs:
        try:
            sheets.append(_run_job(job, args, log, checks))
        except Exception as exc:
            if len(jobs) == 1:
                raise
            failed += 1
            log(f"[{job.sheet}] ERROR: {exc}")
            checks.append(f"{job.sheet}: ERROR - {exc}")
            sheets.append(Sheet(job.sheet, error=str(exc)))

    write_workbook(
        args.output,
        sheets,
        checks=checks,
        step_column=args.step_column,
        step_number_column=args.step_number_column or None,
        number_column=args.number_column,
        pages_column=args.pages_column,
        start_row=args.start_row,
        step_header=None if args.no_header else args.step_header,
        step_number_header=None if args.no_header else args.step_number_header,
        number_header=None if args.no_header else args.number_header,
        separator=args.separator,
        not_found_text=args.not_found,
        sort=not args.no_sort,
    )
    log(f"Excel file written: {args.output} ({len(sheets)} sheet(s))")
    if failed:
        log(f"{failed} sequence(s) FAILED - see the red sheet tabs and the 'Check' sheet.")
    if checks:
        log(f"{len(checks)} point(s) to check are listed above and on the 'Check' sheet.")
    else:
        log("Nothing to check: every text in the step-name column was taken as a step name.")
    return failed


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.gui or not any([args.pdf1, args.pdf2, args.section, args.pages, args.sequence,
                            args.output]):
        from .gui import launch
        return launch(parser)

    missing = [n for n in ("pdf1", "pdf2", "output") if not getattr(args, n)]
    if not args.section and not args.pages and not args.sequence:
        missing.append("section (or --pages or --sequence)")
    if missing:
        parser.error("missing required option(s): " + ", ".join("--" + m for m in missing))
    try:
        return run(args)
    except Exception as exc:  # show a readable message instead of a traceback
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
