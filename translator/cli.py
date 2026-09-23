"""Command line entry point.

Example:
    python -m translator --pdf1 spec.pdf --section 3.1 --pdf2 steps.pdf --output mapping.xlsx
"""

from __future__ import annotations

import argparse
import sys

from .excel_writer import write_mapping
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


def run(args: argparse.Namespace, log=print) -> int:
    pages = parse_page_range(args.pages) if args.pages else None
    where = " and ".join(filter(None, [
        f"section {args.section}" if args.section else "",
        f"pages {args.pages}" if args.pages else "",
    ]))
    log(f"Reading step names from {where} of {args.pdf1} ...")
    warnings: list[str] = []
    steps = extract_step_names(args.pdf1, args.section, args.step_pattern, args.table_strategy,
                               pages=pages, log=log, warnings=warnings)
    log(f"  {len(steps)} step(s) found: {', '.join(steps)}")
    for w in warnings:
        log(f"  CHECK: {w}")

    pdf2_pages = parse_page_range(args.pdf2_pages) if args.pdf2_pages else None
    in_pages = f" (pages {args.pdf2_pages})" if args.pdf2_pages else ""
    log(f"Searching {args.pdf2}{in_pages} for '{args.number_label}' values ...")
    by_step = find_step_numbers(args.pdf2, steps, args.number_label, args.name_label,
                                pages=pdf2_pages)
    results = [by_step[s] for s in steps]

    missing = [r.step for r in results if not r.numbers]
    for r in results:
        if r.numbers:
            variants = ", ".join(dict.fromkeys(m.variant for m in r.matches))
            log(f"  {r.step}: {f' {args.separator} '.join(r.numbers)}  (as {variants})")
    if missing:
        log(f"  WARNING: no number found for {len(missing)} step(s): {', '.join(missing)}")
        warnings.append(f"No number found in PDF 2 for: {', '.join(missing)}")

    write_mapping(
        args.output,
        results,
        sheet_name=args.sheet,
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
        checks=warnings,
    )
    log(f"Excel file written: {args.output}")
    if warnings:
        log(f"{len(warnings)} point(s) to check are listed above and on the 'Check' sheet.")
    else:
        log("Nothing to check: every text in the step-name column was taken as a step name.")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.gui or not any([args.pdf1, args.pdf2, args.section, args.pages, args.output]):
        from .gui import launch
        return launch(parser)

    missing = [n for n in ("pdf1", "pdf2", "output") if not getattr(args, n)]
    if not args.section and not args.pages:
        missing.append("section (or --pages)")
    if missing:
        parser.error("missing required option(s): " + ", ".join("--" + m for m in missing))
    try:
        return run(args)
    except Exception as exc:  # show a readable message instead of a traceback
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
