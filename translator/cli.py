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
    p.add_argument("--output", "-o", help="Excel file to create, e.g. mapping.xlsx")

    g = p.add_argument_group("PDF options")
    g.add_argument("--step-pattern", default=DEFAULT_STEP_PATTERN,
                   help="regex a first-column cell must match to count as a step name "
                        "(default: %(default)s)")
    g.add_argument("--table-strategy", choices=["auto", "lines", "text"], default="auto",
                   help="how tables in PDF 1 are detected: ruling 'lines', 'text' alignment, "
                        "or 'auto' (try both) (default: %(default)s)")
    g.add_argument("--number-label", default=DEFAULT_NUMBER_LABEL,
                   help="label of the number field in PDF 2 (default: %(default)s)")
    g.add_argument("--name-label", default=None,
                   help='label of the step name field in PDF 2, e.g. "name:" (optional, '
                        "makes matching stricter)")

    x = p.add_argument_group("Excel options")
    x.add_argument("--sheet", default="Mapping", help="sheet name (default: %(default)s)")
    x.add_argument("--step-column", default="A", help="column for step names (default: A)")
    x.add_argument("--number-column", default="B", help="column for numbers (default: B)")
    x.add_argument("--pages-column", default=None,
                   help="optional column listing the PDF 2 pages each number came from")
    x.add_argument("--start-row", type=int, default=1,
                   help="row of the header (data starts below it) (default: 1)")
    x.add_argument("--step-header", default="Step", help="header text for the step column")
    x.add_argument("--number-header", default="Number", help="header text for the number column")
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
    steps = extract_step_names(args.pdf1, args.section, args.step_pattern, args.table_strategy,
                               pages=pages, log=log)
    log(f"  {len(steps)} step(s) found: {', '.join(steps)}")

    log(f"Searching {args.pdf2} for '{args.number_label}' values ...")
    by_step = find_step_numbers(args.pdf2, steps, args.number_label, args.name_label)
    results = [by_step[s] for s in steps]

    missing = [r.step for r in results if not r.numbers]
    for r in results:
        if r.numbers:
            variants = ", ".join(dict.fromkeys(m.variant for m in r.matches))
            log(f"  {r.step}: {f' {args.separator} '.join(r.numbers)}  (as {variants})")
    if missing:
        log(f"  WARNING: no number found for {len(missing)} step(s): {', '.join(missing)}")

    write_mapping(
        args.output,
        results,
        sheet_name=args.sheet,
        step_column=args.step_column,
        number_column=args.number_column,
        pages_column=args.pages_column,
        start_row=args.start_row,
        step_header=None if args.no_header else args.step_header,
        number_header=None if args.no_header else args.number_header,
        separator=args.separator,
        not_found_text=args.not_found,
    )
    log(f"Excel file written: {args.output}")
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
