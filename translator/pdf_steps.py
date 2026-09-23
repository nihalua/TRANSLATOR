"""PDF 1: read the step names from the first column of the table in a given section."""

from __future__ import annotations

import re
from dataclasses import dataclass

import pdfplumber
from pypdf import PdfReader

# Default filter for first-column cells: letters followed by digits, e.g. T00, S01, AB12.
# Header cells such as "Step" or "Name" do not match and are skipped.
DEFAULT_STEP_PATTERN = r"^[A-Za-z]{1,10}\d+[A-Za-z0-9]*$"

# A table-of-contents line: dotted leaders and/or a trailing page number.
_TOC_LINE = re.compile(r"(\.{3,}|…|\s\.\s\.)|\s\d+\s*$")
# Any numbered heading, e.g. "3 Results", "3.2 Steps", "4.1.2. Details".
_ANY_HEADING = re.compile(r"^\s*\d+(\.\d+)*\.?\s+\S")


class SectionNotFoundError(Exception):
    pass


class TableNotFoundError(Exception):
    pass


@dataclass
class _Heading:
    page_index: int
    top: float  # y position of the heading on the page (0 = top of page)


def _section_regex(section: str) -> re.Pattern:
    # "3.1" matches "3.1 Title" / "3.1. Title" but not "3.10 ..." or "3.1.2 ...".
    sec = re.escape(section.strip().rstrip("."))
    return re.compile(rf"^\s*{sec}\.?(?![\d.]*\d)(\s+|$)")


def _find_in_outline(pdf_path: str, section: str) -> int | None:
    """Return the page index of the section using the PDF bookmarks, if there are any."""
    try:
        reader = PdfReader(pdf_path)
        outline = reader.outline
    except Exception:
        return None
    pattern = _section_regex(section)

    def walk(items):
        for item in items:
            if isinstance(item, list):
                found = walk(item)
                if found is not None:
                    return found
                continue
            title = getattr(item, "title", "") or ""
            if pattern.match(title):
                try:
                    return reader.get_destination_page_number(item)
                except Exception:
                    return None
        return None

    return walk(outline or [])


def _heading_on_page(page, pattern: re.Pattern) -> float | None:
    for line in page.extract_text_lines():
        text = line["text"]
        if pattern.match(text) and not _TOC_LINE.search(text):
            return line["top"]
    return None


def _locate_section(pdf, pdf_path: str, section: str) -> _Heading:
    pattern = _section_regex(section)

    outline_page = _find_in_outline(pdf_path, section)
    if outline_page is not None:
        top = _heading_on_page(pdf.pages[outline_page], pattern)
        return _Heading(outline_page, top if top is not None else 0.0)

    # No bookmarks: scan the text for the heading, ignoring table-of-contents entries.
    for index, page in enumerate(pdf.pages):
        top = _heading_on_page(page, pattern)
        if top is not None:
            return _Heading(index, top)
    raise SectionNotFoundError(f"Section {section!r} was not found in {pdf_path}")


def _tables(page, strategy: str):
    if strategy == "text":
        settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}
    else:
        settings = {}
    return sorted(page.find_tables(table_settings=settings), key=lambda t: t.bbox[1])


def _first_heading_top(page) -> float | None:
    for line in page.extract_text_lines():
        if _ANY_HEADING.match(line["text"]) and not _TOC_LINE.search(line["text"]):
            return line["top"]
    return None


def _collect_rows(pdf, heading: _Heading, strategy: str) -> list[list]:
    page = pdf.pages[heading.page_index]
    candidates = [t for t in _tables(page, strategy) if t.bbox[1] >= heading.top]
    # The table may start on the following page when the heading sits at the bottom.
    page_index = heading.page_index
    if not candidates and page_index + 1 < len(pdf.pages):
        page_index += 1
        page = pdf.pages[page_index]
        candidates = _tables(page, strategy)
        next_heading = _first_heading_top(page)
        if next_heading is not None:
            candidates = [t for t in candidates if t.bbox[1] < next_heading]
    if not candidates:
        raise TableNotFoundError(
            f"No table found after the section heading on page {heading.page_index + 1}"
        )

    table = candidates[0]
    rows = table.extract()
    n_cols = len(rows[0]) if rows else 0

    # Follow the table onto the next pages while it continues (same number of columns
    # and no new section heading above it).
    is_last_on_page = table is candidates[-1]
    while is_last_on_page and page_index + 1 < len(pdf.pages):
        next_page = pdf.pages[page_index + 1]
        next_tables = _tables(next_page, strategy)
        if not next_tables:
            break
        first = next_tables[0]
        heading_top = _first_heading_top(next_page)
        if heading_top is not None and heading_top < first.bbox[1]:
            break
        cont = first.extract()
        if not cont or len(cont[0]) != n_cols:
            break
        rows.extend(cont)
        page_index += 1
        is_last_on_page = len(next_tables) == 1

    return rows


def extract_step_names(
    pdf_path: str,
    section: str,
    step_pattern: str = DEFAULT_STEP_PATTERN,
    table_strategy: str = "lines",
) -> list[str]:
    """Return the step names from the first column of the table in `section`.

    Cells in the first column that do not match `step_pattern` (headers, blanks)
    are skipped. Duplicates are removed while keeping the original order.
    """
    regex = re.compile(step_pattern)
    with pdfplumber.open(pdf_path) as pdf:
        heading = _locate_section(pdf, pdf_path, section)
        rows = _collect_rows(pdf, heading, table_strategy)

    steps: list[str] = []
    seen: set[str] = set()
    for row in rows:
        if not row:
            continue
        cell = (row[0] or "").strip()
        # A cell may wrap onto several lines; the step name is the first line.
        cell = cell.splitlines()[0].strip() if cell else ""
        if cell and regex.match(cell) and cell not in seen:
            seen.add(cell)
            steps.append(cell)
    return steps
