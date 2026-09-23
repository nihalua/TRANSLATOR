"""PDF 1: read the step names from the first column of the table in a given section."""

from __future__ import annotations

import re
from dataclasses import dataclass

import pdfplumber
from pypdf import PdfReader

# Default filter for first-column cells: letters followed by digits, e.g. T00, S01, AB12.
# Header cells such as "Step" or "Name" do not match and are skipped.
DEFAULT_STEP_PATTERN = r"^[A-Za-z]{1,10}\d+[A-Za-z0-9]*$"

# When no page range is given, follow a section for at most this many pages.
MAX_SECTION_PAGES = 50

# A table-of-contents line: dotted leaders and/or a trailing page number.
_TOC_LINE = re.compile(r"(\.{3,}|…|\s\.\s\.)|\s\d+\s*$")
# Any numbered heading, e.g. "3 Results", "3.2 Steps", "4.1.2. Details".
_ANY_HEADING = re.compile(r"^\s*(?P<num>\d+(?:\.\d+)*)\.?\s+[A-Za-z]")
# Characters trimmed around the step name in a cell, e.g. "T00:" or "(S01)".
_TRIM = " \t:;,.()[]*"


class SectionNotFoundError(Exception):
    pass


class StepsNotFoundError(Exception):
    pass


@dataclass
class _Region:
    page_index: int
    top: float
    bottom: float


def parse_page_range(text: str) -> list[int]:
    """"3-7" -> [3, 4, 5, 6, 7]; "3,5,8-9" is also accepted. Pages are 1-based."""
    pages: list[int] = []
    for part in re.split(r"[,;\s]+", text.strip()):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)(?:\s*[-–]\s*(\d+))?", part)
        if not m:
            raise ValueError(f"Invalid page range: {text!r} (use e.g. 3-7)")
        start = int(m.group(1))
        end = int(m.group(2) or start)
        if start < 1 or end < start:
            raise ValueError(f"Invalid page range: {text!r}")
        pages.extend(p for p in range(start, end + 1) if p not in pages)
    if not pages:
        raise ValueError("Empty page range")
    return pages


def _section_regex(section: str) -> re.Pattern:
    # "3.1" matches "3.1 Title" / "3.1. Title" / "3.1Title" but not "3.10 ..." or "3.1.2 ...".
    sec = re.escape(section.strip().rstrip("."))
    return re.compile(rf"^\s*{sec}\.?(?![\d.]*\d)(\s+|$|(?=[A-Za-z]))")


def _is_subsection(num: str, section: str) -> bool:
    section = section.strip().rstrip(".")
    return num == section or num.startswith(section + ".")


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


class _Page:
    """Text lines and tables of one page, computed once."""

    # Lines this close to the top/bottom edge are page headers/footers, not section headings.
    MARGIN = 0.07

    def __init__(self, page):
        self.page = page
        self.lines = page.extract_text_lines()
        self._tables: dict[str, list] = {}
        self._words = None

    def words(self) -> list[dict]:
        if self._words is None:
            self._words = self.page.extract_words()
        return self._words

    def in_margin(self, top: float) -> bool:
        height = float(self.page.height)
        return not (height * self.MARGIN < top < height * (1 - self.MARGIN))

    def tables(self, strategy: str) -> list:
        if strategy not in self._tables:
            settings = ({"vertical_strategy": "text", "horizontal_strategy": "text"}
                        if strategy == "text" else {})
            try:
                found = self.page.find_tables(table_settings=settings)
            except Exception:
                found = []
            self._tables[strategy] = sorted(found, key=lambda t: t.bbox[1])
        return self._tables[strategy]

    def _in_table(self, line) -> bool:
        return any(t.bbox[1] - 1 <= line["top"] <= t.bbox[3] + 1 for t in self.tables("lines"))

    def headings(self):
        """(top, number, text) of numbered headings outside tables, headers/footers and TOC."""
        out = []
        for line in self.lines:
            if self.in_margin(line["top"]):
                continue
            m = _ANY_HEADING.match(line["text"])
            if m and not _TOC_LINE.search(line["text"]) and not self._in_table(line):
                out.append((line["top"], m.group("num"), line["text"]))
        return out

    def section_top(self, pattern: re.Pattern) -> float | None:
        for line in self.lines:
            if pattern.match(line["text"]) and not _TOC_LINE.search(line["text"]):
                return line["top"]
        return None


def _regions(pdf, pdf_path, section, pages, log):
    cache: dict[int, _Page] = {}

    def get(i):
        if i not in cache:
            cache[i] = _Page(pdf.pages[i])
        return cache[i]

    n = len(pdf.pages)
    if pages:
        bad = [p for p in pages if p > n]
        if bad:
            raise ValueError(f"PDF 1 has only {n} pages (asked for page {bad[0]})")
        indices = [p - 1 for p in pages]
    else:
        indices = None

    start_index, start_top = None, 0.0
    if section:
        pattern = _section_regex(section)
        if indices is not None:
            for i in indices:
                top = get(i).section_top(pattern)
                if top is not None:
                    start_index, start_top = i, top
                    break
            if start_index is None:
                log(f"  Note: heading {section} not found on the given pages; using the whole pages.")
        else:
            outline_page = _find_in_outline(pdf_path, section)
            if outline_page is not None:
                start_index = outline_page
                start_top = get(outline_page).section_top(pattern) or 0.0
            else:
                for i in range(n):
                    top = get(i).section_top(pattern)
                    if top is not None:
                        start_index, start_top = i, top
                        break
            if start_index is None:
                raise SectionNotFoundError(
                    f"Section {section!r} was not found in PDF 1. "
                    "Tip: enter the page numbers of the table instead."
                )
            indices = list(range(start_index, min(n, start_index + MAX_SECTION_PAGES)))
        if start_index is not None:
            log(f"  Section {section} found on page {start_index + 1}.")
    if indices is None:
        raise ValueError("Give a section number and/or a page range for PDF 1")

    regions: list[_Region] = []
    started = start_index is None
    for i in indices:
        page = get(i)
        top, bottom = 0.0, float(page.page.height)
        if not started:
            if i != start_index:
                continue
            started, top = True, start_top
        # Stop at the next heading that is not part of the section.
        if section:
            stop = next((h for h in page.headings()
                         if h[0] > top + 1 and not _is_subsection(h[1], section)), None)
            if stop is not None:
                regions.append(_Region(i, top, stop[0]))
                break
        regions.append(_Region(i, top, bottom))
    return regions, get


def _step_from_cell(cell: str | None, regex: re.Pattern) -> str | None:
    if not cell:
        return None
    first_line = cell.strip().splitlines()[0] if cell.strip() else ""
    for token in (first_line.strip(_TRIM), (first_line.split() or [""])[0].strip(_TRIM)):
        if token and regex.match(token):
            return token
    return None


def extract_step_names(
    pdf_path: str,
    section: str | None = None,
    step_pattern: str = DEFAULT_STEP_PATTERN,
    table_strategy: str = "auto",
    pages: list[int] | None = None,
    log=lambda message: None,
) -> list[str]:
    """Return the step names from the first column of the table.

    The table is located by `section` (e.g. "3.1"), by `pages` (1-based page
    numbers), or both. Cells in the first column that do not match `step_pattern`
    (headers, blanks) are skipped. Duplicates are removed, order is kept.
    """
    regex = re.compile(step_pattern)
    if table_strategy == "auto":
        strategies = ["position", "lines", "text"]
    else:
        strategies = [table_strategy]

    with pdfplumber.open(pdf_path) as pdf:
        regions, get = _regions(pdf, pdf_path, section, pages, log)
        diagnostics: list[str] = []
        for strategy in strategies:
            if strategy == "position":
                steps = _steps_by_position(regions, get, regex, log)
            else:
                steps = _steps_from_tables(regions, get, strategy, regex, diagnostics)
            if steps:
                return steps

        # Last resort: first word of every text line in the region.
        steps = []
        for region in regions:
            for line in get(region.page_index).lines:
                if region.top - 1 <= line["top"] < region.bottom:
                    step = _step_from_cell(line["text"], regex)
                    if step and step not in steps:
                        steps.append(step)
        if steps:
            log("  No step names found in a table; read them from the text lines instead.")
            return steps

    where = ", ".join(str(r.page_index + 1) for r in regions) or "-"
    details = "\n".join(diagnostics[:12]) or "  (no table found on these pages)"
    raise StepsNotFoundError(
        f"No step names found in PDF 1 (looked at page(s) {where}).\n"
        f"First-column cells that were read:\n{details}\n"
        f"Step names must match the pattern {step_pattern!r}."
    )


def _steps_by_position(regions, get, regex, log) -> list[str]:
    """Step names are the words in the leftmost column that match the pattern.

    Works for "tables" that are really framed blocks (one block per step, the name
    in a box on the left), where no regular rows and columns can be detected.
    """
    found = []  # (page_index, top, x0, x1, name)
    for region in regions:
        page = get(region.page_index)
        for w in page.words():
            if page.in_margin(w["top"]) or not (region.top - 1 <= w["top"] < region.bottom):
                continue
            name = w["text"].strip(_TRIM)
            if name and regex.match(name):
                found.append((region.page_index, w["top"], w["x0"], w["x1"], name))
    if not found:
        return []

    # Group the words into vertical columns: a word belongs to a column when its
    # centre lies within the column's first (leftmost) word, or the other way round,
    # so centred names of different lengths still line up.
    columns: list[list] = []
    for word in sorted(found, key=lambda w: w[2]):
        centre = (word[2] + word[3]) / 2
        for col in columns:
            x0, x1 = col[0][2], col[0][3]
            if x0 <= centre <= x1 or word[2] <= (x0 + x1) / 2 <= word[3]:
                col.append(word)
                break
        else:
            columns.append([word])

    # Leftmost column holding at least two names (a stray word left of the table
    # is ignored); a single name only counts when nothing else was found.
    column = next((c for c in columns if len(c) >= 2), columns[0])
    steps: list[str] = []
    for _, _, _, _, name in sorted(column, key=lambda w: (w[0], w[1])):
        if name not in steps:
            steps.append(name)
    log(f"  Step names read from the left-hand column ({len(steps)} found).")
    return steps


def _steps_from_tables(regions, get, strategy, regex, diagnostics) -> list[str]:
    steps: list[str] = []
    for region in regions:
        page = get(region.page_index)
        for table in page.tables(strategy):
            for row_box, row in zip(table.rows, table.extract()):
                # Only rows inside the section (the table may start above the heading).
                top = row_box.bbox[1]
                if not row or page.in_margin(top) or not (region.top - 1 <= top < region.bottom):
                    continue
                step = _step_from_cell(row[0], regex)
                if step:
                    if step not in steps:
                        steps.append(step)
                elif row[0] and len(diagnostics) < 12:
                    diagnostics.append(f"  page {region.page_index + 1} ({strategy}): {row[0]!r}")
    return steps
