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

    # Lines this close to the top/bottom edge are never taken as section headings
    # (page numbers such as "3 of 20" look like headings). Step names are NOT
    # filtered by this margin; see is_repeated().
    MARGIN = 0.07
    # How far (in points) text may move between pages and still be the same header.
    REPEAT_TOLERANCE = 6
    # Pages before/after compared to recognise headers and footers.
    NEIGHBOURS = 2

    def __init__(self, page, index: int, get, n_pages: int):
        self.page = page
        self.index = index
        self._get = get
        self._n = n_pages
        self.lines = page.extract_text_lines()
        self._tables: dict[str, list] = {}
        self._words = None
        self._tops: dict[str, list[float]] | None = None

    def words(self) -> list[dict]:
        if self._words is None:
            self._words = self.page.extract_words()
        return self._words

    def _word_tops(self) -> dict[str, list[float]]:
        if self._tops is None:
            self._tops = {}
            for w in self.words():
                self._tops.setdefault(w["text"].strip(_TRIM), []).append(w["top"])
        return self._tops

    def is_repeated(self, text: str, top: float) -> bool:
        """True for page headers/footers: the same text at the same height on a
        neighbouring page. A step name is unique, so it is never removed this way."""
        for j in range(self.index - self.NEIGHBOURS, self.index + self.NEIGHBOURS + 1):
            if j == self.index or not 0 <= j < self._n:
                continue
            tops = self._get(j)._word_tops().get(text, [])
            if any(abs(t - top) <= self.REPEAT_TOLERANCE for t in tops):
                return True
        return False

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
    n = len(pdf.pages)

    def get(i):
        if i not in cache:
            cache[i] = _Page(pdf.pages[i], i, get, n)
        return cache[i]

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
    warnings: list[str] | None = None,
) -> list[str]:
    """Return the step names from the first column of the table.

    The table is located by `section` (e.g. "3.1"), by `pages` (1-based page
    numbers), or both. Cells in the first column that do not match `step_pattern`
    (headers, blanks) are skipped. Duplicates are removed, order is kept.

    Anything that may need a human look (text in the step-name column that is not
    a step name, step-like words that were skipped) is appended to `warnings`.
    """
    if warnings is None:
        warnings = []
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
                steps = _steps_by_position(regions, get, regex, log, warnings)
            else:
                steps = _steps_from_tables(regions, get, strategy, regex, diagnostics, warnings)
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


def _steps_by_position(regions, get, regex, log, warnings) -> list[str]:
    """Step names are the words in the leftmost column that match the pattern.

    Works for "tables" that are really framed blocks (one block per step, the name
    in a box on the left), where no regular rows and columns can be detected.
    """
    found = []  # (page_index, top, x0, x1, name)
    furniture = []  # step-like words skipped as page header/footer
    for region in regions:
        page = get(region.page_index)
        for w in page.words():
            if not (region.top - 1 <= w["top"] < region.bottom):
                continue
            name = w["text"].strip(_TRIM)
            if name and regex.match(name):
                entry = (region.page_index, w["top"], w["x0"], w["x1"], name)
                (furniture if page.is_repeated(name, w["top"]) else found).append(entry)
    if not found:
        return []

    # Group the words into vertical columns: a word belongs to a column when its
    # centre lies within the column's first (leftmost) word, or the other way round,
    # so centred names of different lengths still line up.
    def same_column(word, first):
        centre = (word[2] + word[3]) / 2
        x0, x1 = first[2], first[3]
        return x0 <= centre <= x1 or word[2] <= (x0 + x1) / 2 <= word[3]

    columns: list[list] = []
    for word in sorted(found, key=lambda w: w[2]):
        for col in columns:
            if same_column(word, col[0]):
                col.append(word)
                break
        else:
            columns.append([word])

    # Leftmost column holding at least two names (a stray word left of the table
    # is ignored); a single name only counts when nothing else was found.
    column = next((c for c in columns if len(c) >= 2), columns[0])
    column.sort(key=lambda w: (w[0], w[1]))
    steps: list[str] = []
    per_page: dict[int, list[str]] = {}
    for page_index, _, _, _, name in column:
        if name not in steps:
            steps.append(name)
            per_page.setdefault(page_index, []).append(name)
    log(f"  Step names read from the left-hand column ({len(steps)} found):")
    for page_index, names in per_page.items():
        log(f"    page {page_index + 1}: {', '.join(names)}")

    # --- Safety checks: never drop something silently. ---
    col_x0 = min(w[2] for w in column)
    col_x1 = max(w[3] for w in column)
    seen: set[str] = set()

    def warn(message: str):
        if message not in seen and len(seen) < 100:
            seen.add(message)
            warnings.append(message)

    skipped: dict[str, list[int]] = {}
    for word in furniture:
        if col_x0 - 5 <= (word[2] + word[3]) / 2 <= col_x1 + 5:
            skipped.setdefault(word[4], [])
            if word[0] + 1 not in skipped[word[4]]:
                skipped[word[4]].append(word[0] + 1)
    for text, page_numbers in skipped.items():
        warn(f"PDF 1: '{text}' was skipped as a page header/footer (same text at the same "
             f"place on neighbouring pages; seen on page(s) {', '.join(map(str, page_numbers))}).")

    taken = {(w[0], round(w[1], 1), w[4]) for w in column}
    for region in regions:
        page = get(region.page_index)
        for w in page.words():
            top = w["top"]
            if not (region.top - 1 <= top < region.bottom):
                continue
            if abs(top - region.top) < 3 and region is regions[0]:
                continue  # the section heading itself
            centre = (w["x0"] + w["x1"]) / 2
            if not (col_x0 - 5 <= centre <= col_x1 + 5):
                continue
            text = w["text"].strip(_TRIM)
            if not text or (region.page_index, round(top, 1), text) in taken:
                continue
            if page.is_repeated(text, top):
                continue  # page header/footer text (step-like ones are reported above)
            warn(f"PDF 1 page {region.page_index + 1}: '{w['text']}' is in the step-name "
                 "column but was not taken as a step name.")
    return steps


def _steps_from_tables(regions, get, strategy, regex, diagnostics, warnings) -> list[str]:
    steps: list[str] = []
    for region in regions:
        page = get(region.page_index)
        for table in page.tables(strategy):
            for row_box, row in zip(table.rows, table.extract()):
                # Only rows inside the section (the table may start above the heading).
                top = row_box.bbox[1]
                if not row or not (region.top - 1 <= top < region.bottom):
                    continue
                step = _step_from_cell(row[0], regex)
                if step and page.is_repeated(step, top):
                    warnings.append(f"PDF 1 page {region.page_index + 1}: '{step}' was skipped as "
                                    "a page header/footer (same text on a neighbouring page).")
                    continue
                if step:
                    if step not in steps:
                        steps.append(step)
                elif row[0] and len(diagnostics) < 12:
                    diagnostics.append(f"  page {region.page_index + 1} ({strategy}): {row[0]!r}")
    return steps
