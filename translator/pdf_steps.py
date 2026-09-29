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
# Dotted leaders only: a heading may end in a number ("3.6 Sequence Stack 2").
_TOC_LEADERS = re.compile(r"\.{3,}|…|\s\.\s\.")
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
    # True when `top` is the section heading: nothing below it is a page header
    # (a footer still can be).
    body_from_top: bool = False

    def furniture(self, page, text: str, top: float) -> bool:
        if self.body_from_top and self.top + 1 < top < float(page.page.height) / 2:
            return False
        return page.is_repeated(text, top)


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


def _ends_section(num: str, section: str) -> bool:
    """True when heading `num` can follow `section` (3.6 -> 3.7, 4, 4.1, ...).

    Numbered text inside the table ("1 Start", "10 s wait") is not a heading: its
    number is lower than the section, or jumps more than one chapter ahead."""
    if _is_subsection(num, section):
        return False
    a = [int(x) for x in num.split(".")]
    b = [int(x) for x in section.strip().rstrip(".").split(".")]
    return a > b and a[0] <= b[0] + 1


# Tolerance (points) for a heading to count as aligned with the section heading.
HEADING_X_TOLERANCE = 8


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
    # Page headers/footers only live in this outer band of the page (top/bottom).
    FURNITURE_BAND = 0.15
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
        """True for page headers/footers: text in the top/bottom band of the page that
        is at the same height on most of the neighbouring pages.

        Similar sequences often use the same step names (T00, S01, ...) at the same
        place on nearby pages, so being repeated once is not enough: a header/footer
        must also sit in the outer band and repeat on at least two neighbouring pages,
        one of them directly before or after (one page when there is only one)."""
        height = float(self.page.height)
        band = height * self.FURNITURE_BAND
        if band < top < height - band:
            return False
        neighbours = [j for j in range(self.index - self.NEIGHBOURS, self.index + self.NEIGHBOURS + 1)
                      if j != self.index and 0 <= j < self._n]
        hits = [j for j in neighbours
                if any(abs(t - top) <= self.REPEAT_TOLERANCE
                       for t in self._get(j)._word_tops().get(text, []))]
        # A header/footer is on every page, so also on the page right before or after.
        adjacent = any(abs(j - self.index) == 1 for j in hits)
        return adjacent and len(hits) >= min(2, len(neighbours))

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
        """(top, number, text, x0) of numbered headings outside tables, headers/footers and TOC."""
        out = []
        for line in self.lines:
            if self.in_margin(line["top"]):
                continue
            m = _ANY_HEADING.match(line["text"])
            if m and not _TOC_LEADERS.search(line["text"]) and not self._in_table(line):
                out.append((line["top"], m.group("num"), line["text"], line["x0"]))
        return out

    def section_line(self, pattern: re.Pattern, strict: bool = True):
        """The heading line of the section. `strict` also rejects headings ending in a
        number (they look like a TOC line); use strict=False on the given pages."""
        toc = _TOC_LINE if strict else _TOC_LEADERS
        for line in self.lines:
            if pattern.match(line["text"]) and not toc.search(line["text"]):
                return line
        return None

    def looks_like_toc(self) -> bool:
        """A table-of-contents page: several numbered lines ending in a page number."""
        entries = [l for l in self.lines
                   if _ANY_HEADING.match(l["text"]) and _TOC_LINE.search(l["text"])]
        return len(entries) >= 3

    def section_top(self, pattern: re.Pattern) -> float | None:
        line = self.section_line(pattern)
        return line["top"] if line else None


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

    start_index, start_top, start_x = None, 0.0, None
    if section:
        pattern = _section_regex(section)
        if indices is not None:
            for i in indices:
                line = get(i).section_line(pattern, strict=False)
                if line is not None:
                    start_index, start_top, start_x = i, line["top"], line["x0"]
                    break
            if start_index is None:
                log(f"  Note: heading {section} not found on the given pages; using the whole pages.")
        else:
            outline_page = _find_in_outline(pdf_path, section)
            if outline_page is not None:
                start_index = outline_page
                line = get(outline_page).section_line(pattern, strict=False)
                if line is not None:
                    start_top, start_x = line["top"], line["x0"]
            else:
                # First headings that cannot be a TOC line; then also headings ending
                # in a number ("3.6 Sequence Stack 2"), skipping table-of-contents pages.
                for strict in (True, False):
                    for i in range(n):
                        if not strict and get(i).looks_like_toc():
                            continue
                        line = get(i).section_line(pattern, strict=strict)
                        if line is not None:
                            start_index, start_top, start_x = i, line["top"], line["x0"]
                            break
                    if start_index is not None:
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
        # On the page of the section heading, everything below the heading is content
        # (never a page header), even if a similar sequence has it at the same place.
        started_here = False
        if not started:
            if i != start_index:
                continue
            started, top = True, start_top
            started_here = start_top > 0
        # Stop at the next heading that is not part of the section. Only a heading
        # aligned with the section heading counts (numbered text in the table does not).
        if section and start_index is not None:
            stop = next((h for h in page.headings()
                         if h[0] > top + 1 and _ends_section(h[1], section)
                         and (start_x is None or abs(h[3] - start_x) <= HEADING_X_TOLERANCE)),
                        None)
            if stop is not None:
                regions.append(_Region(i, top, stop[0], body_from_top=started_here))
                log(f"  Section ends at '{stop[2].strip()}' on page {i + 1}.")
                break
        regions.append(_Region(i, top, bottom, body_from_top=started_here))
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
                (furniture if region.furniture(page, name, w["top"]) else found).append(entry)
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
    col_x0 = min(w[2] for w in column)
    col_x1 = max(w[3] for w in column)
    # Column position per page (normally the same on every page).
    page_cols = {i: (col_x0, col_x1) for i in {w[0] for w in column}}

    # A page where the table sits elsewhere (landscape page, other margins) has no
    # name in that column: take that page's own leftmost column instead, so its
    # steps are not lost without a word.
    prefixes = {re.match(r"[A-Za-z]*", w[4]).group().upper() for w in column}
    for page_index in sorted({w[0] for w in found} - set(page_cols)):
        shifted = _shifted_column(page_index, [w for w in found if w[0] == page_index],
                                  get(page_index), same_column, prefixes)
        if shifted:
            column += shifted
            page_cols[page_index] = (min(w[2] for w in shifted), max(w[3] for w in shifted))
            warnings.append(
                f"PDF 1 page {page_index + 1}: the step names are at another position than on "
                f"the other pages; read from there: {', '.join(w[4] for w in shifted)}. "
                "Please check this page.")
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
    seen: set[str] = set()

    def in_col(page_index, centre):
        x0, x1 = page_cols.get(page_index, (col_x0, col_x1))
        return x0 - 5 <= centre <= x1 + 5

    def warn(message: str):
        if message not in seen and len(seen) < 100:
            seen.add(message)
            warnings.append(message)

    skipped: dict[str, list[int]] = {}
    for word in furniture:
        if in_col(word[0], (word[2] + word[3]) / 2):
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
            if not in_col(region.page_index, (w["x0"] + w["x1"]) / 2):
                continue
            text = w["text"].strip(_TRIM)
            if not text or (region.page_index, round(top, 1), text) in taken:
                continue
            if region.furniture(page, text, top):
                continue  # page header/footer text (step-like ones are reported above)
            warn(f"PDF 1 page {region.page_index + 1}: '{w['text']}' is in the step-name "
                 "column but was not taken as a step name.")
    return steps


def _shifted_column(page_index, words, page, same_column, prefixes) -> list:
    """The leftmost column of step names on a page whose table is shifted, or [].

    Only accepted when it clearly is a step-name column: at least two names, with
    the same prefixes as elsewhere (S, T, ...), each the first word of its line."""
    columns: list[list] = []
    for word in sorted(words, key=lambda w: w[2]):
        for col in columns:
            if same_column(word, col[0]):
                col.append(word)
                break
        else:
            columns.append([word])
    column = next((c for c in columns if len(c) >= 2), None)
    if not column:
        return []
    for w in column:
        if re.match(r"[A-Za-z]*", w[4]).group().upper() not in prefixes:
            return []
        if any(abs(o["top"] - w[1]) < 3 and o["x1"] <= w[2] for o in page.words()):
            return []  # text left of it on the same line: a reference in a description
    return column


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
                if step and region.furniture(page, step, top):
                    warnings.append(f"PDF 1 page {region.page_index + 1}: '{step}' was skipped as "
                                    "a page header/footer (same text on a neighbouring page).")
                    continue
                if step:
                    if step not in steps:
                        steps.append(step)
                elif row[0] and len(diagnostics) < 12:
                    diagnostics.append(f"  page {region.page_index + 1} ({strategy}): {row[0]!r}")
    return steps
