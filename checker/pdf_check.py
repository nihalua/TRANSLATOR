"""Read the PDFs with PyMuPDF and compare them with what the Excel file claims."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pymupdf

STEP_TOKEN = r"[A-Za-z]{1,10}\d+[A-Za-z0-9]*"
_STEP_WORD = re.compile(rf"^{STEP_TOKEN}$")
_STEP_WITH_SUFFIX = re.compile(rf"^(?P<base>{STEP_TOKEN}?)(?P<suffix>(?:_[A-Za-z0-9]+)*)$")
_SECTION_NUMBER = re.compile(r"^(\d+(?:\.\d+)*)\.?$")
_TOC = re.compile(r"\.{3,}|…")
_STRIP = " \t:;,.()[]*"

ERROR, WARNING, INFO = "ERROR", "WARNING", "INFO"


@dataclass
class Issue:
    severity: str
    sequence: str
    message: str
    step: str = ""
    page: str = ""


@dataclass
class Word:
    page: int  # 1-based
    x0: float
    y0: float
    x1: float
    y1: float
    text: str

    @property
    def centre(self) -> float:
        return (self.x0 + self.x1) / 2


@dataclass
class SequenceReport:
    sequence: str
    excel_steps: int = 0
    steps_confirmed: int = 0
    numbers_claimed: int = 0
    numbers_confirmed: int = 0
    issues: list[Issue] = field(default_factory=list)

    def add(self, severity, message, step="", page=""):
        self.issues.append(Issue(severity, self.sequence, message, step, str(page)))

    @property
    def errors(self) -> int:
        return sum(i.severity == ERROR for i in self.issues)

    @property
    def warnings(self) -> int:
        return sum(i.severity == WARNING for i in self.issues)


def parse_pages(text: str | None, n_pages: int, what: str) -> list[int]:
    """"3-7" / "3,5,8-9" -> page numbers (1-based). Empty -> all pages."""
    if not text or not text.strip():
        return list(range(1, n_pages + 1))
    pages: list[int] = []
    for part in re.split(r"[,;\s]+", text.strip()):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)(?:[-–](\d+))?", part)
        if not m:
            raise ValueError(f"{what}: invalid page range {text!r}")
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if a < 1 or b < a or b > n_pages:
            raise ValueError(f"{what}: page range {text!r} is outside the document "
                             f"(1-{n_pages})")
        pages += [p for p in range(a, b + 1) if p not in pages]
    return pages


class Pdf:
    """A PDF opened with PyMuPDF, with cached words per page."""

    def __init__(self, path: str):
        self.doc = pymupdf.open(path)
        self.n = self.doc.page_count
        self._words: dict[int, list[Word]] = {}
        self._text: dict[int, str] = {}

    def words(self, page: int) -> list[Word]:
        if page not in self._words:
            raw = self.doc[page - 1].get_text("words", sort=True)
            self._words[page] = [Word(page, w[0], w[1], w[2], w[3], w[4]) for w in raw]
        return self._words[page]

    def text(self, page: int) -> str:
        if page not in self._text:
            self._text[page] = self.doc[page - 1].get_text("text", sort=True)
        return self._text[page]

    def lines(self, page: int) -> list[list[Word]]:
        """Words grouped into visual lines (same height), left to right."""
        out: list[list[Word]] = []
        for w in sorted(self.words(page), key=lambda w: (round(w.y0), w.x0)):
            if out and abs(out[-1][0].y0 - w.y0) <= 2.5:
                out[-1].append(w)
            else:
                out.append([w])
        for line in out:
            line.sort(key=lambda w: w.x0)
        return out

    def repeats_nearby(self, word: Word, text: str) -> bool:
        """Page header/footer: text in the top/bottom 15% of the page that is at the same
        place on at least two neighbouring pages, one of them right before or after.

        Similar sequences use the same step names at the same place on nearby pages,
        so a single repetition in the body of the page does not make a header."""
        height = self.doc[word.page - 1].rect.height
        if 0.15 * height < word.y0 < 0.85 * height:
            return False
        near = [p for p in (word.page - 2, word.page - 1, word.page + 1, word.page + 2)
                if 1 <= p <= self.n]
        hits = [p for p in near
                if any(other.text.strip(_STRIP) == text and abs(other.y0 - word.y0) <= 6
                       and abs(other.x0 - word.x0) <= 6 for other in self.words(p))]
        return any(abs(p - word.page) == 1 for p in hits) and len(hits) >= min(2, len(near))

    def close(self):
        self.doc.close()


def _prefix(step: str) -> str:
    m = re.match(r"[A-Za-z]+", step)
    return m.group().upper() if m else ""


def _is_toc(line_text: str) -> bool:
    return bool(_TOC.search(line_text))


# ---------------------------------------------------------------------- PDF 1

def _heading_positions(pdf: Pdf, pages: list[int], section: str, others: set[str]):
    """Return (start, end): start = (page, y) of the section heading or None;
    end = (page, y) of the next heading after it or None."""
    section = section.strip().rstrip(".")
    start = None
    heading_x = None
    for p in pages:
        for line in pdf.lines(p):
            m = _SECTION_NUMBER.match(line[0].text)
            if not m or len(line) < 2 or not line[1].text[:1].isalpha():
                continue
            text = " ".join(w.text for w in line)
            if _is_toc(text):
                continue
            number = m.group(1)
            if start is None:
                if number == section:
                    start, heading_x = (p, line[0].y0), line[0].x0
                continue
            if (p, line[0].y0) <= start or number.startswith(section + "."):
                continue
            if abs(line[0].x0 - heading_x) > 8:
                continue  # not aligned with the section heading: numbered text, not a heading
            major = section.split(".")[0]
            if number in others or number == str(int(major) + 1) or (
                    number.count(".") == section.count(".") and number != section):
                return start, (p, line[0].y0)
    return start, None


def _in_region(word: Word, start, end) -> bool:
    pos = (word.page, word.y0)
    if start and pos < (start[0], start[1] - 1):
        return False
    if end and pos >= (end[0], end[1] - 1):
        return False
    return True


def check_pdf1(pdf: Pdf, sequence: str, section: str | None, pages_text: str | None,
               other_sections: set[str], excel_steps: list[str], report: SequenceReport):
    pages = parse_pages(pages_text, pdf.n, f"{sequence} PDF 1 pages")
    start = end = None
    if section:
        start, end = _heading_positions(pdf, pages, section, other_sections)
        if start is None:
            elsewhere = [p for p in range(1, pdf.n + 1) if p not in pages and
                         _heading_positions(pdf, [p], section, other_sections)[0]]
            where = f" It is on page {elsewhere[0]}." if elsewhere else ""
            report.add(WARNING, f"Heading {section} not found on PDF 1 pages {pages_text}."
                       f"{where} Whole pages were checked.")
        if end is None and pages_text and pages[-1] < pdf.n:
            # Does the section continue after the given pages?
            nxt = pages[-1] + 1
            _, next_end = _heading_positions(pdf, [nxt], section, other_sections)
            stop_y = next_end[1] if next_end else None
            cont = [w for w in pdf.words(nxt)
                    if _STEP_WORD.match(w.text.strip(_STRIP)) and (stop_y is None or w.y0 < stop_y)
                    and _prefix(w.text) in {_prefix(s) for s in excel_steps}
                    and not pdf.repeats_nearby(w, w.text.strip(_STRIP))]
            if cont and not _heading_on_top(pdf, nxt, other_sections, section):
                names = ", ".join(dict.fromkeys(w.text.strip(_STRIP) for w in cont))
                report.add(ERROR, f"The section may continue on PDF 1 page {nxt}, which is not in "
                           f"the page range {pages_text}. Step-like words there: {names}",
                           page=nxt)

    candidates: list[Word] = []
    for p in pages:
        for w in pdf.words(p):
            text = w.text.strip(_STRIP)
            if _STEP_WORD.match(text) and _in_region(w, start, end):
                candidates.append(Word(w.page, w.x0, w.y0, w.x1, w.y1, text))

    excel_set = set(excel_steps)
    for step in dict.fromkeys(s for s in excel_steps if excel_steps.count(s) > 1):
        report.add(ERROR, f"{step} appears more than once in the Excel sheet.", step)

    # Where are the Excel step names in the PDF? Their positions give the step column.
    found = {s: [w for w in candidates if w.text == s] for s in excel_set}
    first_hits = [ws[0] for ws in found.values() if ws]
    if first_hits:
        centres = sorted(w.centre for w in first_hits)
        median = centres[len(centres) // 2]
        column = [w for ws in found.values() for w in ws if abs(w.centre - median) <= 40]
        col_x0 = min(w.x0 for w in column) - 8
        col_x1 = max(w.x1 for w in column) + 8
    else:
        col_x0 = col_x1 = None

    def in_column(w: Word) -> bool:
        return col_x0 is not None and col_x0 <= w.centre <= col_x1

    for step in excel_steps:
        hits = found.get(step) or []
        if not hits:
            report.add(ERROR, f"{step} is in the Excel sheet but was not found in the PDF 1 "
                       "section.", step)
        elif not any(in_column(w) for w in hits):
            report.add(WARNING, f"{step} is in the Excel sheet but in PDF 1 it only appears in "
                       "running text, not in the step-name column.", step,
                       ", ".join(str(p) for p in sorted({w.page for w in hits})))
        else:
            report.steps_confirmed += 1

    prefixes = {_prefix(s) for s in excel_steps}
    info: dict[str, set[int]] = {}
    for w in candidates:
        if w.text in excel_set:
            continue
        below_heading = (start is not None and w.page == start[0] and w.y0 > start[1]
                         and w.y0 < pdf.doc[w.page - 1].rect.height / 2)
        if in_column(w) and (below_heading or not pdf.repeats_nearby(w, w.text)):
            report.add(ERROR, f"{w.text} is in the step-name column of PDF 1 but NOT in the "
                       "Excel sheet (missed step?).", w.text, w.page)
        elif _prefix(w.text) in prefixes:
            info.setdefault(w.text, set()).add(w.page)
    for text, pages_seen in info.items():
        report.add(INFO, f"{text} is mentioned in PDF 1 outside the step-name column "
                   "(reference in a description?).", text, ", ".join(map(str, sorted(pages_seen))))


def _heading_on_top(pdf: Pdf, page: int, others: set[str], section: str) -> bool:
    """True when the first real line of `page` is another section's heading."""
    for line in pdf.lines(page):
        text = " ".join(w.text for w in line)
        if pdf.repeats_nearby(line[0], line[0].text.strip(_STRIP)):
            continue  # page header
        m = _SECTION_NUMBER.match(line[0].text)
        return bool(m and m.group(1) != section and m.group(1) in others and not _is_toc(text))
    return False


# ---------------------------------------------------------------------- PDF 2

def _label_regex(label: str) -> re.Pattern:
    label = label.strip().rstrip(":")
    return re.compile(rf"(?i)^{re.escape(label)}\s*:?$")


def check_pdf2(pdf: Pdf, sequence: str, pages_text: str | None, label: str,
               claimed: dict[str, list[str]], report: SequenceReport):
    """claimed: step -> numbers written in the Excel sheet."""
    pages = parse_pages(pages_text, pdf.n, f"{sequence} PDF 2 pages")
    label_re = _label_regex(label)
    bare = label.strip().rstrip(":")
    steps = set(claimed)

    # Every "number:" field in the pages, with the step it belongs to: the nearest
    # step-like name ABOVE it on the page (geometrically), which is independent of
    # the text order the other program uses.
    found: list[tuple[str, str, str, int]] = []  # (base step, variant, value, page)
    unassigned: list[tuple[str, int, str]] = []
    for p in pages:
        words = pdf.words(p)
        lines = pdf.lines(p)
        for li, line in enumerate(lines):
            for wi, w in enumerate(line):
                t = w.text
                glued = t.lower().startswith(bare.lower() + ":") and len(t) > len(bare) + 1
                if not (glued or label_re.match(t) or (
                        t.lower() == bare.lower() and wi + 1 < len(line)
                        and line[wi + 1].text.startswith(":"))):
                    continue
                # value: rest of the text after the label on the same line, or next line
                rest = [x for x in line[wi + 1:] if x.text != ":"]
                if glued:  # "number:101"
                    rest = [Word(p, w.x0, w.y0, w.x1, w.y1, t[len(bare) + 1:])]
                if rest and rest[0].text.startswith(":"):
                    rest[0] = Word(p, rest[0].x0, rest[0].y0, rest[0].x1, rest[0].y1,
                                   rest[0].text[1:])
                    if not rest[0].text:
                        rest = rest[1:]
                if rest:
                    value = rest[0].text
                elif li + 1 < len(lines):
                    value = lines[li + 1][0].text
                else:
                    value = ""
                if not value:
                    report.add(WARNING, "A 'number:' field without a value.", page=p)
                    continue
                owner = None
                best = None
                for cand in words:
                    m = _STEP_WITH_SUFFIX.match(cand.text.strip(_STRIP))
                    if not m or m.group("base") not in steps:
                        continue
                    if cand.y0 > w.y0 + 2:
                        continue  # below the label
                    if abs(cand.y0 - w.y0) <= 2 and cand.x0 > w.x0:
                        continue  # right of the label on the same line
                    dist = (w.y0 - cand.y0, abs(w.x0 - cand.x0))
                    if best is None or dist < best:
                        best, owner = dist, (m.group("base"), cand.text.strip(_STRIP))
                if owner:
                    found.append((owner[0], owner[1], value, p))
                else:
                    unassigned.append((value, p, _page_title(pdf, p)))

    by_step: dict[str, list[tuple[str, str, int]]] = {}
    for base, variant, value, p in found:
        by_step.setdefault(base, []).append((variant, value, p))

    all_claimed = {n for nums in claimed.values() for n in nums}
    for step, numbers in claimed.items():
        report.numbers_claimed += len(numbers)
        pdf_values = {v for _, v, _ in by_step.get(step, [])}
        if not numbers:
            report.add(ERROR, f"{step} has no number in the Excel sheet.", step)
        if len(numbers) > 1:
            report.add(INFO, f"{step} still has alternatives: {' OR '.join(numbers)}.", step)
        for n in numbers:
            if n in pdf_values:
                report.numbers_confirmed += 1
                continue
            pages_with = sorted({pg for _, _, v, pg in found if v == n})
            report.add(ERROR, f"{step}: number {n} from the Excel sheet was not confirmed in "
                       "PDF 2" + (f" (on page {', '.join(map(str, pages_with))} it belongs to "
                                  f"{', '.join(sorted({var for b, var, v, pg in found if v == n}))})"
                                  if pages_with else ""), step,
                       ", ".join(map(str, pages_with)))
        for variant, value, p in by_step.get(step, []):
            if value not in numbers:
                report.add(ERROR, f"PDF 2 page {p}: {variant} has number {value}, which is NOT in "
                           f"the Excel row of {step}.", step, p)

    for value, p, title in unassigned:
        if value not in all_claimed:
            report.add(INFO, f"PDF 2 page {p}: number {value} does not belong to any step of "
                       f"this sheet ({title}).", page=p)


def _page_title(pdf: Pdf, page: int) -> str:
    lines = pdf.lines(page)
    return " ".join(w.text for w in lines[0])[:60] if lines else "empty page"
