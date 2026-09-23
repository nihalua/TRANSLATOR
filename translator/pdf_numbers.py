"""PDF 2: find the "number:" value of every step (and its variants such as S01_1, S01_2)."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pdfplumber

DEFAULT_NUMBER_LABEL = "number:"


@dataclass
class Match:
    variant: str  # step name as written in PDF 2, e.g. "S01_2"
    number: str
    page: int  # 1-based page number


@dataclass
class StepResult:
    step: str
    matches: list[Match] = field(default_factory=list)

    @property
    def numbers(self) -> list[str]:
        out: list[str] = []
        for m in self.matches:
            if m.number not in out:
                out.append(m.number)
        return out

    @property
    def pages(self) -> list[int]:
        return sorted({m.page for m in self.matches})


def _label_regex(label: str) -> re.Pattern:
    label = label.strip()
    if not label.endswith(":"):
        label += ":"
    body = re.escape(label[:-1]).replace(r"\ ", r"\s+")
    # "number:" but not "part_number:" / "serial number:" style prefixes glued to it.
    return re.compile(rf"(?<![A-Za-z0-9_]){body}\s*:[ \t]*(?P<value>[^\n]*)", re.IGNORECASE)


def _step_regex(steps: list[str]) -> re.Pattern:
    # Exact step name, optionally followed by "_<suffix>" parts (S01, S01_1, S01_2a ...),
    # but never a longer name such as S010.
    alternatives = "|".join(re.escape(s) for s in sorted(steps, key=len, reverse=True))
    return re.compile(
        rf"(?<![A-Za-z0-9_])(?P<base>{alternatives})(?P<suffix>(?:_[A-Za-z0-9]+)*)(?![A-Za-z0-9_])"
    )


def _value_after(match: re.Match, text: str) -> str | None:
    value = match.group("value").strip()
    if not value:
        # Value printed on the next non-empty line.
        for line in text[match.end():].splitlines():
            if line.strip():
                value = line.strip()
                break
    if not value:
        return None
    return value.split()[0]


def find_step_numbers(
    pdf_path: str,
    steps: list[str],
    number_label: str = DEFAULT_NUMBER_LABEL,
    name_label: str | None = None,
) -> dict[str, StepResult]:
    """Return {step: StepResult} with every "number:" value found for each step.

    A "number:" field is attributed to the closest step name (or variant) written
    before it on the same page. If `name_label` is given (e.g. "name:"), only step
    names written after that label are taken into account, which is more precise
    when the property pages also mention other steps.
    """
    results = {s: StepResult(s) for s in steps}
    if not steps:
        return results

    number_re = _label_regex(number_label)
    step_re = _step_regex(steps)
    name_re = _label_regex(name_label) if name_label else None

    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            if not number_re.search(text):
                continue

            # (position, base step, variant) of every step name on the page.
            names: list[tuple[int, str, str]] = []
            if name_re:
                for m in name_re.finditer(text):
                    sm = step_re.search(m.group("value"))
                    if sm:
                        names.append((m.start(), sm.group("base"), sm.group(0)))
            else:
                for sm in step_re.finditer(text):
                    names.append((sm.start(), sm.group("base"), sm.group(0)))

            for nm in number_re.finditer(text):
                value = _value_after(nm, text)
                if value is None:
                    continue
                before = [n for n in names if n[0] < nm.start()]
                if not before:
                    continue
                _, base, variant = before[-1]
                results[base].matches.append(Match(variant, value, page_no))

    return results
