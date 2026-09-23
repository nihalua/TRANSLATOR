from __future__ import annotations

import pytest
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from translator.cli import main
from translator.pdf_numbers import find_step_numbers
from translator.pdf_steps import SectionNotFoundError, extract_step_names

STYLES = getSampleStyleSheet()
GRID = TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)])


def _table(rows):
    t = Table(rows, repeatRows=1)
    t.setStyle(GRID)
    return t


def make_pdf1(path, with_bookmarks: bool, many_rows: bool = False):
    step_rows = [["Step", "Description"], ["T00", "Init"], ["T01", "Warm up"],
                 ["S01", "First"], ["S02", "Second"]]
    if many_rows:
        step_rows += [[f"S{i:02d}", f"Step {i}"] for i in range(3, 60)]

    story = [
        Paragraph("Table of contents", STYLES["Heading1"]),
        Paragraph("2 Overview ........ 2", STYLES["Normal"]),
        Paragraph("3.1 Step list ........ 3", STYLES["Normal"]),
        Paragraph("3.2 Other steps ........ 4", STYLES["Normal"]),
        PageBreak(),
        Paragraph("2 Overview", STYLES["Heading1"]),
        _table([["Item", "Value"], ["X01", "wrong table"]]),
        PageBreak(),
        Paragraph("3.1 Step list", STYLES["Heading2"]),
        Spacer(1, 12),
        _table(step_rows),
        Paragraph("3.2 Other steps", STYLES["Heading2"]),
        _table([["Step", "Description"], ["Z99", "not in 3.1"]]),
    ]

    doc = SimpleDocTemplate(str(path), pagesize=A4)
    if with_bookmarks:
        class Doc(SimpleDocTemplate):
            def afterFlowable(self, flowable):
                if isinstance(flowable, Paragraph) and flowable.style.name.startswith("Heading"):
                    text = flowable.getPlainText()
                    if text[0].isdigit():
                        key = text.replace(" ", "_")
                        self.canv.bookmarkPage(key)
                        self.canv.addOutlineEntry(text, key, 0)
        doc = Doc(str(path), pagesize=A4)
    doc.build(story)


def make_pdf2(path):
    pages = [
        ["Step properties", "Name: T00", "number: 100", "type: start"],
        ["Name: T01", "number: 101"],
        ["Name: S01_1", "number: 201", "previous: T01"],
        ["Name: S01_2", "number:", "202"],  # value on the next line
        ["Name: S01_3", "Number: 203"],
        ["Name: S010", "number: 999"],  # must not be taken for S01
        ["Name: S02", "number: 300"],
        ["Name: Z99", "number: 777"],
    ]
    c = canvas.Canvas(str(path), pagesize=A4)
    for lines in pages:
        y = 800
        for line in lines:
            c.drawString(72, y, line)
            y -= 20
        c.showPage()
    c.save()


@pytest.fixture
def pdfs(tmp_path):
    p1, p2 = tmp_path / "steps.pdf", tmp_path / "numbers.pdf"
    make_pdf1(p1, with_bookmarks=False)
    make_pdf2(p2)
    return p1, p2


@pytest.mark.parametrize("with_bookmarks", [False, True])
def test_extract_steps_from_section(tmp_path, with_bookmarks):
    p1 = tmp_path / "steps.pdf"
    make_pdf1(p1, with_bookmarks=with_bookmarks)
    assert extract_step_names(str(p1), "3.1") == ["T00", "T01", "S01", "S02"]
    assert extract_step_names(str(p1), "3.2") == ["Z99"]


def test_table_spanning_pages(tmp_path):
    p1 = tmp_path / "steps.pdf"
    make_pdf1(p1, with_bookmarks=False, many_rows=True)
    steps = extract_step_names(str(p1), "3.1")
    assert steps[:4] == ["T00", "T01", "S01", "S02"]
    assert steps[-1] == "S59"
    assert "Z99" not in steps


def test_missing_section(pdfs):
    with pytest.raises(SectionNotFoundError):
        extract_step_names(str(pdfs[0]), "7.4")


def test_numbers_with_variants(pdfs):
    res = find_step_numbers(str(pdfs[1]), ["T00", "T01", "S01", "S02", "S03"])
    assert res["T00"].numbers == ["100"]
    assert res["T01"].numbers == ["101"]
    assert res["S01"].numbers == ["201", "202", "203"]
    assert [m.variant for m in res["S01"].matches] == ["S01_1", "S01_2", "S01_3"]
    assert res["S02"].numbers == ["300"]
    assert res["S03"].numbers == []


def test_name_label(pdfs):
    res = find_step_numbers(str(pdfs[1]), ["T01", "S01"], name_label="Name")
    assert res["T01"].numbers == ["101"]
    assert res["S01"].numbers == ["201", "202", "203"]


def test_end_to_end_excel(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    rc = main(["--pdf1", str(pdfs[0]), "--section", "3.1", "--pdf2", str(pdfs[1]),
               "--output", str(out)])
    assert rc == 0
    ws = load_workbook(out).active
    rows = [tuple(c.value for c in r) for r in ws.iter_rows(min_row=1, max_col=3)]
    assert rows == [
        ("Step", "Step no.", "Number"),
        ("S01", 1, "201 OR 202 OR 203"),
        ("S02", 2, 300),
        ("T00", 0, 100),
        ("T01", 1, 101),
    ]


def test_custom_layout(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    main(["--pdf1", str(pdfs[0]), "--section", "3.1", "--pdf2", str(pdfs[1]),
          "--output", str(out), "--step-column", "C", "--number-column", "E",
          "--start-row", "3", "--step-header", "Name", "--number-header", "No."])
    ws = load_workbook(out).active
    assert ws["C3"].value == "Name" and ws["E3"].value == "No."
    assert ws["C4"].value == "S01" and ws["E4"].value == "201 OR 202 OR 203"


def test_document_order_and_pdf2_pages(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    main(["--pdf1", str(pdfs[0]), "--section", "3.1", "--pdf2", str(pdfs[1]),
          "--pdf2-pages", "1-3", "--output", str(out), "--no-sort", "--step-number-column", ""])
    ws = load_workbook(out).active
    rows = [tuple(c.value for c in r) for r in ws.iter_rows(min_row=1, max_col=3)]
    # Pages 4-7 are not searched: S01 keeps only the number from page 3, S02 has none.
    assert rows == [
        ("Step", None, "Number"),
        ("T00", None, 100),
        ("T01", None, 101),
        ("S01", None, 201),
        ("S02", None, None),
    ]


def test_natural_sort():
    from translator.excel_writer import natural_key, step_digits

    assert sorted(["T10", "S10", "T0", "S02", "J30", "T2"], key=natural_key) == \
        ["J30", "S02", "S10", "T0", "T2", "T10"]
    assert step_digits("T00") == 0 and step_digits("S02") == 2 and step_digits("T45") == 45


def make_borderless_pdf(path):
    """Table without ruling lines, cells like "T00 Start", spanning two pages with a page header."""
    c = canvas.Canvas(str(path), pagesize=A4)
    c.drawString(72, 800, "Contents")
    c.drawString(72, 780, "3.1 Sequence steps ........ 2")
    c.showPage()

    def header():
        c.drawString(72, 815, "2024 Electrolyser specification")

    header()
    c.drawString(72, 740, "3.1 Sequence steps")
    y = 710
    rows = [("Step", "Description")] + [(f"T0{i} Start", f"Task {i}") for i in range(3)] + \
           [(f"S{i:02d}", f"Step {i}") for i in range(1, 30)]
    for name, desc in rows:
        if y < 80:
            c.showPage()
            header()
            y = 760
        c.drawString(72, y, name)
        c.drawString(250, y, desc)
        y -= 22
    c.drawString(72, y - 20, "3.2 Other section")
    c.drawString(72, y - 50, "Z99")
    c.drawString(250, y - 50, "not in 3.1")
    c.showPage()
    c.save()


def test_borderless_table_spanning_pages(tmp_path):
    p1 = tmp_path / "borderless.pdf"
    make_borderless_pdf(p1)
    expected = ["T00", "T01", "T02"] + [f"S{i:02d}" for i in range(1, 30)]
    assert extract_step_names(str(p1), "3.1") == expected
    assert extract_step_names(str(p1), "3.1", pages=[2, 3]) == expected
    # Page range only (no section): everything on those pages.
    assert extract_step_names(str(p1), pages=[2, 3]) == expected + ["Z99"]


def test_page_range_with_table(pdfs):
    assert extract_step_names(str(pdfs[0]), pages=[3]) == ["T00", "T01", "S01", "S02", "Z99"]
    assert extract_step_names(str(pdfs[0]), "3.1", pages=[3]) == ["T00", "T01", "S01", "S02"]


def test_parse_page_range():
    from translator.pdf_steps import parse_page_range

    assert parse_page_range("3-7") == [3, 4, 5, 6, 7]
    assert parse_page_range("3, 5, 8-9") == [3, 5, 8, 9]
    with pytest.raises(ValueError):
        parse_page_range("7-3")


def test_no_steps_reports_cells(pdfs):
    from translator.pdf_steps import StepsNotFoundError

    with pytest.raises(StepsNotFoundError) as err:
        extract_step_names(str(pdfs[0]), "3.1", step_pattern=r"^Q\d+$")
    assert "'T00'" in str(err.value)


def make_sequence_chart_pdf(path):
    """Layout like a sequence description: one framed block per step, the step name in a
    small box on the left, free text (with step-like words) in a big box on the right."""
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4

    def page_header():
        c.drawString(72, height - 30, "EL2024 Sequence description")  # step-like, in margin

    page_header()
    c.drawString(60, 780, "3.1 Sequence")
    c.drawString(530, 755, "Rev")
    y = 740
    blocks = [("T00", "", 10), ("S02", "S", 6), ("T02", "&", 1), ("S05", "S", 3),
              ("S07", "S", 8), ("T05", "&", 1), ("S09", "S", 4)]
    for name, symbol, n_lines in blocks:
        block_h = 20 + n_lines * 16
        if y - block_h < 70:
            c.showPage()
            page_header()
            y = 780
        c.rect(60, y - block_h, 50, block_h)              # name column
        c.rect(110, y - block_h, 20, block_h)             # symbol column
        c.rect(130, y - block_h, 390, block_h)            # description
        name_y = y - block_h / 2 if name == "S05" else y - 14  # S05 centred like the screenshot
        c.drawString(72, name_y, name)
        c.drawString(115, y - 14, symbol)
        for i in range(n_lines):
            text = f"XV10{i} open valve" if i % 2 else f"- check S03 {i}"
            c.drawString(140, y - 14 - i * 16, text)
        if name == "T00":
            c.drawString(140, y - block_h + 4, "Note: P1 must be running")
        y -= block_h + 15
    c.drawString(60, y - 20, "3.2 Next section")
    c.rect(60, y - 60, 50, 25)
    c.drawString(72, y - 52, "Z99")
    c.showPage()
    c.save()


def test_sequence_chart_layout(tmp_path):
    p1 = tmp_path / "chart.pdf"
    make_sequence_chart_pdf(p1)
    expected = ["T00", "S02", "T02", "S05", "S07", "T05", "S09"]
    assert extract_step_names(str(p1), "3.1") == expected
    assert extract_step_names(str(p1), "3.1", pages=[1, 2]) == expected


def make_edge_pdf(path, extra_in_column: str | None = None):
    """Framed-block layout with a step name at the very bottom of a page (like T70),
    one at the very top of the next page, and a repeating step-like page header."""
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4

    def furniture(page_no):
        c.drawString(72, height - 25, "EL2024")              # header, left, step-like
        c.drawString(72, 20, f"DOC7 page {page_no}")         # footer, left, step-like

    furniture(1)
    c.drawString(60, 780, "3.1 Sequence")
    y = 750
    for name in ["T00", "S02", "T02", "S05"]:
        c.rect(60, y - 60, 50, 60)
        c.rect(110, y - 60, 400, 60)
        c.drawString(72, y - 14, name)
        c.drawString(120, y - 14, "check S03 and XV101")
        y -= 75
    if extra_in_column:
        c.drawString(72, y - 14, extra_in_column)
    # T70 in the bottom 7% of the page, just above the footer
    c.rect(60, 38, 450, 18)
    c.drawString(72, 42, "T70")
    c.drawString(120, 42, "& Empty transition")
    c.showPage()

    furniture(2)
    # S75 at the very top of the next page
    c.drawString(72, height - 45, "S75")
    c.drawString(120, height - 45, "MESSAGE")
    c.drawString(60, height - 120, "3.2 Next section")
    c.drawString(72, height - 150, "Z99")
    c.showPage()
    furniture(3)
    c.drawString(72, 700, "Other text")
    c.showPage()
    c.save()


def test_steps_at_page_edges_are_kept(tmp_path):
    p1 = tmp_path / "edges.pdf"
    make_edge_pdf(p1)
    warnings = []
    steps = extract_step_names(str(p1), "3.1", warnings=warnings)
    assert steps == ["T00", "S02", "T02", "S05", "T70", "S75"]
    # Only the step-like page header/footer is mentioned, once each.
    assert len(warnings) == 2
    assert all("page header/footer" in w for w in warnings)
    assert extract_step_names(str(p1), pages=[1, 2]) == ["T00", "S02", "T02", "S05", "T70", "S75", "Z99"]


def test_unexpected_text_in_step_column_is_reported(tmp_path, pdfs):
    p1 = tmp_path / "edges.pdf"
    make_edge_pdf(p1, extra_in_column="T 60")
    warnings = []
    steps = extract_step_names(str(p1), "3.1", warnings=warnings)
    assert "T70" in steps
    assert any("'T'" in w for w in warnings) and any("'60'" in w for w in warnings)

    out = tmp_path / "mapping.xlsx"
    main(["--pdf1", str(p1), "--section", "3.1", "--pdf2", str(pdfs[1]), "--output", str(out)])
    wb = load_workbook(out)
    assert "Check" in wb.sheetnames
    texts = [c.value for c in wb["Check"]["A"]]
    assert any("'60'" in (t or "") for t in texts)
    assert any("No number found" in (t or "") for t in texts)


def test_multiple_sequences_one_sheet_each(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    rc = main(["--pdf1", str(pdfs[0]), "--pdf2", str(pdfs[1]), "--output", str(out),
               "--sequence", "3.1:3:1-5",      # T00, T01, S01 (S01_1.._3), S02 not in pages
               "--sequence", "3.2:3:",         # Z99, all PDF 2 pages
               "--sequence", "3.9:99:"])       # bad page -> error sheet, others still written
    assert rc == 1
    wb = load_workbook(out)
    assert wb.sheetnames == ["Sequence 01", "Sequence 02", "Sequence 03", "Check"]
    rows = [r for r in wb["Sequence 01"].iter_rows(min_row=2, max_col=3, values_only=True)]
    assert rows == [("S01", 1, "201 OR 202 OR 203"), ("S02", 2, None), ("T00", 0, 100),
                    ("T01", 1, 101)]
    assert [r for r in wb["Sequence 02"].iter_rows(min_row=2, max_col=3, values_only=True)] == \
        [("Z99", 99, 777)]
    assert "ERROR" in wb["Sequence 03"]["A2"].value
    checks = [c.value for c in wb["Check"]["A"]]
    assert any(c.startswith("Sequence 01: No number found in PDF 2 for: S02") for c in checks[1:])
    assert any(c.startswith("Sequence 03: ERROR") for c in checks[1:])
