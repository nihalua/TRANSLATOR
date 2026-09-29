"""Tests for the independent checker (checker package).

The Excel files are produced by TRANSLATOR, then checked; some are tampered with to
make sure the checker notices.
"""

from __future__ import annotations

from openpyxl import load_workbook

from checker.app import main as check_main
from checker.app import run_check
from test_translator import (make_edge_pdf, make_pdf1, make_pdf2, make_sequence_chart_pdf,
                             pdfs)  # noqa: F401 (pytest fixture)
from translator.cli import main as translate


def _cfg(excel, pdf1, pdf2, sequences):
    return {"excel": str(excel), "pdf1": str(pdf1), "pdf2": str(pdf2),
            "sequences": [{"sheet": f"Sequence {i:02d}", "section": s, "pages": p1,
                           "pdf2_pages": p2} for i, (s, p1, p2) in enumerate(sequences, start=1)],
            "options": {}}


def _issues(reports, severity):
    return [i.message for r in reports for i in r.issues if i.severity == severity]


def test_correct_file_passes(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    translate(["--pdf1", str(pdfs[0]), "--pdf2", str(pdfs[1]), "-o", str(out),
               "--sequence", "3.1:3:", "--sequence", "3.2:3:"])
    reports, report_file = run_check(_cfg(out, *pdfs, [("3.1", "3", ""), ("3.2", "3", "")]))
    assert _issues(reports, "ERROR") == [] and _issues(reports, "WARNING") == []
    assert reports[0].steps_confirmed == 4 and reports[0].numbers_confirmed == 6
    assert reports[1].steps_confirmed == 1 and reports[1].numbers_confirmed == 1
    summary = load_workbook(report_file)["Summary"]
    assert [c.value for c in summary["H"] if c.value][-2:] == ["OK", "OK"]
    # command line: exit code 0 when everything is confirmed
    assert check_main(["--excel", str(out), "--pdf1", str(pdfs[0]), "--pdf2", str(pdfs[1]),
                       "--sequence", "3.1:3:", "--sequence", "3.2:3:"]) == 0


def test_tampered_file_is_caught(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    translate(["--pdf1", str(pdfs[0]), "--pdf2", str(pdfs[1]), "-o", str(out),
               "--sequence", "3.1:3:"])
    wb = load_workbook(out)
    ws = wb["Sequence 01"]
    # rows: 1 header, 2 S01, 3 S02, 4 blank, 5 header, 6 T00, 7 T01
    # -> delete T01, change S02's number, add a fake step, break T00's step no.
    ws.delete_rows(7)
    ws["C3"] = 301
    ws["A7"], ws["B7"], ws["C7"] = "S77", 77, 500
    ws["B6"] = 5  # T00 -> step no. should be 0
    wb.save(out)

    reports, _ = run_check(_cfg(out, *pdfs, [("3.1", "3", "")]))
    errors = "\n".join(_issues(reports, "ERROR"))
    assert "T01 is in the step-name column of PDF 1 but NOT in the Excel sheet" in errors
    assert "S02: number 301 from the Excel sheet was not confirmed" in errors
    assert "S02 has number 300, which is NOT in the Excel row of S02" in errors
    assert "S77 is in the Excel sheet but was not found in the PDF 1 section" in errors
    assert "Step no. column shows 5, expected 0" in errors


def test_missed_variant_number_is_caught(pdfs, tmp_path):
    out = tmp_path / "mapping.xlsx"
    translate(["--pdf1", str(pdfs[0]), "--pdf2", str(pdfs[1]), "-o", str(out),
               "--sequence", "3.1:3:"])
    wb = load_workbook(out)
    wb["Sequence 01"]["C2"] = "201 OR 202"  # S01_3 -> 203 dropped
    wb.save(out)
    reports, _ = run_check(_cfg(out, *pdfs, [("3.1", "3", "")]))
    assert any("S01_3 has number 203, which is NOT in the Excel row of S01" in e
               for e in _issues(reports, "ERROR"))


def test_page_range_too_short_is_caught(tmp_path, pdfs):
    p1 = tmp_path / "chart.pdf"
    make_sequence_chart_pdf(p1)
    out = tmp_path / "mapping.xlsx"
    translate(["--pdf1", str(p1), "--pdf2", str(pdfs[1]), "-o", str(out),
               "--sequence", "3.1:1:"])  # the section continues on page 2
    reports, _ = run_check(_cfg(out, p1, pdfs[1], [("3.1", "1", "")]))
    assert any("may continue on PDF 1 page 2" in e for e in _issues(reports, "ERROR"))


def test_step_at_page_bottom_confirmed(tmp_path, pdfs):
    p1 = tmp_path / "edges.pdf"
    make_edge_pdf(p1)
    out = tmp_path / "mapping.xlsx"
    translate(["--pdf1", str(p1), "--pdf2", str(pdfs[1]), "-o", str(out),
               "--sequence", "3.1:1-2:"])
    reports, _ = run_check(_cfg(out, p1, pdfs[1], [("3.1", "1-2", "")]))
    rep = reports[0]
    assert rep.excel_steps == 6 and rep.steps_confirmed == 6
    assert not any("step-name column of PDF 1" in e for e in _issues(reports, "ERROR"))


def test_missed_step_in_similar_sequences_is_caught(tmp_path, pdfs):
    """A step left out of the sheet is reported even when a similar sequence on a
    nearby page has the same name at the same place (it is not a page header)."""
    from test_translator import make_similar_sequences_pdf

    p1 = tmp_path / "similar.pdf"
    names, ranges = make_similar_sequences_pdf(p1, n_sequences=4)
    out = tmp_path / "mapping.xlsx"
    translate(["--pdf1", str(p1), "--pdf2", str(pdfs[1]), "-o", str(out),
               "--sequence", f"3.3:{ranges[3]}:"])
    wb = load_workbook(out)
    ws = wb["Sequence 01"]
    sheet_steps = [c.value for c in ws["A"]]
    assert set(names) <= set(sheet_steps)
    ws.delete_rows(sheet_steps.index("T00") + 1)  # T00 sits right below the heading
    wb.save(out)
    reports, _ = run_check(_cfg(out, p1, pdfs[1], [("3.3", ranges[3], "")]))
    errors = "\n".join(_issues(reports, "ERROR"))
    assert "T00 is in the step-name column of PDF 1 but NOT in the Excel sheet" in errors
