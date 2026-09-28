"""Run the check from the command line or from a small window.

    python -m checker --excel mapping.xlsx --pdf1 spec.pdf --pdf2 steps.pdf \\
        --sequence "3.1:3-7:100-250" --sequence "3.2:8-9:251-300"

Without arguments a window opens; it is pre-filled from TRANSLATOR's saved settings.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
from pathlib import Path

from .excel_io import read_sheet, write_report
from .pdf_check import ERROR, WARNING, Pdf, SequenceReport, check_pdf1, check_pdf2

TRANSLATOR_SETTINGS = Path.home() / ".translator_settings.json"
CHECKER_SETTINGS = Path.home() / ".translator_checker_settings.json"


def report_path(excel: str) -> str:
    p = Path(excel)
    return str(p.with_name(p.stem + "_CHECK_REPORT.xlsx"))


def run_check(cfg: dict, log=print) -> tuple[list[SequenceReport], str]:
    """cfg: excel, pdf1, pdf2, sequences [{sheet, section, pages, pdf2_pages}], options."""
    o = cfg.get("options", {})
    step_col = (o.get("step_column") or "A").upper()
    step_no_col = (o.get("step_number_column", "B") or "").upper() or None  # "" = no such column
    number_col = (o.get("number_column") or "C").upper()
    header_row = int(o.get("start_row") or 1)
    label = o.get("number_label") or "number:"

    sequences = [s for s in cfg["sequences"] if s.get("pages") or s.get("pdf2_pages")]
    if not sequences:
        raise ValueError("No sequence has page numbers filled in")
    all_sections = {(s.get("section") or "").strip().rstrip(".") for s in cfg["sequences"]}
    all_sections.discard("")

    pdf1, pdf2 = Pdf(cfg["pdf1"]), Pdf(cfg["pdf2"])
    reports: list[SequenceReport] = []
    try:
        for seq in sequences:
            name = seq["sheet"]
            rep = SequenceReport(name)
            reports.append(rep)
            log(f"[{name}] checking ...")
            sheet = read_sheet(cfg["excel"], name, step_col, step_no_col, number_col, header_row)
            if sheet.error:
                rep.add(ERROR, f"The Excel sheet reports an error: {sheet.error}")
            rep.excel_steps = len(sheet.rows)
            steps = [r[0] for r in sheet.rows]

            if step_no_col:
                for step, step_no, _ in sheet.rows:
                    digits = re.search(r"\d+", step)
                    expected = int(digits.group()) if digits else None
                    try:
                        actual = int(step_no) if step_no is not None else None
                    except (TypeError, ValueError):
                        actual = step_no
                    if actual != expected:
                        rep.add(ERROR, f"Step no. column shows {step_no!r}, expected {expected}.",
                                step)
            try:
                others = all_sections - {(seq.get("section") or "").strip().rstrip(".")}
                check_pdf1(pdf1, name, seq.get("section"), seq.get("pages"), others, steps, rep)
                claimed: dict[str, list[str]] = {}
                for step, _, numbers in sheet.rows:
                    claimed.setdefault(step, []).extend(numbers)
                check_pdf2(pdf2, name, seq.get("pdf2_pages"), label, claimed, rep)
            except Exception as exc:
                rep.add(ERROR, f"Check could not run: {exc}")
            status = "OK" if not rep.errors and not rep.warnings else \
                f"{rep.errors} error(s), {rep.warnings} warning(s)"
            log(f"  steps {rep.steps_confirmed}/{rep.excel_steps} confirmed in PDF 1, numbers "
                f"{rep.numbers_confirmed}/{rep.numbers_claimed} confirmed in PDF 2 -> {status}")
            for issue in rep.issues:
                if issue.severity in (ERROR, WARNING):
                    page = f" (page {issue.page})" if issue.page else ""
                    log(f"    {issue.severity}: {issue.message}{page}")
    finally:
        pdf1.close()
        pdf2.close()

    out = report_path(cfg["excel"])
    write_report(out, reports, {"Checked file": cfg["excel"], "PDF 1": cfg["pdf1"],
                                "PDF 2": cfg["pdf2"]})
    errors = sum(r.errors for r in reports)
    warnings = sum(r.warnings for r in reports)
    if errors or warnings:
        log(f"RESULT: {errors} error(s) and {warnings} warning(s) - see {out}")
    else:
        log(f"RESULT: ALL CHECKS PASSED ({len(reports)} sequence(s)). Report: {out}")
    return reports, out


# ------------------------------------------------------------------------ CLI

def _load(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="checker", description="Independent check of a TRANSLATOR "
                                                             "Excel file against the two PDFs.")
    p.add_argument("--excel")
    p.add_argument("--pdf1")
    p.add_argument("--pdf2")
    p.add_argument("--sequence", action="append", metavar="SECTION:PAGES1:PAGES2",
                   help='as in TRANSLATOR; sheets are "Sequence 01", "Sequence 02", ...')
    p.add_argument("--step-column", default="A")
    p.add_argument("--step-number-column", default="B")
    p.add_argument("--number-column", default="C")
    p.add_argument("--start-row", type=int, default=1)
    p.add_argument("--number-label", default="number:")
    args = p.parse_args(argv)

    if not (args.excel or args.pdf1 or args.pdf2 or args.sequence):
        return launch_gui()
    if not (args.excel and args.pdf1 and args.pdf2 and args.sequence):
        p.error("--excel, --pdf1, --pdf2 and at least one --sequence are required")
    seqs = []
    for i, text in enumerate(args.sequence, start=1):
        parts = (text.split(":") + ["", ""])[:3]
        seqs.append({"sheet": f"Sequence {i:02d}", "section": parts[0], "pages": parts[1],
                     "pdf2_pages": parts[2]})
    cfg = {"excel": args.excel, "pdf1": args.pdf1, "pdf2": args.pdf2, "sequences": seqs,
           "options": {"step_column": args.step_column,
                       "step_number_column": args.step_number_column,
                       "number_column": args.number_column, "start_row": args.start_row,
                       "number_label": args.number_label}}
    try:
        reports, _ = run_check(cfg)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    return 1 if any(r.errors or r.warnings for r in reports) else 0


# ------------------------------------------------------------------------ GUI

def launch_gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    saved = _load(CHECKER_SETTINGS) or {}
    tr = _load(TRANSLATOR_SETTINGS)

    root = tk.Tk()
    root.title("CHECKER - independent check of the step/number Excel file")

    files = tk.LabelFrame(root, text="Files", padx=6, pady=4)
    files.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=4)
    fvars: dict[str, tk.StringVar] = {}
    for r, (key, label, default) in enumerate([
        ("excel", "Excel file to check", tr.get("output", "")),
        ("pdf1", "PDF 1 (step names)", tr.get("pdf1", "")),
        ("pdf2", "PDF 2 (step numbers)", tr.get("pdf2", "")),
    ]):
        tk.Label(files, text=label, anchor="w").grid(row=r, column=0, sticky="w")
        var = tk.StringVar(value=saved.get(key) or default)
        tk.Entry(files, textvariable=var, width=80).grid(row=r, column=1, padx=4, pady=2)
        types = [("Excel files", "*.xlsx")] if key == "excel" else [("PDF files", "*.pdf")]

        def browse(v=var, t=types):
            path = filedialog.askopenfilename(filetypes=t)
            if path:
                v.set(path)
        tk.Button(files, text="Browse...", command=browse).grid(row=r, column=2)
        fvars[key] = var

    seq_frame = tk.LabelFrame(root, text="Sequences (as in TRANSLATOR; empty rows are skipped)",
                              padx=6, pady=4)
    seq_frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
    canvas = tk.Canvas(seq_frame, height=300, width=470, highlightthickness=0)
    bar = tk.Scrollbar(seq_frame, orient="vertical", command=canvas.yview)
    table = tk.Frame(canvas)
    table.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=table, anchor="nw")
    canvas.configure(yscrollcommand=bar.set)
    canvas.grid(row=0, column=0)
    bar.grid(row=0, column=1, sticky="ns")
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
    for c, text in enumerate(["Sheet", "Section", "PDF 1 pages", "PDF 2 pages"]):
        tk.Label(table, text=text, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=c)

    rows: list[dict[str, tk.StringVar]] = []

    def fill_rows(seqs: list[dict]):
        for widget in list(table.grid_slaves()):
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        rows.clear()
        for i in range(max(17, len(seqs))):
            s = seqs[i] if i < len(seqs) else {"section": f"3.{i + 1}"}
            tk.Label(table, text=f"Sequence {i + 1:02d}").grid(row=i + 1, column=0, sticky="w")
            rv = {}
            for c, (key, width) in enumerate([("section", 8), ("pages", 14), ("pdf2_pages", 14)],
                                             start=1):
                var = tk.StringVar(value=s.get(key, ""))
                tk.Entry(table, textvariable=var, width=width).grid(row=i + 1, column=c, padx=3)
                rv[key] = var
            rows.append(rv)

    fill_rows(saved.get("sequences") or tr.get("sequences") or [])

    def copy_from_translator():
        t = _load(TRANSLATOR_SETTINGS)
        if not t:
            messagebox.showinfo("Not found", "No saved TRANSLATOR settings found.")
            return
        fill_rows(t.get("sequences") or [])
        for key, tkey in (("excel", "output"), ("pdf1", "pdf1"), ("pdf2", "pdf2")):
            if t.get(tkey):
                fvars[key].set(t[tkey])
        for key in ovars:
            if key in (t.get("options") or {}):
                ovars[key].set(t["options"][key])

    tk.Button(seq_frame, text="Copy everything from TRANSLATOR", command=copy_from_translator
              ).grid(row=1, column=0, sticky="w", pady=4)

    opts = tk.LabelFrame(root, text="Excel layout", padx=6, pady=4)
    opts.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)
    ovars: dict[str, tk.StringVar] = {}
    saved_o = saved.get("options") or tr.get("options") or {}
    for r, (key, label, default) in enumerate([
        ("step_column", "Step column", "A"),
        ("step_number_column", "Step no. column", "B"),
        ("number_column", "Number column", "C"),
        ("start_row", "Header row", "1"),
        ("number_label", "Number label in PDF 2", "number:"),
    ]):
        tk.Label(opts, text=label, anchor="w").grid(row=r, column=0, sticky="w")
        var = tk.StringVar(value=str(saved_o.get(key, default)))
        tk.Entry(opts, textvariable=var, width=12).grid(row=r, column=1, padx=4, pady=1)
        ovars[key] = var

    log = scrolledtext.ScrolledText(root, width=110, height=14, state="disabled")
    log.grid(row=3, column=0, columnspan=2, padx=6, pady=6)

    def write_log(message: str):
        def append():
            log.configure(state="normal")
            log.insert("end", message + "\n")
            log.see("end")
            log.configure(state="disabled")
        root.after(0, append)

    def collect() -> dict:
        return {**{k: v.get().strip() for k, v in fvars.items()},
                "sequences": [{"sheet": f"Sequence {i:02d}",
                               **{k: v.get().strip() for k, v in r.items()}}
                              for i, r in enumerate(rows, start=1)],
                "options": {k: v.get().strip() for k, v in ovars.items()}}

    def save(cfg):
        try:
            CHECKER_SETTINGS.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        except Exception:
            pass

    def start():
        cfg = collect()
        save(cfg)
        for key, label in (("excel", "Excel file"), ("pdf1", "PDF 1"), ("pdf2", "PDF 2")):
            if not cfg[key]:
                messagebox.showerror("Missing input", f"Please fill in: {label}")
                return
        button.configure(state="disabled")

        def work():
            try:
                reports, out = run_check(cfg, log=write_log)
                errors = sum(r.errors for r in reports)
                warnings = sum(r.warnings for r in reports)
                if errors or warnings:
                    root.after(0, lambda: messagebox.showwarning(
                        "Check finished", f"{errors} error(s), {warnings} warning(s).\n\n"
                        f"Details: {out}"))
                else:
                    root.after(0, lambda: messagebox.showinfo(
                        "Check finished", f"ALL CHECKS PASSED ({len(reports)} sequence(s)).\n\n"
                        f"Report: {out}"))
            except Exception as exc:
                msg = str(exc)
                write_log(f"ERROR: {msg}")
                root.after(0, lambda: messagebox.showerror("Error", msg))
            finally:
                root.after(0, lambda: button.configure(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    button = tk.Button(root, text="Run check", command=start, width=20)
    button.grid(row=2, column=0, columnspan=2, pady=4)
    root.protocol("WM_DELETE_WINDOW", lambda: (save(collect()), root.destroy()))
    root.mainloop()
    return 0
