"""Tkinter window so the tool can be used without the command line."""

from __future__ import annotations

import argparse
import json
import threading
from pathlib import Path

SETTINGS_FILE = Path.home() / ".translator_settings.json"
DEFAULT_SEQUENCES = 17


def _load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_settings(data: dict) -> None:
    try:
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass  # remembering the inputs is a convenience only


def launch(parser: argparse.ArgumentParser) -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    from .cli import Job, run

    saved = _load_settings()
    root = tk.Tk()
    root.title("PDF step -> number mapper")

    # ---------------------------------------------------------------- files
    files = tk.LabelFrame(root, text="Files", padx=6, pady=4)
    files.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=4)
    file_vars: dict[str, tk.StringVar] = {}

    def browse_pdf(var):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            var.set(path)

    def browse_xlsx(var):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
        if path:
            var.set(path)

    for r, (key, label, browse) in enumerate([
        ("pdf1", "PDF 1 (step names)", browse_pdf),
        ("pdf2", "PDF 2 (step numbers)", browse_pdf),
        ("output", "Output Excel file", browse_xlsx),
    ]):
        tk.Label(files, text=label, anchor="w").grid(row=r, column=0, sticky="w")
        var = tk.StringVar(value=saved.get(key, ""))
        tk.Entry(files, textvariable=var, width=80).grid(row=r, column=1, padx=4, pady=2)
        tk.Button(files, text="Browse...", command=lambda v=var, b=browse: b(v)).grid(row=r, column=2)
        file_vars[key] = var

    # ------------------------------------------------------------ sequences
    seq_frame = tk.LabelFrame(root, text="Sequences (one Excel sheet each; empty rows are skipped)",
                              padx=6, pady=4)
    seq_frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
    canvas = tk.Canvas(seq_frame, height=330, width=470, highlightthickness=0)
    scrollbar = tk.Scrollbar(seq_frame, orient="vertical", command=canvas.yview)
    table = tk.Frame(canvas)
    table.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=table, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

    for c, text in enumerate(["Sheet", "Section", "PDF 1 pages", "PDF 2 pages"]):
        tk.Label(table, text=text, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=c, padx=3)

    rows: list[dict[str, tk.StringVar]] = []

    def add_sequence_row(values: dict | None = None):
        n = len(rows) + 1
        values = values or {"section": f"3.{n}", "pages": "", "pdf2_pages": ""}
        tk.Label(table, text=f"Sequence {n:02d}").grid(row=n, column=0, padx=3, sticky="w")
        row_vars = {}
        for c, (key, width) in enumerate([("section", 8), ("pages", 14), ("pdf2_pages", 14)],
                                         start=1):
            var = tk.StringVar(value=values.get(key, ""))
            tk.Entry(table, textvariable=var, width=width).grid(row=n, column=c, padx=3, pady=1)
            row_vars[key] = var
        rows.append(row_vars)

    saved_rows = saved.get("sequences") or []
    for i in range(max(DEFAULT_SEQUENCES, len(saved_rows))):
        add_sequence_row(saved_rows[i] if i < len(saved_rows) else None)

    def clear_pages():
        if messagebox.askyesno("Clear", "Clear all page numbers in the sequence table?"):
            for row_vars in rows:
                row_vars["pages"].set("")
                row_vars["pdf2_pages"].set("")

    buttons = tk.Frame(seq_frame)
    buttons.grid(row=1, column=0, columnspan=2, sticky="w", pady=4)
    tk.Button(buttons, text="Add row", command=add_sequence_row).pack(side="left")
    tk.Button(buttons, text="Clear page numbers", command=clear_pages).pack(side="left", padx=6)

    # -------------------------------------------------------------- options
    opts = tk.LabelFrame(root, text="Excel layout and options", padx=6, pady=4)
    opts.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)
    opt_vars: dict[str, tk.StringVar] = {}
    saved_opts = saved.get("options", {})
    for r, (key, label, default) in enumerate([
        ("step_column", "Step column", "A"),
        ("step_number_column", "Step no. column (T10 -> 10)", "B"),
        ("number_column", "Number column", "C"),
        ("start_row", "Header row", "1"),
        ("step_header", "Step header", "Step"),
        ("step_number_header", "Step no. header", "Step no."),
        ("number_header", "Number header", "Number"),
        ("name_label", "Name label in PDF 2 (optional)", ""),
        ("pages_column", "Pages column (optional, e.g. D)", ""),
    ]):
        tk.Label(opts, text=label, anchor="w").grid(row=r, column=0, sticky="w")
        var = tk.StringVar(value=saved_opts.get(key, default))
        tk.Entry(opts, textvariable=var, width=14).grid(row=r, column=1, padx=4, pady=1)
        opt_vars[key] = var
    sort_var = tk.BooleanVar(value=saved_opts.get("sort", True))
    tk.Checkbutton(opts, text="Sort steps alphabetically", variable=sort_var).grid(
        row=20, column=0, columnspan=2, sticky="w")

    # ------------------------------------------------------------------ log
    log = scrolledtext.ScrolledText(root, width=110, height=12, state="disabled")
    log.grid(row=3, column=0, columnspan=2, padx=6, pady=6)

    def write_log(message: str):
        def append():
            log.configure(state="normal")
            log.insert("end", message + "\n")
            log.see("end")
            log.configure(state="disabled")
        root.after(0, append)

    def collect() -> dict:
        return {
            **{k: v.get().strip() for k, v in file_vars.items()},
            "sequences": [{k: v.get().strip() for k, v in r.items()} for r in rows],
            "options": {**{k: v.get().strip() for k, v in opt_vars.items()},
                        "sort": sort_var.get()},
        }

    def start():
        data = collect()
        _save_settings(data)
        for key, label in (("pdf1", "PDF 1"), ("pdf2", "PDF 2"), ("output", "Output Excel file")):
            if not data[key]:
                messagebox.showerror("Missing input", f"Please fill in: {label}")
                return

        # A row counts when it has PDF 1 pages or PDF 2 pages; the section alone is
        # pre-filled and does not make a row active.
        jobs = [Job(f"Sequence {i:02d}", seq["section"] or None, seq["pages"] or None,
                    seq["pdf2_pages"] or None)
                for i, seq in enumerate(data["sequences"], start=1)
                if seq["pages"] or seq["pdf2_pages"]]
        if not jobs:
            messagebox.showerror("Missing input",
                                 "Please fill in the page numbers of at least one sequence")
            return

        o = data["options"]
        argv = ["--pdf1", data["pdf1"], "--pdf2", data["pdf2"], "--output", data["output"],
                "--step-column", o["step_column"] or "A",
                "--step-number-column", o["step_number_column"],
                "--number-column", o["number_column"] or "C",
                "--start-row", o["start_row"] or "1",
                "--step-header", o["step_header"],
                "--step-number-header", o["step_number_header"],
                "--number-header", o["number_header"]]
        if not o["sort"]:
            argv.append("--no-sort")
        if o["name_label"]:
            argv += ["--name-label", o["name_label"]]
        if o["pages_column"]:
            argv += ["--pages-column", o["pages_column"]]
        args = parser.parse_args(argv)
        create.configure(state="disabled")

        def work():
            try:
                failed = run(args, log=write_log, jobs=jobs)
                if failed:
                    root.after(0, lambda: messagebox.showwarning(
                        "Done with errors",
                        f"Created {args.output}\n\n{failed} sequence(s) failed: see the red "
                        "sheet tabs and the 'Check' sheet."))
                else:
                    root.after(0, lambda: messagebox.showinfo(
                        "Done", f"Created {args.output} with {len(jobs)} sheet(s)."))
            except Exception as exc:
                message = str(exc)
                write_log(f"ERROR: {message}")
                root.after(0, lambda: messagebox.showerror("Error", message))
            finally:
                root.after(0, lambda: create.configure(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    create = tk.Button(root, text="Create Excel", command=start, width=20)
    create.grid(row=2, column=0, columnspan=2, pady=4)
    root.protocol("WM_DELETE_WINDOW", lambda: (_save_settings(collect()), root.destroy()))
    root.mainloop()
    return 0
