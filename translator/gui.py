"""Minimal Tkinter window so the tool can be used without the command line."""

from __future__ import annotations

import argparse
import threading


def launch(parser: argparse.ArgumentParser) -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    from .cli import run

    root = tk.Tk()
    root.title("PDF step -> number mapper")

    fields: dict[str, tk.StringVar] = {}

    def add_row(r, label, key, default="", browse=None):
        tk.Label(root, text=label, anchor="w").grid(row=r, column=0, sticky="w", padx=6, pady=3)
        var = tk.StringVar(value=default)
        tk.Entry(root, textvariable=var, width=55).grid(row=r, column=1, padx=6, pady=3)
        if browse:
            tk.Button(root, text="Browse...", command=lambda: browse(var)).grid(row=r, column=2, padx=6)
        fields[key] = var

    def open_pdf(var):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            var.set(path)

    def save_xlsx(var):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
        if path:
            var.set(path)

    add_row(0, "PDF 1 (step names)", "pdf1", browse=open_pdf)
    add_row(1, "Section of the table (e.g. 3.1)", "section")
    add_row(2, "PDF 2 (step numbers)", "pdf2", browse=open_pdf)
    add_row(3, "Output Excel file", "output", browse=save_xlsx)
    add_row(4, "Step column", "step_column", "A")
    add_row(5, "Number column", "number_column", "B")
    add_row(6, "Header row", "start_row", "1")
    add_row(7, "Step header", "step_header", "Step")
    add_row(8, "Number header", "number_header", "Number")
    add_row(9, "Name label in PDF 2 (optional, e.g. name:)", "name_label")
    add_row(10, "Pages column (optional, e.g. C)", "pages_column")

    log = scrolledtext.ScrolledText(root, width=90, height=15, state="disabled")
    log.grid(row=12, column=0, columnspan=3, padx=6, pady=6)

    def write_log(message: str):
        def append():
            log.configure(state="normal")
            log.insert("end", message + "\n")
            log.see("end")
            log.configure(state="disabled")
        root.after(0, append)

    def start():
        values = {k: v.get().strip() for k, v in fields.items()}
        for key in ("pdf1", "section", "pdf2", "output"):
            if not values[key]:
                messagebox.showerror("Missing input", f"Please fill in: {key}")
                return
        argv = ["--pdf1", values["pdf1"], "--section", values["section"],
                "--pdf2", values["pdf2"], "--output", values["output"],
                "--step-column", values["step_column"] or "A",
                "--number-column", values["number_column"] or "B",
                "--start-row", values["start_row"] or "1",
                "--step-header", values["step_header"],
                "--number-header", values["number_header"]]
        if values["name_label"]:
            argv += ["--name-label", values["name_label"]]
        if values["pages_column"]:
            argv += ["--pages-column", values["pages_column"]]
        args = parser.parse_args(argv)

        def work():
            try:
                run(args, log=write_log)
                root.after(0, lambda: messagebox.showinfo("Done", f"Created {args.output}"))
            except Exception as exc:
                message = str(exc)
                write_log(f"ERROR: {message}")
                root.after(0, lambda: messagebox.showerror("Error", message))

        threading.Thread(target=work, daemon=True).start()

    tk.Button(root, text="Create Excel", command=start).grid(row=11, column=1, pady=6)
    root.mainloop()
    return 0
