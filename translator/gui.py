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
    add_row(2, "Pages of the table (e.g. 3-7)", "pages")
    add_row(3, "PDF 2 (step numbers)", "pdf2", browse=open_pdf)
    add_row(4, "Pages to search in PDF 2 (optional)", "pdf2_pages")
    add_row(5, "Output Excel file", "output", browse=save_xlsx)
    add_row(6, "Step column", "step_column", "A")
    add_row(7, "Step no. column (T10 -> 10)", "step_number_column", "B")
    add_row(8, "Number column", "number_column", "C")
    add_row(9, "Header row", "start_row", "1")
    add_row(10, "Step header", "step_header", "Step")
    add_row(11, "Step no. header", "step_number_header", "Step no.")
    add_row(12, "Number header", "number_header", "Number")
    add_row(13, "Name label in PDF 2 (optional, e.g. name:)", "name_label")
    add_row(14, "Pages column (optional, e.g. D)", "pages_column")
    sort_var = tk.BooleanVar(value=True)
    tk.Checkbutton(root, text="Sort steps alphabetically", variable=sort_var).grid(
        row=15, column=1, sticky="w", padx=6)
    tk.Label(root, text="PDF 1: fill in the section, the pages, or both.", fg="gray").grid(
        row=16, column=1, sticky="w", padx=6)

    log = scrolledtext.ScrolledText(root, width=90, height=15, state="disabled")
    log.grid(row=18, column=0, columnspan=3, padx=6, pady=6)

    def write_log(message: str):
        def append():
            log.configure(state="normal")
            log.insert("end", message + "\n")
            log.see("end")
            log.configure(state="disabled")
        root.after(0, append)

    def start():
        values = {k: v.get().strip() for k, v in fields.items()}
        for key, label in (("pdf1", "PDF 1"), ("pdf2", "PDF 2"), ("output", "Output Excel file")):
            if not values[key]:
                messagebox.showerror("Missing input", f"Please fill in: {label}")
                return
        if not values["section"] and not values["pages"]:
            messagebox.showerror("Missing input", "Please fill in the section or the pages of the table")
            return
        argv = ["--pdf1", values["pdf1"],
                "--pdf2", values["pdf2"], "--output", values["output"],
                "--step-column", values["step_column"] or "A",
                "--step-number-column", values["step_number_column"],
                "--number-column", values["number_column"] or "C",
                "--start-row", values["start_row"] or "1",
                "--step-header", values["step_header"],
                "--step-number-header", values["step_number_header"],
                "--number-header", values["number_header"]]
        if not sort_var.get():
            argv.append("--no-sort")
        if values["pdf2_pages"]:
            argv += ["--pdf2-pages", values["pdf2_pages"]]
        if values["section"]:
            argv += ["--section", values["section"]]
        if values["pages"]:
            argv += ["--pages", values["pages"]]
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

    tk.Button(root, text="Create Excel", command=start).grid(row=17, column=1, pady=6)
    root.mainloop()
    return 0
