"""Independent checker for the Excel files created by TRANSLATOR.

This package deliberately shares no code with `translator`: it reads the PDFs with
a different engine (PyMuPDF instead of pdfplumber) and uses its own rules, so a
mistake in one program is unlikely to be repeated by the other.
"""

__version__ = "0.1.0"
