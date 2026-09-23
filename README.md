# TRANSLATOR – PDF step name → step number mapper

Creates a new Excel file that maps each **step name** (from a table in PDF 1) to its
**step number(s)** (from the `number:` field in PDF 2).

| Step | Number            |
| ---- | ----------------- |
| T00  | 100               |
| S01  | 201 OR 202 OR 203 |

## How it works

1. **PDF 1 – step names.** The table is found by its section number (e.g. `3.1`).
   The section is located via the PDF bookmarks when there are some, otherwise by
   searching the pages for the heading (table-of-contents lines are ignored).
   The first table after the heading is read, including rows continued on the
   following pages. Only the **first column** is used; cells that do not look like
   a step name (header text such as "Step", empty cells) are skipped.
2. **PDF 2 – step numbers.** Every `number:` field is assigned to the closest step
   name written before it on the same page. Variants such as `S01_1`, `S01_2`, … are
   matched to `S01` (but `S010` is not).
3. **Excel.** One row per step. Several numbers go into the same cell separated by
   `OR`. Steps without a number are highlighted in yellow and listed in the console.

## Windows exe (no Python needed)

Every push builds `TRANSLATOR.exe` on GitHub:

1. Open the repository on GitHub → **Actions** → **Build Windows exe**.
2. Click the latest successful run and download **TRANSLATOR-windows** under *Artifacts*.
3. Unzip it and double-click `TRANSLATOR.exe`. Windows SmartScreen may warn that the
   app is unrecognised (it is not code-signed): click *More info* → *Run anyway*.

## Installation (Python)

Python 3.9+ is required.

```bash
pip install -r requirements.txt
```

## Usage

Graphical interface (file pickers, column / row settings):

```bash
python -m translator
```

Command line:

```bash
python -m translator --pdf1 spec.pdf --section 3.1 --pdf2 steps.pdf --output mapping.xlsx
```

### Options

| Option | Default | Meaning |
| ------ | ------- | ------- |
| `--step-column` | `A` | Excel column for the step names |
| `--number-column` | `B` | Excel column for the numbers |
| `--start-row` | `1` | Row of the header; data starts on the row below |
| `--step-header` / `--number-header` | `Step` / `Number` | Header texts |
| `--no-header` | | Do not write a header row |
| `--sheet` | `Mapping` | Sheet name |
| `--pages-column` | | Extra column listing the PDF 2 pages each number was read from (handy for checking) |
| `--separator` | `OR` | Word between multiple numbers |
| `--not-found` | empty | Text written when no number is found |
| `--number-label` | `number:` | Label of the number field in PDF 2 |
| `--name-label` | | Label of the step name field in PDF 2 (e.g. `name:`). Use it if the property pages also mention other steps (e.g. "previous: S01") so only the real name is used |
| `--step-pattern` | `^[A-Za-z]{1,10}\d+[A-Za-z0-9]*$` | Regex a first-column cell must match to count as a step |
| `--table-strategy` | `lines` | Use `text` when the table in PDF 1 has no ruling lines |

Example with a custom layout:

```bash
python -m translator --pdf1 spec.pdf --section 3.1 --pdf2 steps.pdf -o mapping.xlsx \
    --step-column C --number-column D --start-row 2 --name-label "name:" --pages-column E
```

## Development

```bash
pip install -r requirements-dev.txt
pytest
```

The tests generate sample PDFs on the fly (TOC, several sections, a table spanning
pages, step variants, `number:` value on the next line, look-alike names like `S010`).
