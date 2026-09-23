# TRANSLATOR – PDF step name → step number mapper

Creates a new Excel file that maps each **step name** (from a table in PDF 1) to its
**step number(s)** (from the `number:` field in PDF 2).

| Step | Step no. | Number            |
| ---- | -------- | ----------------- |
| S01  | 1        | 201 OR 202 OR 203 |
| T10  | 10       | 100               |

## How it works

1. **PDF 1 – step names.** The table is found by its section number (e.g. `3.1`),
   by its page numbers (e.g. `3-7`, as shown by the PDF viewer), or both.
   The section is located via the PDF bookmarks when there are some, otherwise by
   searching the pages for the heading (table-of-contents lines are ignored), and
   runs until the next heading that is not a sub-section. Only the **first column**
   of the tables in that part is used; cells that do not look like a step name
   (header text such as "Step", empty cells) are skipped, and `T00 Start` gives `T00`.
   The step names are first taken from their position: the leftmost column of
   step-like words (this also works for "tables" that are really framed blocks, one
   per step, with the name in a box on the left). If that gives nothing, tables with
   and without ruling lines are tried, and finally the first word of each text line.
   Page headers/footers are recognised because they repeat at the same place on the
   neighbouring pages, so a step name right at the top or bottom of a page is kept.
4. **Safety check.** Nothing is dropped silently: any other text found in the
   step-name column (e.g. a name split as `T 60`) and any step-like word skipped as a
   header/footer is listed in the log and on a red **Check** sheet in the Excel file,
   together with the steps that got no number. If the Check sheet is missing, there
   was nothing to verify. When no step is found the program stops and
   shows the first-column cells it did read, so the problem is easy to spot.
2. **PDF 2 – step numbers.** Every `number:` field is assigned to the closest step
   name written before it on the same page. Variants such as `S01_1`, `S01_2`, … are
   matched to `S01` (but `S010` is not).
3. **Excel.** One row per step, sorted alphabetically (S02 before S10). Column B
   holds the number inside the step name (T10 → 10). Several numbers go into the
   same cell separated by `OR`. Steps without a number are highlighted in yellow and listed in the console.

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

Graphical interface (file pickers, sequence table, column / row settings):

```bash
python -m translator
```

The window has a **sequence table** (17 rows by default, "Add row" for more): each row
is one table of PDF 1 and becomes one sheet (`Sequence 01`, `Sequence 02`, ...) of the
same Excel file. Fill in the PDF 1 pages and/or PDF 2 pages of a row to include it;
the section is pre-filled with `3.1`, `3.2`, ... and can be changed. Rows without page
numbers are skipped. All inputs are remembered for the next run
(`~/.translator_settings.json`). If one sequence fails, the other sheets are still
written; the failed one gets a red tab with the error, and is listed on the Check sheet.

Command line:

```bash
python -m translator --pdf1 spec.pdf --section 3.1 --pdf2 steps.pdf --output mapping.xlsx
python -m translator --pdf1 spec.pdf --pages 3-7 --pdf2 steps.pdf --output mapping.xlsx

# several sequences -> sheets "Sequence 01", "Sequence 02", ... (section:PDF 1 pages:PDF 2 pages)
python -m translator --pdf1 spec.pdf --pdf2 steps.pdf -o mapping.xlsx \
    --sequence "3.1:3-7:100-250" --sequence "3.2:8-9:251-300"
```

### Options

| Option | Default | Meaning |
| ------ | ------- | ------- |
| `--pdf2-pages` | all | PDF page numbers to search in PDF 2, e.g. `100-250` |
| `--step-column` | `A` | Excel column for the step names |
| `--step-number-column` | `B` | Excel column for the number inside the step name (T10 → 10); `""` leaves it out |
| `--number-column` | `C` | Excel column for the numbers from PDF 2 |
| `--no-sort` | | Keep the order of PDF 1 instead of sorting alphabetically |
| `--start-row` | `1` | Row of the header; data starts on the row below |
| `--step-header` / `--step-number-header` / `--number-header` | `Step` / `Step no.` / `Number` | Header texts |
| `--no-header` | | Do not write a header row |
| `--sheet` | `Mapping` | Sheet name |
| `--pages-column` | | Extra column listing the PDF 2 pages each number was read from (handy for checking) |
| `--separator` | `OR` | Word between multiple numbers |
| `--not-found` | empty | Text written when no number is found |
| `--number-label` | `number:` | Label of the number field in PDF 2 |
| `--name-label` | | Label of the step name field in PDF 2 (e.g. `name:`). Use it if the property pages also mention other steps (e.g. "previous: S01") so only the real name is used |
| `--step-pattern` | `^[A-Za-z]{1,10}\d+[A-Za-z0-9]*$` | Regex a first-column cell must match to count as a step |
| `--pages` | | PDF page numbers of the table in PDF 1, e.g. `3-7` or `3,5,8-9` |
| `--table-strategy` | `auto` | How step names in PDF 1 are read: `position` (leftmost column), `lines` (table with ruling lines), `text` (table by alignment) or `auto` (in that order) |

Example with a custom layout:

```bash
python -m translator --pdf1 spec.pdf --section 3.1 --pdf2 steps.pdf -o mapping.xlsx \
    --pdf2-pages 100-250 --start-row 2 --name-label "name:" --pages-column D
```

## Development

```bash
pip install -r requirements-dev.txt
pytest
```

The tests generate sample PDFs on the fly (TOC, several sections, a table spanning
pages, step variants, `number:` value on the next line, look-alike names like `S010`).
