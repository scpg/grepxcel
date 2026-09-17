# grepxcel Web Wizard — Guide

The web wizard is the recommended way to build a pattern file for non-developers
and for anyone who prefers a mouse-driven interface over a terminal workflow.

It opens your Excel file in the browser as a live grid.  You click cells to
classify them, then download the finished pattern with one button.

---

## Requirements

The web extras must be installed:

```bash
pip install "grepxcel[web]"
```

This adds FastAPI and uvicorn.  The core extraction engine is unaffected.

---

## Starting the wizard

```bash
grepxcel web-wizard data.xlsx
```

The browser opens automatically at `http://localhost:8765`.  To choose a
different port or suppress the auto-open:

```bash
grepxcel web-wizard data.xlsx --port 9000 --no-browser
```

To pre-populate the wizard with an existing pattern (useful for editing):

```bash
grepxcel web-wizard data.xlsx -p existing-pattern.xlsx
```

---

## The interface

The wizard has four tabs on the right panel:

| Tab | Purpose |
|---|---|
| **Classify** | The main workspace — click cells to classify them |
| **Extract** | Run a live extraction and preview the JSON output |
| **Config** | Set options (read direction, currency symbol, case handling) |
| **Logs** | Session event log for debugging |

The bottom status bar shows how many cells are classified by type:
**Labels · Variables · Headers · Mini Tables · Ignored · Unclassified**

---

## Classifying cells

Click any cell to select it.  The right panel shows the cell's value and
inferred type.  Then press a key (or click the corresponding button) to classify:

| Key | Type | What it means |
|---|---|---|
| `L` | **Label** | Static text the engine uses to locate itself on the sheet (e.g. `Invoice No:`). Never in the output. Enter a name when prompted. |
| `V` | **Value** | A data cell you want to extract. Enter the output field name (dot notation for nesting: `invoice.number`). |
| `C` | **Column header** | A decorative section title — consumed but not extracted. |
| `T` | **Mini-table** | Marks the top-left corner of a repeating table block. Starts a three-step table definition flow. |
| `I` | **Ignore** | Cell is consumed but skipped. |
| `R` | **Remove** | Clear the classification on this cell. |

### Keyboard shortcuts

| Key | Action |
|---|---|
| Arrow keys | Navigate the grid |
| `Ctrl+Z` | Undo the last classification |
| `Ctrl+G` | Jump to a specific cell (e.g. `C7`) |

---

## Saving the pattern

When you have classified all the cells you need, click:

- **Save CSV** — saves a `.csv` pattern next to the data file (use this for
  pipelines and version control; CSV is plain text and diff-friendly).
- **Save Excel** — saves a coloured `.xlsx` pattern next to the data file
  (easier to read in a spreadsheet app; colour-coded by cell role).

The saved filename is `<data-stem>_pattern-from-web.csv` (or `.xlsx`).

---

## Running the extraction

After saving the pattern, run:

```bash
grepxcel extract -p <stem>_pattern-from-web.csv data.xlsx
```

Or click the **Extract** tab in the wizard to preview the JSON output
directly in the browser without leaving the tool.

---

## The pattern concept

A pattern file is just a spreadsheet (CSV or xlsx) that describes the layout
of your data file.  It has two kinds of rows:

- **Definition rows** (`lbl:`, `var:`) — declare labels and output variables.
- **Instruction rows** (`START:` / `cell:1` / `table:*` / `END:`) — tell the
  engine in what order to find each field.

You do not need to understand this structure to use the web wizard — the
wizard builds it for you.  But if you want to edit the pattern directly,
see [pattern-file.md](pattern-file.md).

---

## Common match patterns (regex)

The **Match pattern** field in each variable modal accepts a regex.  Common
values:

| Pattern | Matches |
|---|---|
| `.*` | Anything (default — accept any value) |
| `\d+` | Integers only (e.g. `42`, `1000`) |
| `[\d.,]+` | Decimal / currency numbers (e.g. `1,234.56`) |
| `\S+@\S+\.\S+` | Email addresses |
| `https?://\S+` | URLs |
| `\d{4}-\d{2}-\d{2}` | ISO dates (`YYYY-MM-DD`) |
| `[A-Z]{2,3}-\d+` | Reference codes (e.g. `PO-1234`, `INV-007`) |

Leave it as `.*` if you don't need to validate the value format.

---

## Merged cells

Merged cells in Excel are read from the **top-left cell** of the merge range.
The remaining cells of the merge are invisible to the wizard and engine.
If an anchor label sits on a merged cell, classify the top-left cell as normal.

---

## Troubleshooting

**The wizard shows no data** — check that the file path is correct and the
sheet has data.  Run `grepxcel lint data.xlsx` for a quick diagnosis.

**A pattern I saved doesn't extract correctly** — open the **Extract** tab in
the wizard to see what the engine found.  Run `grepxcel extract -p pattern.csv data.xlsx -v`
for a cell-by-cell trace showing which anchor fired and what value was read.

**The browser tab shows an error** — check the terminal where you started the
wizard for the Python traceback.

---

## P3 roadmap (planned enhancements)

- **Regex preset menu** — a small dropdown of common patterns (integer, email,
  URL, date) so non-technical users don't need to know regex syntax.
- **Zero-assumption install guide** — step-by-step instructions for Windows
  users who have never used a terminal, covering Python installation through
  first extraction.

See [CONTRIBUTING.md](../CONTRIBUTING.md) to help prioritise or implement these.
