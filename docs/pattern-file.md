# Pattern File Reference

A pattern file is an `.xlsx` workbook that tells grepxcel what to look for and extract from a data file. Think of it as a schema: it describes the layout, field names, types, and validation rules.

---

## Rules for pattern files

- Every cell must contain **plain text**. Numbers, dates, booleans, and formulas are rejected.
- Cell values must not exceed **1 000 characters**.
- Formatting (fonts, colours, borders) is ignored and may be used freely.
- Only the **active sheet** is read; all other sheets are ignored.

---

## Overall structure

A pattern file has three areas, written top-to-bottom:

```
┌─────────────────────────────────────────┐
│  config: rows   (optional, 0 or more)   │
│  def: rows      (required, 1 or more)   │
│  START:                                 │
│    cell / table instructions            │
│  END:                                   │
└─────────────────────────────────────────┘
```

Each row uses **columns A, B, C, D, …** as fields. Column A is always the row-type keyword.

---

## config: rows

Optional. Placed before `START:`. Each `config:` row sets one global option.

| Column A  | Column B           | Column C         | Notes                        |
|-----------|--------------------|------------------|------------------------------|
| `config:` | `read.direction`   | `LR` or `TD`     | Default: `LR`                |
| `config:` | `currency.sign`    | e.g. `€` or `$`  | Default: `€`                 |
| `config:` | `empty.aliases`    | e.g. `N/A`       | Repeat the row for each alias|

**`read.direction`** controls how the data sheet is scanned for `cell:` instructions:

- `LR` — left-to-right, top-to-bottom (row by row). Most common.
- `TD` — top-to-bottom, left-to-right (column by column).

**`currency.sign`** is prepended to the string representation of currency values when matching their regex.

**`empty.aliases`** lists strings that should be treated as empty cells (e.g. `N/A`, `-`, `—`). Add one alias per row.

---

## def: rows

Required. Each `def:` row defines one field that can be referenced in the `START:` section.

| Column A | Column B      | Column C   | Column D              |
|----------|---------------|------------|-----------------------|
| `def:`   | `FieldName`   | type       | regex                 |

- **FieldName** — any plain-text identifier, e.g. `InvoiceNo`, `Amount`, `Date`. Must be unique.
- **type** — one of the values below.
- **regex** — a Python `re.fullmatch` pattern applied to the string representation of the cell value. Use `.*` to accept anything. Nested unbounded quantifiers (e.g. `(a+)+`) are rejected as unsafe.

### Supported types

| Type        | Accepts                                      | Regex target           |
|-------------|----------------------------------------------|------------------------|
| `string`    | Any text                                     | The cell text as-is    |
| `integer`   | Whole numbers (Excel integers or whole floats)| `str(int_value)`       |
| `currency`  | Any number (int or float)                    | `str(numeric_value)`   |
| `date`      | Excel date cells                             | n/a (type check only)  |
| `datetime`  | Excel datetime cells                         | n/a (type check only)  |
| `timestamp` | Same as `datetime`                           | n/a (type check only)  |

For `date`/`datetime`/`timestamp`, the regex column is ignored — only the Python type is checked.

---

## START: and END:

Mark the extraction sequence. Everything between `START:` and `END:` is processed in order, top to bottom.

```
START:
  … instructions …
END:
```

`END:` is required. Any rows after it are ignored.

---

## cell: instructions

Extract a single cell from the data file.

| Column A  | Column B    |
|-----------|-------------|
| `cell:1`  | `FieldName` |

- **Column A** — always `cell:1`. The `1` means one cell. (Only `1` is supported currently.)
- **Column B** — the field name (must exist in a `def:` row) **or** the special keyword `IGNORE`.

`IGNORE` tells the engine to consume the next non-empty cell without extracting its value. Useful for skipping headers or labels that appear in the data file but that you don't need.

The engine advances through the data sheet in `read.direction` order, skipping empty cells, and assigns each non-empty cell to the next `cell:1` instruction in sequence.

---

## table: instructions

Extract one or more instances of a repeating mini-table. The engine searches the data sheet greedily and collects every matching block.

| Column A   | Column B *(leave blank)* |
|------------|--------------------------|
| `table:*`  |                          |

The `*` means "zero or more instances". Immediately below the `table:*` row, add the template rows that describe the mini-table's layout. **Column A must be blank** for all template rows — that is how the parser knows they belong to the table.

### Template row types

Each template row occupies one physical row in the pattern file. Column B holds the row-type keyword; columns C onwards hold field names (one per column).

```
    B            C         D         E  …
┌──────────┬──────────┬──────────┬────────┐
│ HEADER:1 │ field1   │ field2   │ …      │  ← strict: missing = no match
│ SPLITTER:1│         │          │ …      │  ← all cells must be empty
│ DATA:*   │ field1   │ field2   │ …      │  ← repeats until empty row / footer
│ FOOTER:1 │ field1   │ field2   │ …      │  ← strict: missing = no match
└──────────┴──────────┴──────────┴────────┘
```

**HEADER rows** (`HEADER:1`) — matched strictly: if any non-`EMPTY` column is missing from the data, the entire mini-table candidate is rejected and the engine tries the next anchor.

**SPLITTER rows** (`SPLITTER:1`) — every column in the row must be empty in the data. Used to represent a blank separator row between sections.

**DATA rows** (`DATA:*`) — collected until the row is entirely empty **or** a FOOTER row is detected. Missing values in DATA rows produce a warning but do not reject the match.

**FOOTER rows** (`FOOTER:1`) — matched strictly, like HEADER.

Only one DATA template row is supported per table block. HEADER and FOOTER rows may be omitted.

### Column order

Columns C, D, E, … in a template row correspond to the first, second, third, … column of the mini-table as it appears in the data, starting from the anchor (the first non-empty cell found).

### Special field names in tables

| Name     | Meaning                                                                 |
|----------|-------------------------------------------------------------------------|
| `IGNORE` | Consume the cell but do not extract its value.                          |
| `EMPTY`  | Assert the cell is empty. If it is not, the mini-table match fails.     |

### Table-level config

Optionally override settings for one table block:

```
    B          C                  D
┌──────────┬──────────────────┬──────┐
│ config:  │ read.direction   │ TD   │
└──────────┴──────────────────┴──────┘
```

Place this row inside the table block (column A blank), before any template rows.

---

## Complete example

```
A          B                C          D
─────────────────────────────────────────────────────────
config:    read.direction   LR
config:    currency.sign    €
config:    empty.aliases    N/A

def:       InvoiceNo        string     INV-\d{4,8}
def:       IssueDate        date
def:       Vendor           string     .{2,100}
def:       Amount           currency   \d+(\.\d{1,2})?
def:       ProductCode      string     [A-Z]{2}\d{4}
def:       Qty              integer    \d+
def:       UnitPrice        currency   \d+(\.\d{1,2})?
def:       Total            currency   \d+(\.\d{1,2})?

START:
cell:1     InvoiceNo
cell:1     IssueDate
cell:1     Vendor
cell:1     Amount

table:*
           HEADER:1         ProductCode  Qty  UnitPrice
           DATA:*           ProductCode  Qty  UnitPrice  Total
           FOOTER:1         IGNORE       IGNORE  IGNORE  Total
END:
```

In this example:
- The engine first reads four individual cells in scan order (InvoiceNo, IssueDate, Vendor, Amount).
- Then it searches the sheet for every mini-table that starts with a HEADER row of three fields and ends with a FOOTER row.

---

## Validation and error behaviour

| Situation                                  | Result                        |
|--------------------------------------------|-------------------------------|
| Formula in any pattern cell                | Fatal error — processing stops |
| Non-string value (number, date, bool)      | Fatal error — processing stops |
| Cell value longer than 1 000 characters    | Fatal error — processing stops |
| Regex with nested quantifiers `(a+)+`      | Fatal error — processing stops |
| Invalid regex syntax                       | Fatal error — processing stops |
| `def:` field referenced in `START:` but missing | Fatal error             |
| Data cell fails type/regex validation      | Warning logged, value kept     |
| Data cell empty where a value was expected | Warning logged, `null` stored  |

Fatal errors stop processing immediately and return a partial (possibly empty) result. Warnings are collected and reported in the summary.
