# Pattern File Reference

A pattern file is an `.xlsx` workbook that tells grepxcel what to look for and extract from a data file. Think of it as a schema: it describes the layout, field names, types, and validation rules.

---

## Rules for pattern files

- Every cell must contain **plain text**. Numbers, dates, booleans, and formulas are rejected.
- Cell values must not exceed **1 000 characters**.
- Formatting (fonts, colours, borders) is ignored and may be used freely.
- By default only the **active sheet** of the *data* file is read. Use `--sheet`
  to target another sheet, or `--all-sheets` to process every sheet with the
  same pattern (output is then keyed by sheet name).

---

## Overall structure

A pattern file has three areas, written top-to-bottom:

```
┌──────────────────────────────────────────────┐
│  config: rows        (optional, 0 or more)   │
│  lbl: / var: / doc:  (definitions, 1 or more)│
│  START:                                      │
│    cell / table instructions                 │
│  END:                                        │
└──────────────────────────────────────────────┘
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

## Field definition rows: `lbl:`, `var:`, `doc:`

At least one definition row is required. Each defines one field that can be
referenced in the `START:` section, or documents the pattern.

| Column A | Column B      | Column C   | Column D              | Role |
|----------|---------------|------------|-----------------------|------|
| `lbl:`   | `FieldName`   | type       | regex                 | **Anchor** — matched for position only; **never written to output JSON** |
| `var:`   | `field.name`  | type       | regex                 | **Variable** — extracted and written to output JSON |
| `doc:`   | (free text)   |            |                       | **Comment** — ignored by the engine |
| `def:`   | `FieldName`   | type       | regex                 | Backward-compatible alias for `var:` |

- **`lbl:`** — use for literal text that marks *where* a value lives: labels like
  `Invoice No:` or column headers like `Product`, `Qty`. Matched but stripped
  from output.
- **`var:`** — use for every value you want to capture. **Dot notation creates
  nested JSON**: `po.number` → `{"po": {"number": …}}`. In a table, the group
  prefix (`line` in `line.qty`) becomes the output array key.
- **FieldName** — a unique plain-text identifier.
- **regex** — a Python `re.fullmatch` pattern applied to the string
  representation of the cell value. Use `.*` to accept anything. Nested
  unbounded quantifiers (e.g. `(a+)+`) are rejected as unsafe (ReDoS guard).

### Supported types

| Type         | Accepts                                       | Regex target           |
|--------------|-----------------------------------------------|------------------------|
| `string`     | Any text                                      | The cell text as-is    |
| `integer`    | Whole numbers (Excel integers or whole floats)| `str(int_value)`       |
| `currency`   | Any number (int or float)                     | `str(numeric_value)`   |
| `percentage` | Any number (int or float). Excel stores a percentage as a fraction, e.g. 62.5% → `0.625` | `str(numeric_value)` |
| `date`       | Excel date cells                              | n/a (type check only)  |
| `datetime`   | Excel datetime cells                          | n/a (type check only)  |
| `timestamp`  | Same as `datetime`                            | n/a (type check only)  |

`percentage` validates identically to `currency` (both require a numeric cell);
it exists to document intent — a reader sees that the field holds a percentage.
For `date`/`datetime`/`timestamp`, the regex column is ignored.

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

Extract a single cell from the data file. There are two addressing modes.

| Column A   | Column B    | Mode |
|------------|-------------|------|
| `cell:next`| `FieldName` | **Sequential** — the next non-empty cell in scan order (`cell:1` is an alias) |
| `cell:B5`  | `FieldName` | **Absolute** — jump directly to cell B5 (A1-notation) |

- **Column B** — a defined `var:`/`lbl:` field name, **or** the keyword `IGNORE`.

**Sequential (`cell:next` / `cell:1`)** — the engine advances through the data
sheet in `read.direction` order, skipping empty and already-consumed cells, and
assigns each non-empty cell to the next instruction in sequence.

**Absolute (`cell:B5`)** — the cursor jumps directly to the named coordinate and
the cursor advances past it. This is self-documenting: you can read the pattern
without mentally tracing the scan order. Rules:

- Absolute references must appear in forward reading order relative to each
  other; an out-of-order reference is rejected at parse time.
- Referencing a cell the cursor has already passed is a fatal error.
- On an empty target: `lbl:` → fatal (the anchor was expected); `var:` →
  records `null` and continues; `IGNORE` → skipped silently.

You can freely mix the two modes — e.g. jump with `cell:A1`, then read the rest
with `cell:next`.

`IGNORE` consumes a cell without extracting its value — useful for skipping
labels you don't need.

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

lbl:       inv_label        string     Invoice No:
lbl:       col_code         string     Code
lbl:       col_qty          string     Qty
lbl:       col_price        string     Unit Price
var:       inv.number       string     INV-\d{4,8}
var:       inv.date         date
var:       inv.vendor       string     .{2,100}
var:       line.code        string     [A-Z]{2}\d{4}
var:       line.qty         integer    \d+
var:       line.price       currency   \d+(\.\d{1,2})?
var:       line.total       currency   \d+(\.\d{1,2})?
var:       footer.total     currency   \d+(\.\d{1,2})?

START:
cell:next  inv_label
cell:next  inv.number
cell:next  inv.date
cell:next  inv.vendor

table:*
           HEADER:1         col_code   col_qty  col_price
           DATA:*           line.code  line.qty  line.price  line.total
           FOOTER:1         IGNORE     IGNORE   IGNORE      footer.total
END:
```

In this example:
- The engine reads the `Invoice No:` label (an `lbl:` anchor, dropped from output),
  then the invoice number, date and vendor (`var:` fields, kept in output).
- Then it searches the sheet for every mini-table whose HEADER matches the three
  column labels, collects the `line.*` DATA rows into a `"line"` array, and
  captures the footer total.
- `lbl:` fields never appear in the output; `var:` fields do, nested by their
  dot-notation prefix.

---

## Validation and error behaviour

| Situation                                  | Result                        |
|--------------------------------------------|-------------------------------|
| Formula in any pattern cell                | Fatal error — processing stops |
| Non-string value (number, date, bool)      | Fatal error — processing stops |
| Cell value longer than 1 000 characters    | Fatal error — processing stops |
| Regex with nested quantifiers `(a+)+`      | Fatal error — processing stops |
| Invalid regex syntax                       | Fatal error — processing stops |
| Field referenced in `START:` but not defined | Fatal error                  |
| Absolute `cell:` references out of reading order | Fatal error (at parse time) |
| `cell:` absolute target already passed by cursor | Fatal error              |
| `lbl:` absolute target is empty            | Fatal error                    |
| Data cell fails type/regex validation      | Warning logged, value kept     |
| Data cell empty where a value was expected | Warning logged, `null` stored  |

Fatal errors stop processing immediately and return a partial (possibly empty) result. Warnings are collected and reported in the summary.
