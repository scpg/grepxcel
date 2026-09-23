# Pattern File Reference

A pattern file is an **`.xlsx` workbook** or a **`.csv` text file** that tells grepxcel what to look for and extract from a data file. Think of it as a schema: it describes the layout, field names, types, and validation rules. Both formats are read into the same internal grid and behave identically.

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

| Column A  | Column B           | Column C                          | Notes                        |
|-----------|--------------------|-----------------------------------|------------------------------|
| `config:` | `pattern.version`  | integer (e.g. `1`)                 | Default: `1`                 |
| `config:` | `read.direction`   | `LR` or `TD`                      | Default: `LR`                |
| `config:` | `currency.sign`    | e.g. `€` or `$`                   | Default: `€`                 |
| `config:` | `ignore.case`      | `true` or `false`                  | Default: `false`             |
| `config:` | `trim.whitespace`  | `true` or `false`                  | Default: `false`             |
| `config:` | `lbl.match`        | `literal`, `glob`, or `regexp`     | Default: `literal`           |
| `config:` | `var.match`        | `glob`, or `regexp`               | Default: `glob`              |
| `config:` | `empty.aliases`    | e.g. `N/A`                        | Repeat the row for each alias|

**`read.direction`** controls how the data sheet is scanned for `cell:` instructions:

- `LR` — left-to-right, top-to-bottom (row by row). Most common.
- `TD` — top-to-bottom, left-to-right (column by column).

**`currency.sign`** sets the expected currency symbol for `currency` fields (e.g. `€`, `$`). The sign is stored in the pattern and available to custom regex — Excel stores currency cell values as plain numbers, so a regex like `\d+\.\d{2}` matches the raw value directly without the sign prefix.

**`empty.aliases`** lists strings that should be treated as empty cells (e.g. `N/A`, `-`, `—`). Add one alias per row.

**`ignore.case`** makes all regex matching case-insensitive (applies to both `lbl:` and `var:` fields).

**`trim.whitespace`** strips leading and trailing whitespace from every cell value before matching and before writing to the JSON output. Off by default (opt-in). Per-field override: add `trim-whitespace` to the column-A modifiers. See [Whitespace handling](#whitespace-handling) below.

**`lbl.match`** controls how `lbl:` field values are matched against data cells:

- `literal` (default) — exact string match. Labels like `Term (months):` or `Invoice No.` work without escaping.
- `glob` — shell-style wildcards: `*` matches any text (including newlines), `?` matches one character.
- `regexp` — full Python `re.search` behaviour (the pre-v0.1.0 default).

Per-field overrides are also supported: write `lbl:literal`, `lbl:glob`, or `lbl:regexp` in column A instead of plain `lbl:`.

**`pattern.version`** declares a forward-compatibility version. Currently only version `1` is defined; absent defaults to `1`. Future versions may add new syntax.

---

## Field definition rows: `lbl:`, `var:`, `doc:`

At least one definition row is required. Each defines one field that can be
referenced in the `START:` section, or documents the pattern.

| Column A | Column B      | Column C   | Column D              | Role |
|----------|---------------|------------|-----------------------|------|
| `lbl:` *[modifiers]* | `FieldName` | type | match pattern | **Anchor** — matched for position only; **never written to output JSON** |
| `var:` *[modifiers]* | `field.name` | type | pattern/regex | **Variable** — extracted and written to output JSON |
| `doc:`   | (free text)   |            |                       | **Comment** — ignored by the engine |
| `def:`   | `FieldName`   | type       | regex                 | Backward-compatible alias for `var:` |

Column A supports **order-independent colon-separated modifiers** (see [Column A modifiers](#column-a-modifiers) below).

- **`lbl:`** — use for literal text that marks *where* a value lives: labels like
  `Invoice No:` or column headers like `Product`, `Qty`. Matched but stripped
  from output. Column D is matched according to the `lbl.match` config:
  - **`literal`** (default) — exact string match. Write the label text as-is.
  - **`glob`** — shell wildcards (`*`, `?`).
  - **`regexp`** — full Python `re.search`.
  - Per-field override: use `lbl:literal`, `lbl:glob`, `lbl:regexp`, or `lbl:re` (alias for `regexp`) in column A.
  - Constraint: add `not-null` to make a missing/empty label a fatal error (e.g. `lbl:not-null`).
  - Whitespace: add `trim-whitespace` to strip cell spaces before matching (e.g. `lbl:trim-whitespace`).
- **`var:`** — use for every value you want to capture. **Dot notation creates
  nested JSON**: `po.number` → `{"po": {"number": …}}`. In a table, the group
  prefix (`line` in `line.qty`) becomes the output array key.
  - Default: column D is a Python `re.fullmatch` regex. Use `.*` to accept anything.
  - `var:glob` — column D is a shell-style glob (`*`, `?`). Type check still runs first.
  - `var:literal` — column D is an exact string match. Special regex characters are literal.
  - `var:re` / `var:regexp` — explicit alias for the default regex mode.
  - Constraint: add `not-null` or `not-empty` (synonyms) to make an empty/null value a **fatal error**, always, regardless of `--strict`.
  - Whitespace: add `trim-whitespace` to strip leading/trailing spaces before matching and in the extracted JSON value.
- **FieldName** — a unique plain-text identifier.
- **regex** (for `var:` / `def:`) — a Python `re.fullmatch` pattern applied to
  the string representation of the cell value. Use `.*` to accept anything.
  Nested unbounded quantifiers (e.g. `(a+)+`) are rejected as unsafe (ReDoS guard).

---

### Column A modifiers

Modifiers are colon-separated tokens after the row keyword. They are **order-independent** and can be combined freely:

| Example column A | Meaning |
|------------------|---------|
| `lbl:` | Anchor, literal match (default) |
| `lbl:literal` | Anchor, explicit literal match |
| `lbl:glob` | Anchor, glob match for this field |
| `lbl:regexp` / `lbl:re` | Anchor, regex match for this field |
| `lbl:not-null` | Anchor, literal match; fatal if the cell is empty |
| `lbl:not-null:glob` | Anchor, glob match, required |
| `lbl:not-null:regexp` | Anchor, regex match, required |
| `lbl:trim-whitespace` | Anchor — strip cell spaces before matching |
| `lbl:not-null:trim-whitespace` | Anchor — required + strip spaces |
| `var:` | Variable, regex match in column D (default) |
| `var:re` / `var:regexp` | Variable, explicit regex mode (same as default) |
| `var:glob` | Variable, glob match in column D |
| `var:literal` | Variable, exact string match in column D |
| `var:not-null` | Variable, required — fatal error if value is empty/null |
| `var:not-empty` | Synonym for `var:not-null` |
| `var:not-null:re` | Variable, required, explicit regex |
| `var:not-null:glob` | Variable, glob match, required (order-independent) |
| `var:literal:not-empty` | Variable, literal match, required |
| `var:trim-whitespace` | Variable — strip cell spaces before matching and in extracted JSON |
| `var:not-null:trim-whitespace` | Variable — required + strip spaces |
| `var:glob:trim-whitespace` | Variable — glob match + strip spaces |
| `var:literal:trim-whitespace` | Variable — literal match + strip spaces |

**`not-null` / `not-empty` are synonyms.** Both trigger a fatal error (not a warning) when the extracted value is empty or null — regardless of whether `--strict` is used. This is stronger than the default behaviour where missing values are silently set to `null`.

**`var:glob` and `var:literal`** apply the type check first (column C), then match column D against the string representation of the value. The regex safety guard does not apply (column D is never compiled as a regex). Special characters in column D (`(`, `)`, `*`, `.`, etc.) are interpreted literally in `var:literal` mode, and as glob wildcards in `var:glob` mode.

**`trim-whitespace`** strips leading and trailing whitespace from the cell value before matching and, for `var:` fields, before writing to the JSON output. This is **opt-in** — off by default. Apply it per-field (column A) or globally with `config: | trim.whitespace | yes`.

```
var:not-null          invoice.number  string  INV-\d+    ← regex, required
var:glob              sku             string  PROD-*     ← glob, optional
var:literal           status          string  Active     ← exact match, optional
var:not-null:literal  currency        string  EUR        ← exact match, required
var:trim-whitespace   company         string  .*         ← strip " Acme Corp  " → "Acme Corp"
var:not-null:re       code            string  [A-Z]{3}   ← explicit regex, required
lbl:not-null          inv_label       string  Invoice:   ← anchor must be present
lbl:trim-whitespace   header          string  Date       ← strip spaces before matching
```

---

### Whitespace handling

Excel cells occasionally contain invisible leading or trailing spaces — typed accidentally, pasted from another source, or left by a formula. These are **not stripped by default**: the raw cell value is matched and extracted as-is.

**When whitespace causes a mismatch, grepxcel tells you.** The validation warning for a mismatched field includes a targeted hint:

```
⚠  B3  [company / string]
   Found:    ' Acme Corp  '
   Expected: matches /\w+/
   → The cell value has leading or trailing whitespace.
     The trimmed value 'Acme Corp' DOES match /\w+/.
     Add 'trim-whitespace' to this field (e.g. 'var:trim-whitespace') or set
     'config: | trim.whitespace | yes' to strip whitespace globally.
```

**To fix it, use `trim-whitespace`** (opt-in, not the default):

| Scope | Syntax | Effect |
|-------|--------|--------|
| Per field | `var:trim-whitespace` | Strip before matching and in extracted JSON |
| Per label | `lbl:trim-whitespace` | Strip before matching the anchor (never in output) |
| Whole pattern | `config: \| trim.whitespace \| yes` | Strip all fields globally |

**What trim-whitespace does:**
- Strips **leading and trailing** whitespace only (Python `str.strip()`).
- For `var:` fields: the **extracted JSON value** is the trimmed string.
- For `lbl:` fields: only the matching comparison is trimmed; labels are never in output anyway.
- Applies to `string`/`text` cells only. Numeric, date, and boolean cells are Python objects — no whitespace to strip.

**Why it is opt-in:** Most patterns expect exact values. Automatically stripping could silently change the extracted data, breaking downstream consumers that rely on the raw cell text. Making it explicit keeps behaviour predictable and puts the author in control.

### Supported types

| Type         | Accepts                                       | Regex target           |
|--------------|-----------------------------------------------|------------------------|
| `string`     | Any text                                      | The cell text as-is    |
| `text`       | Alias for `string`                            | The cell text as-is    |
| `integer`    | Whole numbers (Excel integers or whole floats)| `str(int_value)`       |
| `number` / `float` / `decimal` | Any number, int or float (not a boolean)    | `str(numeric_value)`   |
| `currency`   | Any number (int or float)                     | `str(numeric_value)`   |
| `percentage` | Any number (int or float). Excel stores a percentage as a fraction, e.g. 62.5% → `0.625` | `str(numeric_value)` |
| `boolean` / `bool` | Excel `TRUE`/`FALSE`                    | n/a (type check only)  |
| `time`       | Excel **time-only** cells (no date part)      | n/a (type check only)  |
| `duration`   | Excel `[h]:mm` durations (timesheets, elapsed time) | n/a (type check only) |
| `date`       | Excel date cells                              | n/a (type check only)  |
| `datetime`   | Excel datetime cells                          | n/a (type check only)  |
| `timestamp`  | Same as `datetime`                            | n/a (type check only)  |

`text` is a synonym for `string`. `number`/`float`/`decimal` behave like
`currency`/`percentage` (any numeric cell) but carry no money/percentage intent —
use them for plain decimals such as quantities or measurements. `percentage`
validates identically to `currency`; it exists to document intent. For
`boolean`/`time`/`duration`/`date`/`datetime`/`timestamp`, the regex column is
ignored. `time` matches **time-only** cells (e.g. `14:30`); a cell that also has
a date is a `datetime`, not a `time`. `duration` matches Excel `[h]:mm` formatted
cells (e.g. `8:30` meaning 8 hours 30 minutes) — useful for timesheets and
elapsed-time columns.

> An unknown type name (e.g. a typo like `currncy`) is rejected when the pattern
> file is parsed, so a mistyped type fails fast instead of silently mis-validating.

---

## Comments

Two ways to annotate a pattern file:

**Comment rows** — a `doc:` (or `info:`) in column A makes the whole row a
comment, ignored by the engine. Use these for headings or notes that sit on their
own line (including above a `table:` block).

**Trailing `#` comments** — on a `config:`, `var:`/`def:`, `lbl:`, `cell:`, or
`START:` row, a cell whose text starts with `#` begins a comment that runs to the
end of the row. The comment must come **after** the row's real columns:

| Row | Real columns | Comment goes in |
|-----|--------------|-----------------|
| `var:` / `lbl:` / `def:` | A–D (keyword, name, type, regex) | column **E** onward |
| `cell:` | A–B (instruction, field) | column **C** onward |
| `config:` | A–C (keyword, key, value) | column **D** onward |
| `START:` | A | column **B** onward |

```
var:       amount.net    currency   \d+(\.\d{2})?   # net amount, before VAT
cell:next  amount.net                               # first value after the label
```

Rules:

- A `#` in a **value column is a value, not a comment** — e.g. the regex in
  `var: | code | string | #\d+` matches `#123`; the `#` is part of the regex.
- **`table:` rows do not support `#` comments.** `HEADER:` / `DATA:` / `FOOTER:`
  rows use their trailing columns for the table's own columns, so `#` there is a
  literal value (free to use). Annotate a table with a `doc:` row above it instead.
- **Fail fast:** any non-empty cell in a comment position that does *not* start
  with `#` is rejected as an error — this catches a value typed into the wrong
  column instead of silently dropping it.

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

## seek: instruction

Reposition the scanner cursor to a specific cell **without reading it**. The
next `cell:next` after a `seek:` starts scanning from the new position.

| Column A   | Meaning |
|------------|---------|
| `seek:G5`  | Move cursor to G5; does not consume the cell |

**When to use it:** when you need to read scattered absolute cells from
different regions of the sheet and the natural scan order would require jumping
backward. A typical pattern is: read a group of cells from one column, then
`seek:` back to an earlier row in a different column and read those cells with
absolute references.

```
cell:B7    | company.phone     ← last cell in left column (row 7)
seek:I4                        ← reposition: back to row 4, right column
cell:I4    | employee.name     ← absolute ref at the seek position is valid
cell:I5    | employee.manager
cell:I6    | employee.week_starting
seek:B9                        ← reposition again before the table
table:*
```

**Ordering rules after `seek:`:** the abs-ref ordering constraint resets
completely. The first `cell:abs` after a `seek:` may be at, before, or after
the seek target — the parser accepts it; backward references are caught at
runtime by the engine's "already passed" check.

**`seek:` does not read the target cell.** It only moves the cursor. Already-
consumed cells are still skipped by subsequent `cell:next` instructions.

`seek:` is not valid inside a table block. Annotate it with a `doc:` row above
if you want to explain why the reposition is needed.

---

## dir: instruction

Switch the scan direction partway through extraction. This affects all subsequent
`cell:next` instructions (until another `dir:` changes it again).

| Column A   | Meaning |
|------------|---------|
| `dir:LR`   | Switch to left-to-right, top-to-bottom (row by row) |
| `dir:TD`   | Switch to top-to-bottom, left-to-right (column by column) |

**When to use it:** when a data sheet has a header region that reads left-to-right
but a body that reads top-to-bottom (or vice versa). Instead of using absolute
`cell:` references for every field, switch direction inline and continue with
`cell:next`.

```
config:    read.direction   LR           ← start reading row-by-row
lbl:       header_label     string       Name:
var:       name             string       .*

START:
cell:next  header_label
cell:next  name
dir:TD                                   ← switch to column-by-column
cell:next  first_column_value
```

`dir:` is not valid inside a table block. The table's own `config: read.direction`
controls direction within the table.

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

**DATA rows** — come in three forms:

| Multiplicity | Meaning |
|---|---|
| `DATA:*` | Greedy — collect rows until empty row or FOOTER detected |
| `DATA:1` | Collect exactly one physical row per instance |
| `DATA:{n,m}` | Scan at most *m* physical rows total; warn if fewer than *n* non-skipped rows are found |

Missing values in DATA rows produce a warning but do not reject the match.

**SKIP_IF rows** — silently exclude a data row from output when the row matches
the condition. Useful for skipping empty or irrelevant rows without rejecting the
whole table instance. Column positions use the same field keywords as DATA:

| Column keyword | Meaning in SKIP_IF |
|---|---|
| `EMPTY` | This column's cell must be empty for the row to be skipped |
| `IGNORE` | Do not check this column |

A row is skipped when **all** non-`IGNORE` columns satisfy their condition.
`SKIP_IF` is valid with `DATA:*` and `DATA:{n,m}`. It is not valid with `DATA:1`.

With `DATA:{n,m}`: skipped rows still count toward the `{n,m}` bounds.
With `DATA:*`: skipped rows are silently filtered; scanning continues.

```
    | SKIP_IF   | IGNORE | EMPTY | EMPTY | IGNORE |   ← skip rows with empty col2 and col3
    | DATA:{0,7}| date   | timeIn | breaks | total |
```

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
| Consecutive absolute `cell:` references out of reading order | Fatal error (at parse time) |
| `cell:` absolute target already passed by cursor | Fatal error              |
| `seek:` target outside the sheet's used range | Fatal error                |
| `lbl:` absolute target is empty            | Fatal error                    |
| Data cell fails type/regex validation      | Warning logged, value kept     |
| Data cell empty where a value was expected | Warning logged, `null` stored  |

Fatal errors stop processing immediately and return a partial (possibly empty) result. Warnings are collected and reported in the summary.

---

## Merged cells

Excel cells can span multiple columns or rows (merged ranges). grepxcel reads merged
cells from the **top-left cell** of the merge range only. The remaining cells in the
range appear empty to the engine.

### Labels on merged cells

If an anchor label sits on a merged cell, place the `lbl:` definition for that cell's
top-left address in your pattern. The engine will find it normally.

### Values spanning merged cells

If a data value occupies a merged range, the value is read from the top-left cell.
The other cells of the range are invisible — do not define `var:` entries for them.

### Tables with merged header rows

When a table has a header row where some column headers span multiple cells, the
header labels for the merged cells will appear in the leftmost column of the merge.
Subsequent columns of the merge will look empty. Use `IGNORE` or omit them from
the `HEADER:1` row in your pattern.

### Pattern file itself

The pattern file should not use merged cells. The parser reads each cell by its
exact address; merged ranges in a pattern file may lead to blank cells where the
parser expects instructions.

### `grepxcel lint` merged-cell report

Running `grepxcel lint data.xlsx` reports all merged ranges in the file. This is
useful to identify which cells in a complex sheet are top-left anchors before
building a pattern.
