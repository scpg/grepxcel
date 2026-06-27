# Design: `grepxcel wizard` command

**Date:** 2026-06-27  
**Status:** Approved  
**Scope:** New interactive command that walks a user cell-by-cell through an Excel file and produces a valid CSV pattern.

---

## Problem

The `draft` command auto-generates a pattern using a local LLM. When no LLM is available, or when the user wants full manual control, there is no guided path — they must understand the pattern format from docs and write the file from scratch. This is a high barrier for non-technical users.

## Solution

`grepxcel wizard data.xlsx` — a terminal wizard that reads the data file, asks the user to classify each cell (label, variable, skip, table…), and writes a complete, valid CSV pattern.

---

## Invocation

```bash
grepxcel wizard data.xlsx
grepxcel wizard data.xlsx --sheet Sheet2
grepxcel wizard data.xlsx -o my-pattern.csv
```

If `-o` is omitted, the output file is suggested as `pattern-<data-stem>.csv` in the same directory as the data file. The user can confirm or change it at the end of the session.

---

## Flow

```
Phase 1 — Config
Phase 2 — Cell walk  (loops until user types E)
  └── Mini-table sub-flow  (entered when user types T)
Phase 3 — Summary + save
```

---

## Phase 1: Config

The wizard loads the workbook and asks config parameters one at a time. The current/default value is shown in brackets; pressing Enter accepts it.

```
Sheet [Sheet1]:
Scan direction (LR / TD) [LR]:
Ignore case (yes/no) [no]:
Currency symbol [€]:
```

- If the workbook has only one sheet, the sheet question is skipped.
- Scan direction determines the order cells are visited in Phase 2: LR = left-to-right across each row, then down; TD = top-to-bottom down each column, then right.

---

## Phase 2: Cell walk

The wizard maintains a **cursor** over the sheet and visits cells in reading-direction order. For each cell it displays:

```
──────────────────────────────────────────────
Cell B2  │  "Invoice Number"
Proposed: label  (short text, looks like a heading)
──────────────────────────────────────────────
  [L] Label     [V] Variable   [S] Skip
  [T] Table     [I] Ignore     [G] Goto <ref>
  [E] End
>
```

### Type proposal logic

Based on the openpyxl cell value (using the existing `infer_cell_type` utility):

| Cell value | Proposed type |
|------------|--------------|
| Short string (≤ 40 chars) | label |
| Longer string | var:string |
| Integer or whole-number float | var:integer |
| Fractional float | var:currency |
| `datetime.date` / `datetime.datetime` | var:date or var:datetime |
| `None` / empty | Skip |

### User commands

| Input | Meaning |
|-------|---------|
| `L` | Mark as `lbl:` — prompts for label name (default: auto-slug of cell value) |
| `V` | Mark as `var:` — prompts for field name, type (default: proposed), and match pattern (default: `.*`) |
| `S` | Emit `cell:1, SKIP` — cell is consumed but its value is not extracted |
| `I` | Emit `cell:1, IGNORE` — cell is read and discarded |
| `T` | Enter mini-table sub-flow at this position |
| `G A5` or just `A5` | Move cursor to cell A5 without emitting anything |
| `E` / `end` | End the walk and go to Phase 3 |
| *(Enter)* | Accept proposed type with prompted defaults |

### Label prompt

```
Label name [invoice_number]: inv.number_lbl
```

Auto-slug default: lower-case cell value, spaces → underscores, non-alphanumeric stripped.

The wizard emits:
- A `lbl:` field definition in the header section.
- A `cell:1, <label_name>` instruction in the START/END body.

### Variable prompt

```
Field name: inv.number
Type [string]: 
Match pattern [.*]: [A-Z]{2}[0-9]{6}
```

The wizard emits:
- A `var:` field definition in the header section with the given name, type, and match pattern.
- A `cell:1, <field_name>` instruction in the START/END body.

---

## Mini-table sub-flow (entered with T)

### Step 1: Multiplicity

```
How many instances?
  1      — exactly one
  N      — a specific number (you will enter it)
  N..M   — a range (e.g. 1..10)
  *      — any number (most common)
> *
```

Maps to `table:1`, `table:N`, `table:{N,M}`, or `table:*` in the pattern.

### Step 2: HEADER row

The wizard visits every cell in the current row and asks for a field name. The user can also mark a cell as IGNORE:

```
--- HEADER row (row 5) ---
Cell A5  "Item"    → field name [item]: row.item
Cell B5  "Qty"     → field name [qty]:  row.qty
Cell C5  "Price"   → field name [price]: row.price
Cell D5  "Total"   → field name [total]: row.total
Cell E5  (empty)   → [S]kip / [I]gnore / [done]: done
```

`done` ends the header definition. Any cells marked S or I are recorded as SKIP / IGNORE in the HEADER row.

### Step 3: DATA row template

The wizard visits the next row and for each cell proposes the matching header field name. The user confirms, changes, or marks as SKIP/IGNORE:

```
--- DATA row (row 6, template) ---
Cell A6  "Laptop"  → [D]ata (row.item) / [S]kip / [I]gnore: D
Cell B6   2        → [D]ata (row.qty)  / [S]kip / [I]gnore: D
Cell C6   1200.00  → [D]ata (row.price) / [S]kip / [I]gnore: D
Cell D6   2400.00  → [D]ata (row.total) / [S]kip / [I]gnore: D
```

### Step 4: FOOTER (optional)

```
FOOTER row? (yes/no) [no]: yes
--- FOOTER row ---
Cell A11  "Grand Total"  → field name [footer.label]: footer.label
Cell B11  (empty)        → [I]gnore: I
Cell C11  (empty)        → [I]gnore: I
Cell D11   3030          → field name [footer.value]: footer.value
End of footer.
```

### After the table

The cursor resumes at the first cell after the table range. The main walk continues.

---

## Phase 3: Summary and save

```
─────────────────────────────────────────────────
Pattern summary
  Direction : LR
  Currency  : €
  Labels    : 3
  Variables : 8
  Tables    : 1  (multiplicity=*)

Save pattern to [pattern-invoice.csv]:
─────────────────────────────────────────────────
Pattern written to: pattern-invoice.csv

Try it:
  grepxcel extract -p pattern-invoice.csv invoice.xlsx
  grepxcel validate-pattern pattern-invoice.csv
─────────────────────────────────────────────────
```

---

## Output format

Always CSV (same format as `pattern-from-draft.csv`). Structure:

```
config:,read.direction,LR
config:,currency.sign,€
lbl:,<name>,string,<cell_value>
...
var:,<name>,<type>,<match>
...
START:
cell:1,<name_or_SKIP_or_IGNORE>
...
table:*
,HEADER:1,field1,field2,...
,DATA:*,field1,field2,...
,FOOTER:1,field1,IGNORE,...   ← only if user added footer
END:
```

All `var:` field definitions for table columns are also written in the header section so the engine resolves them.

---

## Implementation

### New file

**`grepxcel/wizard.py`** — self-contained module.  
Public entry point: `run_wizard(data_file: str, sheet: str | None, output: str | None) -> int`

Internal structure:
- `WizardState` dataclass — holds config, accumulated lbl/var definitions, cell instructions, table instructions, cursor position
- `_ask(prompt, default)` — single-line input helper (respects TTY, handles Ctrl-C gracefully)
- `_propose_type(cell_value)` — single-cell type inference (wraps `infer_cell_type`)
- `_run_config_phase(ws, state)` — Phase 1
- `_run_cell_walk(ws, state)` — Phase 2 outer loop
- `_run_table_subflow(ws, state, start_row)` — mini-table sub-flow
- `_write_pattern(state, output_path)` — Phase 3 output

Uses only stdlib + openpyxl (already a hard dependency). ANSI colour helpers copied from `quickstart.py` style (no extra deps).

### Changes to existing files

**`grepxcel/cli.py`**
- Add `_add_wizard_subparser(sub)` — registers `wizard` with `data_file`, `--sheet`, `-o/--output` args
- Add `'wizard'` dispatch case in `main()` → `from .wizard import run_wizard; return run_wizard(...)`
- Add one-line entry in the help text block

---

## Deferred (not in this release)

- `SKIP_IF` / conditional rows
- `SPLITTER` rows
- Multi-sheet table spanning
- `--resume` (save and continue a session)
- XLSX output (CSV is sufficient; the engine accepts both)

---

## Testing

- `tests/unit/test_wizard.py` — unit tests for `_propose_type`, `_write_pattern`, and the state accumulator using scripted input (monkeypatched `input()`)
- No integration test that drives the full interactive flow (TTY interaction is hard to test end-to-end); the unit tests cover the logic; the pattern the wizard writes is validated by running it through `validate-pattern`
