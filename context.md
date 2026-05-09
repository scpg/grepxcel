# Excel Processing Project — Context & Knowledge Base

## Goal
Build a **pattern-based Excel extraction engine** in Python. A human defines a pattern once; the engine finds all instances of that pattern across any conforming Excel file and outputs clean, structured data.

Core philosophy: no automatic recognition — explicit, human-defined patterns applied by a deterministic engine. Like `grep` but for Excel structure.

---

## Excel File Types Encountered

### Type 1 — Simple flat table
- One header row, data rows below
- Content: text, images, URLs
- Straightforward to extract

### Type 2 — Multi-entity mini-table sheet
- A single sheet contains **multiple repetitions of the same table structure**
- Each mini-table represents the same kind of information for a different entity (e.g. one per European country)
- Users chose this layout deliberately — seeing all entities on one sheet helps human processing
- Mini-tables are **not in fixed positions** — layout varies; spatial position is not semantically meaningful
- Each mini-table has a **consistent header row** (same column headers every time)

---

## Pattern Engine Design (v2 — based on pattern-example-2.xlsx)

The pattern file is a **meta-description** — its instructions tell the engine what to find and how to extract it. Real data files contain none of these markers.

---

### Section 1: `config:` — global engine settings

| Key | Example value | Meaning |
|---|---|---|
| `read.direction` | `LR` | Direction to scan the whole Excel file. `LR` = left-to-right (mini-tables side by side). `TD` = top-down (stacked). |
| `currency.sign` | `€` | Currency symbol applied to `currency` type fields |
| `empty.aliases` | `-, N/A, n/a` | Additional values treated as empty (organisation-specific placeholders) |
| *(future)* | *(TBD)* | Direction to scan within a mini-table group |

---

### Section 2: `def:` — named field registry

Format: `def: <field.name> <type> <regex>`

- Field names use dot notation for hierarchy: `requestor.name.data`, `header.item.qty.label`
- **`.label` suffix** = fixed-text locator cell — regex is the text to match in the real file (used to find/anchor the cell)
- **`.data` suffix** = actual value to extract — regex validates the extracted value
- Types: `string`, `integer`, `currency`, `date`, `datetime`, `timestamp`
- Regex used for both **location** (labels) and **validation** (data fields)

---

### Section 3: `START:` — sequential layout description

Describes the real Excel content **in order**, as the engine encounters it reading in `read.direction`. The engine follows this sequence top-to-bottom, matching each element in turn.

**Top-level elements:**

| Tag | Meaning |
|---|---|
| `cell:1` | The next **non-empty cell** encountered while scanning in `read.direction`. Position-agnostic — the engine simply advances until it finds a cell with a value. |
| `table:*` | A group of mini-tables sharing the same structure (multiple instances expected) |

**Within a `table:*` block:**

| Tag | Meaning |
|---|---|
| `HEADER:1` | Exactly 1 header row |
| `HEADER:*` | Multiple header rows |
| `SPLITTER:1` | Exactly 1 splitter row (empty in real file — visual padding only) |
| `DATA:1` | Exactly 1 data row |
| `DATA:*` | Unbounded data rows — stop when all non-IGNORE columns are empty, or next HEADER found |
| `DATA:<n>` | Exactly N data rows — validated, fails if count doesn't match |
| `FOOTER:1` | Exactly 1 footer row |
| `FOOTER:*` | Multiple footer rows |

**Cell-level keywords:**

| Keyword | Meaning |
|---|---|
| `<field.name>` | Reference to a `def:` entry — locate, extract and validate this cell |
| `EMPTY` | This cell must be empty in the real file |
| `IGNORE` | Find the next cell that has a value and skip it — do not extract or validate |

---

### Real Excel layout described by pattern-example-2

```
cell:1   → merchandise name (data)
cell:1   → requestor name label
cell:1   → requestor name data
cell:1   → requestor dept label
cell:1   → requestor dept data
cell:1   → IGNORE (skip a row)
table:*  → first group of mini-tables (HEADER×2 + SPLITTER + DATA:* + FOOTER×2), repeated N times
cell:1   → IGNORE (skip a row)
table:*  → second group of mini-tables (same structure), repeated N times
```

---

### Mini-table sub-routine

Invoked when the engine hits a `table:*` entry in the template. Mini-tables can be arranged side by side, stacked, or in a grid — the engine handles all cases without needing to know the layout.

**Full validation before commit** — the engine attempts a complete match across all sections. Only if everything passes are cells marked consumed. Any failure at any section → reject candidate, advance, try next.

**Validation sequence for each candidate position:**

1. **HEADER rows** — for each HEADER row: do ALL labels match their regex? Do all data-type fields match their type? → any failure: reject
2. **SPLITTER rows** — are ALL cells considered empty (see EMPTY definition below)? → any failure: reject
3. **DATA rows** — does each cell respect its defined data type per the `def:` entry? → any failure: reject
4. **FOOTER rows** — do ALL labels match their regex? Do all data-type fields validate? → any failure: reject
5. **All pass** → commit: extract data, mark entire region as consumed

After each successful commit, advance and attempt the next candidate. When no candidate produces a full match → `table:*` group exhausted, control returns to parent sequence.

### EMPTY cell definition

**Hardcoded as always-empty (non-configurable):**
- `None` / truly blank cell
- Whitespace-only strings (` `, `\t`, `\n`, `\r`)
- Non-breaking space (`\xa0`) — common Excel paste artifact
- Zero-width characters (`​`, `‌`, `‍`, `﻿` BOM) — invisible copy-paste artifacts
- Formula evaluating to `""`

**Configurable via `config: empty.aliases`:**
- Business-specific placeholder values that should be treated as empty (e.g. `"-"`, `"N/A"`, `"n/a"`, `"---"`)
- Vary by organisation — not hardcoded

### End-of-mini-table detection (for `DATA:*`)
A data section ends when:
1. All non-`IGNORE` columns in the current row are empty, **or**
2. The next `HEADER` fingerprint anchor is found (start of next mini-table instance)

**Open question:** can empty rows appear legitimately *inside* a DATA section? To be determined when real data files are available. If yes, a stricter mode will be needed.

---

### Core engine algorithm
1. Start at first non-empty cell, scanning in `read.direction`
2. Look at current template entry:
   - **`cell:1`** → read cell, validate against `def:` (type or content), mark consumed, advance to next non-empty cell
   - **`EMPTY`** → skipped automatically (engine only lands on non-empty cells)
   - **`table:*`** → invoke mini-table sub-routine (see below)
3. Move to next template entry, repeat until template exhausted

### Greedy vs lazy matching — priority rule
Complex patterns always take priority over simple ones. A `cell:1 string` matches almost any non-empty cell, including cells that belong to a mini-table header — so simple matches must never run before complex ones.

**Rule: when a `table:*` is active, the engine attempts a full mini-table match at every candidate position BEFORE considering any `cell:1` entries.**

- `table:*` is **greedy** — keeps consuming mini-table instances as long as matches succeed. The group is only considered exhausted when a mini-table match fails.
- `cell:1` is the **fallback** — only matches cells that the greedy `table:*` pass could not claim.
- The failed mini-table match is itself the signal that the `table:*` group is done and control passes to the next template entry.

This mirrors regex greedy/lazy semantics and prevents simple type matches from incorrectly consuming mini-table structure cells.

### Key engine behaviours
- Engine follows `START:` sequence in order — position-agnostic, advances to next non-empty cell
- **Cell consumption is atomic** — a cell is marked consumed only after a successful match. Failed matches cause no consumption (rollback).
- **Mini-table matching is transactional** — the engine tentatively peeks at candidate cells to check if the full mini-table pattern matches. Only on a complete positive match are all cells in that region marked consumed and data extracted. Partial matches are rolled back entirely.
- Consumed cells are never re-processed
- `def:` regex patterns used for both locating labels and validating data values
- Type mismatches should be flagged as errors
- `DATA:<n>` count mismatch should be flagged as an error

---

## Output Format
Still to be confirmed. JSON is the leading candidate, with field names from the `def:` registry as keys.

---

## Python Stack
| Library | Purpose |
|---|---|
| `openpyxl` | Low-level cell-by-cell access (required for spatial scanning) |
| `pandas` | Optional: data transformation after extraction |
| `SQLAlchemy` | Optional: load output into a database |

Note: `pandas.read_excel()` is too high-level — it assumes one flat table. `openpyxl` gives the direct cell access needed for the pattern scanner.

---

## Open Questions
1. What is the final consumer of the extracted data? (another app, a database, a web app?)
2. Can empty rows appear legitimately inside a `DATA:*` section?
3. Direction config for scanning within a mini-table group (to be added later)

---

## Broader Landscape (for later phases)
If the goal eventually becomes a full web app on top of the data:
- **Baserow** — spreadsheet-database hybrid with built-in app builder, self-hostable via Docker
- **NocoDB** — visual layer on top of existing PostgreSQL/MySQL
- **Grist** — best for calculation-heavy workflows (supports Python in formulas)
- **Budibase** — rapid CRUD app builder with auto-scaffolded UI

Recommended pipeline for complex cases:
`openpyxl (extraction)` → `PostgreSQL` → `Baserow or NocoDB (app layer)`

---

## Environment
- OS: WSL2 (Linux on Windows)
- Docker Desktop: installed and running
- Working directory: `/mnt/c/dev/2026/excel`
- Language: Python
- Virtual environment: `.venv` — always use `.venv/bin/python` and `.venv/bin/pip`
- Test scripts: `tmp-scripts/` folder — reuse before recreating
