# The grepxcel Wizard — User Guide

> **STATUS: DRAFT FOR REVIEW**
> This document describes the current behaviour of `grepxcel wizard`.
> Items marked ⚠️ are known gaps or bugs that need fixing.
> Please correct the text and add notes — the code will be updated to match.

---

## What the wizard does

The wizard walks you through a real Excel file, cell by cell, and lets you classify each
non-empty cell as a **label**, **value to extract**, **table column**, or **ignore**.
At the end it writes a **pattern file** — a plain CSV (or xlsx) that tells
`grepxcel extract` what to find and where.

The wizard does not extract data itself. It produces the pattern; you then run
`grepxcel extract -p pattern.csv data.xlsx` to get JSON output.

If you want to know all paramenters for the wizard command, then check
```bash
grepxcel wizard -h
grepxcel wizard --help
```


---

## Starting the wizard

```bash
grepxcel wizard data.xlsx                             # write pattern to stdout
grepxcel wizard -o pattern.csv data.xlsx              # write to file
grepxcel wizard -o pattern.xlsx data.xlsx             # ⚠️ see note below
grepxcel wizard --sheet "Sheet2" data.xlsx            # choose a specific sheet

```

> ⚠️ **Known issue — output format:** The `-o` flag accepts any filename, but
> the wizard always writes **CSV text** regardless of the extension. If you write
> `-o pattern.xlsx` the file will contain CSV, not a real Excel workbook. You can
> open it in Excel (it will warn you), or rename it `.csv`.
> Fix pending: the wizard should write a real xlsx workbook when the extension is `.xlsx`.

The wizard launches a full-screen terminal UI (requires a terminal; falls back to a
sequential prompt if Textual is not installed).

---

## Step 1: Configuration modal

The first screen asks three questions:

| Field | Default | Meaning |
|---|---|---|
| **Read direction** | `LR` | How the scanner walks the data sheet: `LR` = row-by-row (left-right, top-bottom); `TD` = column-by-column |
| **Template mode** | `yes` | If the Excel file is a *template* (contains headers and structure but no real data), the wizard uses the structure to suggest names. Say `yes` for blank or partially-filled forms. |
| **Sheet** | first sheet | Which sheet to classify. |

> ⚠️ **To confirm and verify - for CLAUDE** 
> In template mode, the next field to jump to it is not the next non ampty field if the user just defined a cell as a label
> In template mode once a cell has been defined as a label, then the next cell should be the adjacent cell (in the direction defined for the sheet processing (next cell to the right or next cell down) depending on the directions)
> in the case of templates, this is the cell that is most probably going to contain the data

Press ENTER on each field, then ENTER on the last to proceed.

---

## Step 2: The cell grid

After configuration the wizard shows the worksheet as a grid. Non-empty cells appear in
colour; empty cells are dim. You navigate with arrow keys and classify cells with
single-key commands.

### Navigation

| Key | Action |
|---|---|
| Arrow keys | Move one cell |
| `G` | Jump to a specific cell reference (e.g. `C7`) |
| `N` | Advance to next non-empty cell (same as ENTER in sequential mode) |
| `U` / `Z` | Undo the last classification |
| `E` | End session and save the pattern |
| `ESC` | Cancel (no file written) |
| `F11` | Copy the session log to the clipboard |
| `F2` | Add a free-text note to the current cell (saved to the log) |

> ⚠️ **To confirm and verify - for CLAUDE** 
> the above help should be easily seen by the user of the awizzard

### Classification keys

| Key | Type | What it means |
|---|---|---|
| `L` | **Label** | Static text that identifies a nearby value (e.g. `DEPARTMENT:`). The engine uses it to *find* where you are on the sheet; it is **never written to the JSON output**. |
| `V` | **Value** | A cell whose content you want to extract. Written to the JSON output under the variable name you give it. |
| `C` | **Header** / section title | A decorative section title, no extraction value. Consumed by the engine to advance the cursor. |
| `T` | **Table** | Marks the start of a repeating row block. Opens a three-step table definition flow (see below). |
| `I` | **Ignore** | Consume the cell without extracting anything. |
| `R` | **Remove** | Clear the current cell's classification (or remove the whole table if on a table cell). |
| ENTER | **Accept proposal** | Accept the wizard's auto-suggestion (shown in the right panel). |

---

## Classifying scalar (non-table) cells

### Labels and Values go in pairs
The label will be used to confirm that the position in the excel sheet is correct and then the data will be read
One of the most common pattern is: a label cell identifies a value cell nearby.

```
       Col B           Col C
ROW 3: DEPARTMENT:     Accounting        ← B3 = label, C3 = value
ROW 4: CONTACT PHONE:  +1-555-0100       ← B4 = label, C4 = value
```

**Label (L)** — Press `L` on the cell that says `DEPARTMENT:`. A modal asks for:
- **Name** — a short identifier, e.g. `department_label`
- **Type** — almost always `string`
  > ⚠️ **To confirm and verify - for CLAUDE** 
  > The cell type sugestion should be based on the celly type / format in the *.xlsx being read
- **Match** — what text the label cell should contain. Defaults to the cell's current text. This is how the engine anchors the pattern to the right location on the sheet.
- **Notes** — optional free-text note saved to the session log

**Value (V)** — Press `V` on the cell next to or below the label. A modal asks for:
- **Name** — the variable name in the output JSON, e.g. `inventory.department`
  Use dot notation to create nested JSON: `inventory.department` → `{"inventory": {"department": ...}}`
- **Type** — `string`, `number`, `date`, `boolean`, etc. The wizard auto-suggests based on the Excel cell format.
  > ⚠️ **To confirm and verify - for CLAUDE** 
  > The cell type sugestion should be based on the celly type / format in the *.xlsx being read
- **Match pattern** — a regex the extracted value must satisfy. `.*` accepts anything.
  > ⚠️ **To confirm and verify - for CLAUDE** 
  > AS this tools is going to be used by many normal users it would be greate to provide here some presets
  > I have been thinking her eon URL, eMail by now, maybe the wizzard can already help the user on complex regexp 

- **Notes** — optional

> **Tip:** Name related variables with a common prefix to group them in the output JSON:
> `inventory.department`, `inventory.contact_person`, `inventory.date_of_order`
> → `{"inventory": {"department": "Accounting", "contact_person": "Jane", "date_of_order": "2026-01-15"}}`

### Headers (C)

Use `C` for section dividers (e.g. a merged cell that says "PART 1 — SUPPLIER INFO").
These tell the engine where it is without adding output fields.

### Ignore (I)

Use `I` for cells that the scanner must consume but that you don't care about —
decorative text, repeated titles, merged-cell artefacts.

---

## Classifying tables (T)

A *mini-table* is a repeating block of rows: 
* cero or more header rows (column titles)
* a possible splitter between header and data, but onyl if header exists
* one or more data rows (the actual records)
* a possible splitter between data and footer, but onyl if header exists
* and optionally a footer row (totals, notes).

### When to use T

Use `T` when the same columns repeat for multiple records — an invoice line-items table,
a product list, a weekly timesheet, etc.

Press `T` on the **top-left cell of the table block** (the anchor). A three-step flow
opens.

### Step T-1: Table setup

| Field | Example | Meaning |
|---|---|---|
| **Table name** | `items` | A short label used in the session log. ⚠️ See note below about JSON output. **Important to AI-CLAUDE this name will define the name of the "variable/structure" that will be used for generating the json structures of the mini tables recognized **|
| **Full table range** | `B11:F29` | The bounding box of the *entire* block, including header and footer rows. The main role for this is toassist  the user in the creation of the pattern |
| **Multiplicity** | `*` | How many instances of this table the engine should look for. `*` = any number; `1` = exactly one; `{n,m}` = between n and m. |

> ⚠️ **Table name and JSON output:** The table name you enter here is stored in the
> session log and used as a label in the wizard UI, but it does **not** appear in the
> pattern file. The grepxcel pattern format has no named-table concept.
>
> To make the extracted rows appear under a specific key in the JSON output, give your
> DATA variables a common dot-notation prefix:
> - `items.item_no`, `items.description`, `items.qty` → `{"items": [{"item_no": ..., ...}]}`
> - Or just plain names: `item_no`, `description` → rows appear at the top level
>
> Fix pending: the wizard should offer to auto-prefix DATA variable names with the table
> name, or display this guidance inline.

### Step T-2: Row classification

The wizard shows a mini-table with all rows in the range. For each row, press one key:

> ⚠️ **ADDITIONAL IDEA FOR CLAUDE:** I believe once a mini table is defined the first thing to ask the user is to define us a label for ecah column (the main purpose of the column) a and a descrition for it
> based on this the iterations below will always suggest something based on this definition and then the user doe snot need to remember if column N is meaning what, the user can follow the naming of the column
> it should be clear that in the wizard each possible mini-table is an object in itself, and needs to be ttreated as such

| Key | Role | Meaning |
|---|---|---|
| `H` | **Header** | A row of column titles (e.g. "ITEM NO.", "DESCRIPTION", "QTY"). The engine uses this to locate and verify the table. |
| `D` | **Data** | A row containing actual records to extract. There is usually one D row in the pattern representing all repeating records. |
| `F` | **Footer** | A summary row at the bottom (totals, sub-totals). The engine uses it to know where the data ends. |
| `S` | **Skip** | A structural row to consume but ignore (blank separator, etc.). |

Classify every row in the range, then press ENTER or the confirm key to proceed.
> ⚠️ **Rows clasification and JSON output:** The user needs to clasify all row types (H D F S) if they exist
> each mini table is a small matrix , and in general headers cells defines in a certain way the type of content that data cell will have
> footers do not follow  this pattern (headers do not follow it either, but the probablility of this happening is lower)
> all in all headers will in general be labels (if only one headr appears , or if more thatn one header is defined then for the latest header), if more than one header apperas it might happen that some of the header data is related to vvariables, and not labels, or then some of the header rows needs to be empty or ignored
> similar applies to footers

### Step T-3: Column definition

The wizard then asks you to define each column, one modal per column, for each row type
(H rows first, then D rows, then F rows).

Every column modal has the same four fields:

| Field | Default | Meaning |
|---|---|---|
| **Name** | Auto-suggested | See role prefixes below |
| **Type** | Auto-inferred | `string`, `number`, `date`, `boolean` — inferred from the Excel cell format |
| **Match / pattern** | Depends on role | For labels: the exact cell text. For variables: a regex (`.*` = accept anything). |
| **Notes** | (empty) | Free-text note saved to the session log |

#### Column roles — the Name field

Any column in any row type can play one of three roles. The role is controlled by the
value you type in the **Name** field:

| What you type | Role | Meaning |
|---|---|---|
| A plain name, e.g. `qty_label` | **Label** (default for H/F) | The cell is matched by the engine to verify the table structure. Never in the JSON output. |
| `var:item_no` | **Variable** | The cell value is extracted to the JSON output under the name `item_no`. |
| (empty) or `IGNORE` | **Ignored** | The column is consumed but neither verified nor extracted. |

In **DATA rows** the defaults are reversed: plain names are variables, `lbl:name` makes it a label.

| What you type in a D row | Role |
|---|---|
| A plain name, e.g. `item_no` | **Variable** (default for D) |
| `lbl:subtotal_label` | **Label** |
| (empty) or `IGNORE` | **Ignored** |

> **Example — invoice line items table (B11:F29):**
>
> Header row (H): ITEM NO. · DESCRIPTION · QTY · UNIT PRICE · TOTAL
> - Column B: `item_no_label` (label, matches "ITEM NO.")
> - Column C: `desc_label` (label, matches "DESCRIPTION")
> - Column D: `qty_label` (label, matches "QTY")
> - Column E: `unit_price_label` (label, matches "UNIT PRICE")
> - Column F: `total_label` (label, matches "TOTAL")
>
> Data row (D):
> - Column B: `items.item_no` (variable, type: string)
> - Column C: `items.description` (variable, type: string)
> - Column D: `items.qty` (variable, type: number)
> - Column E: `items.unit_price` (variable, type: currency)
> - Column F: `items.total` (variable, type: currency)
>
> Footer row (F):
> - Column B: IGNORE
> - Column C: IGNORE
> - Column D: IGNORE
> - Column E: `subtotal_label` (label, matches "SUBTOTAL")
> - Column F: `subtotal_value` (variable, type: currency)

### After defining all columns

The entire table range is highlighted in the grid (dim magenta for data rows, solid
magenta for anchor and header/footer columns). You can press `T` on any cell in the
range to reopen the editor and change definitions.

---

## The pattern file — what does it contain?
(the example below shows that csv format, this is correct, but only if the option for exporting CSV files is used, currently there is a **BUG** in this area, which need to be corrected )

The pattern file produced by `grepxcel wizard` has this structure:

```
config:,read.direction,LR
config:,currency.sign,€

lbl:,dept_label,string,DEPARTMENT:           ← label definitions
var:,inventory.department,string,.*           ← variable definitions

START:
cell:1,dept_label                             ← find the DEPARTMENT: label
cell:1,inventory.department                  ← extract the value next to it
table:*                                       ← begin repeating table block
,HEADER:1,item_no_label,desc_label,...       ← assert header row columns
,DATA:*,items.item_no,items.description,...  ← extract data row columns
,FOOTER:1,IGNORE,IGNORE,IGNORE,subtotal_lbl ← assert footer row
cell:1,comments_label
cell:1,inventory.comments
END:
```

### Key facts about the pattern format

- **`lbl:` rows** define labels: text the engine uses to find its place on the sheet.
  They are *never* in the JSON output. Each must have a **unique name** — if two labels
  share the same name, only one match text is remembered and the pattern will fail to
  match some cells correctly. ⚠️ The wizard should validate this and prevent duplicates.

- **`var:` rows** define variables: the values extracted to JSON. Dot notation creates
  nested JSON. In a table, the prefix becomes the array key.

- **`HEADER:1` rows** in the table block tell the engine to find a row that matches all
  the listed label texts before treating it as a valid table instance.

- **`FOOTER:1` rows** — similar to HEADER:1 but at the bottom of the table.
  ⚠️ **Current bug:** the wizard TUI emits `HEADER:1` for footer rows too. Should be
  `FOOTER:1`. Fix pending.

- **`DATA:*`** collects all rows between the HEADER and FOOTER (or until an empty row).

---

## Ending the session

Press `E` to end the session. The wizard shows a summary and saves the pattern file.

```
✓  Pattern written: pattern.csv
   Try:  grepxcel extract -p pattern.csv data.xlsx
         grepxcel validate-pattern pattern.csv
```

After saving, run `grepxcel validate-pattern pattern.csv` to check the pattern for errors
before trying an extraction.

---

## Session log

Every wizard session writes a log file to `logs/<basename>.wizard.<timestamp>.log`.
The log records every classification decision, note, and table definition. You can
review it to understand what happened and to spot mistakes.

Example log entries:
```
20:24:15  LABEL    B6   "CONTACT EMAIL"  →  contact_email_label  [string, CONTACT EMAIL]
20:24:24  VALUE    C6   (empty)          →  inventory.contact_email  [string, .*]
20:32:14  TABLE    B11:F29  name=items  mult=*  H=1 D=17 F=1 S=0  vars=[item_no, ...]
```

Notes you add with `F2` (or via the Notes field in any modal) also appear in the log:
```
19:54:57  NOTE  B12: The suggested name of the variable could come from the associated header
```

---

## Quick reference card

```
L = Label (anchor text, never in output)
V = Value to extract (goes to JSON)
C = Section header (skip, no output)
T = Table (opens 3-step flow)
I = Ignore this cell

Inside the T modal:
  H = Header row    D = Data row    F = Footer row    S = Skip row

Column name prefixes:
  plain name    → label (in H/F rows) or variable (in D rows)
  var:name      → always a variable
  lbl:name      → always a label
  IGNORE        → skip this column

Dot notation in variable names:
  po.number, po.date  →  {"po": {"number": ..., "date": ...}}
  items.qty           →  {"items": [{"qty": ...}, ...]}   (in a table)
```

---

## Known bugs and gaps (to fix after document review)

### Confirmed bugs (wrong behaviour, must fix)

| # | Where | Issue | Impact |
|---|---|---|---|
| B1 | Save | `-o pattern.xlsx` writes CSV text, not a real xlsx workbook | File opens with a warning in Excel; contents look like CSV |
| B2 | Pattern output | Footer rows are emitted as `HEADER:1` in the pattern instead of `FOOTER:1` | Engine treats footer as a second header; table matching may fail |
| B3 | L modal | Type field for Labels defaults to `string` — not inferred from Excel cell format | User must manually change type even when Excel clearly shows a date or number |
| B4 | Pattern output | Duplicate `lbl:` names (e.g. three columns all named `empty`) cause silent wrong matches | Engine resolves by name; only the first definition is used for all three columns |

### Missing features (behaviour that should exist but doesn't)

| # | Where | Gap | Expected behaviour |
|---|---|---|---|
| F1 | Template mode — `_advance()` | After classifying a cell as **L** (label), the cursor jumps to the next *non-empty* cell. In template mode (blank form) the value cell is usually the cell immediately adjacent to the label (next right in LR, next down in TD) — but that cell is empty, so the cursor skips past it. | After L in template mode: advance one cell in the configured direction (even if empty), so the user lands on the value cell automatically. |
| F2 | Table name | The table name entered in Step T-1 is stored in the session log but **not used in the pattern**. The user expects it to define the JSON output key (e.g. name `items` → `{"items": [...]}`) | Auto-prefix all DATA variable names with `<table_name>.` when writing the pattern — e.g. variable `item_no` with table name `items` becomes `items.item_no` in the pattern. If the user already typed a dot-prefix that starts with the table name, do not double-prefix. |
| F3 | V / table modals — Match field | No presets offered for common regex patterns; non-technical users must know regex | Offer a small preset menu or hint line below the Match field: `.*` = any, email, URL, integer, decimal, date (ISO), phone. Selecting one fills the field. |
| F4 | Step T-2 row classification | No way to mark a row as a **Splitter** (blank separator between header and data, or data and footer) | Add `P` key (or `X`) for SPLITTER — a row where all cells must be blank. The pattern emits `SPLITTER:1` for it. |
| F5 | Step T-3 column modals | No column legend: the user must remember what each column number means while filling in per-column modals | After Step T-2, show a brief column-labelling step where the user gives each column a short title/description. These titles appear as context hints in all subsequent per-column modals. |
| F6 | Navigation keys | The key bindings (L, V, T, I, R, G, U, E…) are not prominently displayed while classifying cells | Show a persistent key-reference line or panel section listing all active keys and their action at all times. |
| F7 | Wizard | No duplicate label-name warning: two labels with the same name silently produce a broken pattern | Before saving, warn if any two `lbl:` entries share the same name and prompt the user to rename. |
