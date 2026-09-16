# grepxcel

[![PyPI](https://img.shields.io/pypi/v/grepxcel?color=blue)](https://pypi.org/project/grepxcel/)
[![Python](https://img.shields.io/pypi/pyversions/grepxcel)](https://pypi.org/project/grepxcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/scpg/grepxcel/blob/main/LICENSE)
[![Downloads](https://img.shields.io/pypi/dm/grepxcel)](https://pypi.org/project/grepxcel/)

**Stop writing custom Python for every Excel template.** Define the layout once in a pattern spreadsheet, extract clean JSON from any file that follows it.

```
pip install grepxcel
```

**Capabilities at a glance:**
- **Pattern file** — describe any spreadsheet's layout once; reuse it across all matching files with no code changes
- **Anchor-based** — finds data even when row or column positions shift between files
- **LLM drafter** — auto-generate a starter pattern from any file in ~30 seconds (`grepxcel draft file.xlsx`)
- **Visual wizard** — browser UI: click cells to classify, download the pattern (`grepxcel web-wizard file.xlsx`)
- **CI-ready** — deterministic, `--strict` exit-2 mode, JSON / CSV / colored-Excel output, Python 3.11–3.15
- **Security-hardened** — ReDoS guard, ZIP-bomb protection, CSRF hardening on the web wizard
- **MCP server** — expose grepxcel as a tool in Claude, Cursor, and any MCP-compatible AI agent (`pip install 'grepxcel[mcp]'`)

---

## The problem

You get Excel files from suppliers, clients, or other departments. Every template is different. You write Python to parse each one — and it silently breaks the moment someone moves a column or renames a header.

## The solution

Describe the layout once in a **pattern file**. Run `grepxcel extract`. Get clean, structured JSON. When a file doesn't match the pattern, grepxcel flags exactly what went wrong — it is designed to fail loudly rather than produce silent errors, so automated pipelines can catch problems early.

```
pattern.xlsx  +  data.xlsx  →  { "po": { "number": "PO-2026" }, "line": [ {…} ] }
```

## See it in action (2 minutes)

**1. Generate example files:**

```bash
grepxcel generate-examples
cd grepxcel-examples/01_simple_invoice
```

**2. Look at what you have** — a pattern file and a data file:

```
pattern.xlsx:  "Here is what the invoice looks like"
data.xlsx:     "Here is an actual invoice"
```

**3. Extract:**

```bash
grepxcel extract -p pattern.xlsx data.xlsx
```

**4. Get clean JSON:**

```json
{
  "inv":    { "number": "AB123456", "date": "2026-03-15" },
  "client": { "name": "Alice Wonderland", "email": "alice@wonderland.example" },
  "amount": { "net": 120, "vat": 24, "gross": 144 }
}
```

That's it. Same pattern works on next month's invoice, and the one after that.

## Who is this for?

- **Data teams** drowning in Excel reports from different sources
- **Finance and accounting** — extract invoices, timesheets, price lists automatically
- **ETL pipelines** — replace brittle `openpyxl` scripts that break when a column shifts
- **Research and clinical data** — reject files that deviate from the expected structure before they corrupt your dataset

You don't need to be a programmer to use grepxcel.

**Non-developers:** use the browser-based web wizard (`pip install "grepxcel[web]"` then `grepxcel web-wizard data.xlsx`) — click cells in the browser to classify them, download the pattern, and run the extraction. No coding required after install. See [docs/web-wizard-guide.md](https://github.com/scpg/grepxcel/blob/main/docs/web-wizard-guide.md) for a walkthrough.

## Install

```bash
pip install grepxcel
```

That's all you need to extract data. Optional extras add more capabilities:

| You want to… | Install |
|---|---|
| Extract data from Excel files | `pip install grepxcel` |
| Get results as pandas DataFrames | `pip install 'grepxcel[pandas]'` |
| Get results as polars DataFrames | `pip install 'grepxcel[polars]'` |
| Auto-draft patterns with a local AI model | `pip install 'grepxcel[suggest]'` |
| Auto-draft patterns with a cloud model | `pip install 'grepxcel[draft-cloud]'` |
| Use grepxcel as an MCP server for AI agents | `pip install 'grepxcel[mcp]'` |

Requires Python **3.11+**. Works on Linux, macOS, and Windows.

<details>
<summary>Installing from source (for contributors)</summary>

```bash
git clone https://github.com/scpg/grepxcel.git
cd grepxcel
python -m venv .venv
.venv/bin/pip install -e .
```

Windows (PowerShell): use `.venv\Scripts\` instead of `.venv/bin/`.

</details>

## How patterns work

A pattern file is a simple spreadsheet (`.xlsx` or `.csv`) with four types of rows:

| Row type | What it does |
|---|---|
| **Settings** (`config:`) | How to read the file — scan direction, currency symbol, case sensitivity |
| **Landmarks** (`lbl:`) | Text you use to find your place — "Invoice Number", "Total" |
| **Fields** (`var:`) | Values you want to extract — the invoice number, the date, the amount |
| **Notes** (`doc:`) | Your own comments — ignored by the engine |

Between `START:` and `END:` you tell grepxcel the order to read cells:

```
START:
  cell:1  IGNORE           ← skip this cell (it's a label)
  cell:1  inv.number       ← read the next cell into "inv.number"
  cell:1  IGNORE
  cell:1  inv.date
END:
```

Dot notation creates nested output: `inv.number` → `{"inv": {"number": …}}`.

> **Full reference:** [docs/pattern-file.md](https://github.com/scpg/grepxcel/blob/main/docs/pattern-file.md) covers every
> instruction, type, and config option in detail.

## Describing what cells look like

Each field can have a **match pattern** that describes what the cell value should look like. You don't need to know regular expressions — here's what most people need:

| You want to match… | Write this | Example |
|---|---|---|
| Anything at all | *(leave blank)* | any value |
| A whole number | `\d+` | `7`, `2026` |
| Exactly 4 digits | `\d{4}` | `2026` |
| A code like `PO-2026` | `PO-\d+` | `PO-1`, `PO-999` |
| One of a few words | `paid\|unpaid` | `paid` |
| A price like `19.99` | `\d+\.\d{2}` | `19.99` |

**Tips:** Start loose (blank = accept anything), run `extract`, see what comes out, then tighten. Add `config: | ignore.case | yes` to match regardless of capitalization.

> **Under the hood** these are standard Python regular expressions. If you
> already know regex, anything from the
> [Python `re` syntax](https://docs.python.org/3/library/re.html) works.

### Optional quality gates (column A modifiers)

Column A supports **order-independent modifiers** after the `var:` or `lbl:` keyword:

| Column A | What it does |
|---|---|
| `var:not-null` | **Fatal error** if value is empty — always, not just with `--strict` |
| `var:glob` | Column D is a shell glob (`PROD-*` matches `PROD-42`) |
| `var:literal` | Column D is an exact string — no regex escaping needed |
| `lbl:not-null` | Fatal error if the anchor label is missing |

Modifiers combine freely: `var:not-null:glob`, `var:literal:not-empty`, etc.

```
var:not-null   invoice.number   string   INV-\d+   # required regex field
var:glob       sku              string   PROD-*     # glob match
var:literal    status           string   Active     # exact string
```

## CLI commands

```bash
# Extract data
grepxcel extract -p pattern.xlsx data.xlsx
grepxcel extract -p pattern.xlsx data.xlsx -o output/     # write JSON to a directory
grepxcel extract -p pattern.xlsx data.xlsx --format csv -o output/   # CSV (single-table)
grepxcel extract -p pattern.xlsx data.xlsx --format xlsx -o output/  # colored Excel report
grepxcel extract -p pattern.xlsx data.xlsx --strict        # fail on any missing field
grepxcel extract -p pattern.xlsx data_dir/ -r              # process a whole directory

# Tools
grepxcel generate-examples                     # create example files to try
grepxcel validate-pattern pattern.xlsx         # check a pattern without extracting
grepxcel lint data.xlsx                        # inspect a file before extraction
grepxcel schema pattern.xlsx -o schema.json    # generate JSON Schema for validation
grepxcel draft data.xlsx                       # AI-draft a starter pattern
grepxcel docs                                  # generate the pattern reference file
grepxcel quickstart                            # guided tutorial in your terminal
grepxcel doctor                                # check your environment is ready

# For AI agents
grepxcel mcp                                   # start MCP server
grepxcel mcp-config                            # print config for Claude / Cursor
```

Run `grepxcel <command> --help` for all options.

## Python API

```python
import grepxcel

# Get a dict (always available)
result = grepxcel.extract("pattern.xlsx", "data.xlsx")
print(result["po"]["number"])        # "PO-2026"

# Flat tables — skip the {"_source": ..., "data": [...]} envelope
result = grepxcel.extract("pattern.xlsx", "data.xlsx", flat_tables=True)
for row in result["line"]:          # each row dict directly, no unwrapping
    print(row["item"])

# Get DataFrames (pip install 'grepxcel[pandas]' or 'grepxcel[polars]')
frames = grepxcel.extract_df("pattern.xlsx", "data.xlsx")
frames["line"]                       # pandas or polars DataFrame of all table rows
frames["_scalars"]                   # one-row DataFrame of scalar fields
```

## Output format

Fields with dot notation are grouped into nested objects. Tables produce arrays.

```json
{
  "po":     { "number": "PO-2026", "date": "2026-05-01" },
  "vendor": { "name": "Acme Supplies" },
  "line": [
    {
      "data": [
        { "item": "Laptop", "qty": 2, "price": 1200.00, "total": 2400.00 },
        { "item": "Dock",   "qty": 6, "price":   75.00, "total":  450.00 }
      ],
      "footer": { "label": "Grand Total", "value": 3030.00 }
    }
  ]
}
```

Each table instance is wrapped with `{"data": [...], "footer": {...}}` so per-instance
metadata (source location, header/footer subtotals) survives. If you just want the rows,
use `--no-source` on the CLI or `flat_tables=True` in the Python API — the wrapper is
removed and each table key maps directly to a list of row dicts.

## What's new in v0.3.0

- **Browser-based wizard** — `grepxcel web-wizard file.xlsx` opens a point-and-click UI to classify cells visually
- **Security hardening** — CSRF protection on the web wizard, config validation, ReDoS guard on custom match patterns, narrowed exception handling
- **Preload fixes** — footer rows with no label columns now pre-populate correctly in the wizard

Full history: [CHANGELOG.md](https://github.com/scpg/grepxcel/blob/main/CHANGELOG.md)

## Learn more

| Topic | Where to look |
|---|---|
| **Web wizard guide** (mouse-driven UI, non-developers) | [docs/web-wizard-guide.md](https://github.com/scpg/grepxcel/blob/main/docs/web-wizard-guide.md) |
| Pattern file reference (all instructions, types, config) | [docs/pattern-file.md](https://github.com/scpg/grepxcel/blob/main/docs/pattern-file.md) |
| Advanced patterns: bounded rows `DATA:{n,m}`, `SKIP_IF` | [docs/pattern-file.md#bounded-and-sparse-rows](https://github.com/scpg/grepxcel/blob/main/docs/pattern-file.md) |
| Merged cells — how grepxcel handles them | [docs/pattern-file.md#merged-cells](https://github.com/scpg/grepxcel/blob/main/docs/pattern-file.md#merged-cells) |
| Engineering philosophy and design principles | [docs/MINDSET.md](https://github.com/scpg/grepxcel/blob/main/docs/MINDSET.md) |
| Draft command evaluation (model quality, benchmarks) | [docs/EVALUATION.md](https://github.com/scpg/grepxcel/blob/main/docs/EVALUATION.md) |
| CLI flags and options (full reference) | [docs/cli-reference.md](https://github.com/scpg/grepxcel/blob/main/docs/cli-reference.md) |
| Security model and vulnerability reporting | [SECURITY.md](https://github.com/scpg/grepxcel/blob/main/SECURITY.md) |
| Contributing guidelines | [CONTRIBUTING.md](https://github.com/scpg/grepxcel/blob/main/CONTRIBUTING.md) |

## Support

grepxcel is free and open source. If it saves you time and you'd like to say thanks,
you can [buy me a coffee](https://buymeacoffee.com/scpg.dev) — entirely optional and
always appreciated.

<a href="https://buymeacoffee.com/scpg.dev" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" height="41" width="174"></a>

## License

MIT — see [LICENSE](https://github.com/scpg/grepxcel/blob/main/LICENSE).
