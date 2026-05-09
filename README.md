# grepxcel

> Pattern-based data extraction for Excel — define what to look for, run, get structured data.

---

## What it does

You describe the layout of your Excel sheet in a **pattern file** (itself an Excel file). The engine reads any matching data file and extracts cells and tables into clean, structured output — no coding required to define new patterns.

Think of it as *grep for Excel*.

---

## How it works

```
pattern.xlsx  +  data.xlsx  →  { cells: {…}, tables: [{…}] }
```

The pattern file has three sections:

| Section | Purpose |
|---|---|
| `config:` | Read direction, currency symbol, empty-cell aliases |
| `def:` | Field registry — name, type (`string`, `integer`, `currency`, `date`), and a validation regex |
| `START:` … `END:` | Extraction sequence — `cell:1` for single cells, `table:*` for repeating mini-tables |

The engine scans the data file in the configured direction, matches every instruction in order, and returns all extracted values.

---

## Quick start

```bash
git clone https://github.com/scpg/grepxcel.git
cd grepxcel
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

```python
from engine import Engine, Logger, VerbosityLevel

logger = Logger(level=VerbosityLevel.NORMAL)
result = Engine().process("pattern.xlsx", "data.xlsx", logger=logger)

print(result["cells"])
print(result["tables"])
```

---

## Verbosity levels

| Level | What you see |
|---|---|
| `QUIET` | Nothing — records still collected in memory |
| `NORMAL` | Summary + all validation warnings with hints *(default)* |
| `VERBOSE` | + every cell and table match step by step |
| `DEBUG` | + every anchor probe and rejection reason |

---

## Project layout

```
engine/          ← importable Python package
tests/
  fixtures/      ← pattern + data xlsx pairs (one folder per scenario)
  unit/          ← pytest unit tests
  integration/   ← pytest integration tests
samples/         ← reference pattern files
samples.local/   ← local-only files, never synced  (gitignored)
logs/            ← log file output                 (gitignored)
output/          ← JSON extraction results         (gitignored)
```

---

## Running the tests

```bash
pytest
```

91 tests across 3 fixture scenarios — all green.

---

## Requirements

- Python 3.10+
- openpyxl
