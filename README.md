# grepxcel

> Pattern-based data extraction for Excel — define what to look for, run, get structured JSON.

![CI](https://github.com/scpg/grepxcel/actions/workflows/ci.yml/badge.svg)

---

## What it does

You describe the layout of your Excel sheet in a **pattern file** — either an Excel workbook (`.xlsx`) or a plain `.csv` (handy for hand-editing and git diffs). The engine reads any matching data file and extracts cells and tables into clean, hierarchical JSON — no coding required to define new patterns.

Think of it as *grep for Excel*.

---

## How it works

```
pattern.xlsx  +  data.xlsx  →  { "po": { "number": "PO-2026" }, "line": [ {…} ] }
```

The pattern file has four row types:

| Row type | Purpose |
|---|---|
| `config:` | Global settings: read direction, currency symbol, empty-cell aliases, case-insensitive matching (`ignore.case`) |
| `lbl:` | Anchor label — matched for position, **never written to output JSON** |
| `var:` | Data field — extracted and written to output JSON |
| `doc:` | Comment / documentation row — ignored by the engine |

Between `START:` and `END:` you list the extraction sequence:

- `cell:next fieldName` — read the next non-empty cell into a field (alias: `cell:1`)
- `cell:B5 fieldName` — jump directly to an absolute cell (A1-notation reference)
- `table:*` — match all instances of a repeating mini-table block

Dot notation in `var:` field names creates nested output: `po.number` → `{"po": {"number": …}}`.

### Pattern file formats

A pattern can be authored as **`.xlsx`** or **`.csv`** — both are read into the
same internal grid, so they behave identically. CSV is convenient for
hand-editing and produces clean git diffs. When writing CSV:

- One pattern row per CSV line; column A is the keyword (`config:`, `lbl:`, `var:`, `cell:…`).
- Table-template rows start with an **empty first field** (blank column A), e.g. `,HEADER:1,col_a,col_b`.
- **Quote any regex containing a comma**, e.g. `var,line.qty,integer,"\d{1,3}"`.
- Plain text only — formulas (a leading `=`) are rejected, exactly as in `.xlsx`.

---

## Writing the match pattern (regex) — the simple version

The last column of a `lbl:` or `var:` row is a **regex**: a little pattern that
describes *what the cell should look like*. You don't need to know regex to start —
here's everything most people need.

**Two rules to remember:**

1. **The whole cell must match.** If your pattern is `\d{4}` (four digits), the cell
   `2026` matches but `2026-05` does **not** (the `-05` is left over). You don't add
   `^` or `$` anchors — grepxcel does that for you.
2. **Leave it blank to accept anything.** No pattern in the last column = "any value
   is fine". For `date`/`datetime` fields, always leave it blank (the type is checked,
   not the text).

**The building blocks you'll actually use:**

| You want to match… | Write this | Matches |
|---|---|---|
| Any value at all | *(leave blank)* or `.*` | anything |
| A whole number | `\d+` | `7`, `2026`, `100` |
| Exactly 4 digits | `\d{4}` | `2026` (not `26`) |
| Between 1 and 3 digits | `\d{1,3}` | `5`, `42`, `999` |
| Letters only | `[A-Za-z]+` | `Acme` |
| A code like `PO-2026` | `PO-\d+` | `PO-1`, `PO-2026` |
| One of a few words | `paid\|unpaid\|pending` | `paid` |
| A price like `19.99` | `\d+\.\d{2}` | `19.99` |

Cheat-sheet: `\d` = a digit, `[A-Z]` = one capital letter, `+` = "one or more",
`{4}` = "exactly four", `{1,3}` = "between one and three", `.` = any character,
`.*` = "anything", `|` = "or".

**Tips for beginners:**

- **Start loose, then tighten.** Begin with a blank pattern (accept anything), run
  `extract`, see what comes out, then add a pattern only where you need to be strict.
- **Use `[A-Z]+`, not `([A-Z])+`.** Both look similar, but the bracket form is faster
  and the parentheses-with-a-`+` form is rejected by grepxcel as unsafe (it can make
  matching hang). The tool will tell you if you hit this.
- **Case doesn't matter?** Add one config row at the top of the pattern file:
  `config: | ignore.case | yes`. Then `PAID`, `Paid`, and `paid` all match the same
  pattern. (Default is case-sensitive.)
- **CSV pattern files:** if your pattern contains a comma (like `\d{1,3}`), wrap it in
  quotes — `"\d{1,3}"` — so the comma isn't read as a new column.

> Under the hood these are standard [Python `re`](https://docs.python.org/3/library/re.html)
> patterns, so anything from that syntax works if you already know regex.

---

## Quick start

### Install

```bash
git clone https://github.com/scpg/grepxcel.git
cd grepxcel
python -m venv .venv
.venv/bin/pip install -e .                            # core (extract + docs)
.venv/bin/python3 scripts/install_llm_deps.py         # optional: local draft (auto-detects GPU)
```

**What you actually need** — grepxcel ships in layers, so you only install what you use:

| You want to… | Install |
|---|---|
| Extract data / generate the docs reference | `pip install grepxcel` *(core — nothing extra)* |
| Draft patterns with the **local** model (offline) | `pip install 'grepxcel[suggest]'` |
| Draft patterns with a **cloud** model (Claude) | `pip install 'grepxcel[draft-cloud]'` |

The `extract` and `docs` commands work with the core install alone. The `draft`
command is optional; if its dependencies are missing, grepxcel tells you exactly
what to install (and points you at the cloud option as an alternative).

> **Windows (PowerShell):** use `.venv\Scripts\` instead of `.venv/bin/`, e.g.
> `.venv\Scripts\pip install -e .` and `.venv\Scripts\grepxcel ...`.

Requires Python **3.11+**.

### CLI usage

```bash
# Extract data from an Excel file using a pattern
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx

# Write output to a directory instead of stdout
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx -o output/

# Process a specific sheet, or every sheet at once
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx --sheet 2025
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx --all-sheets

# Verbose mode (step-by-step match log)
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx -v

# Legacy flat output format ({"cells":{}, "tables":[]})
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx --format legacy

# Generate a colour-coded pattern reference file
.venv/bin/grepxcel docs -o pattern-reference.xlsx

# Use a local LLM to draft a starter pattern for an unseen Excel file
.venv/bin/grepxcel draft data.xlsx -o draft-pattern.xlsx
```

### Python API

```python
from grepxcel import Engine, Logger, VerbosityLevel

logger = Logger(level=VerbosityLevel.NORMAL)
result = Engine().process("pattern.xlsx", "data.xlsx", logger=logger)

# Fields are grouped by dot-notation prefix
print(result["po"]["number"])        # "PO-2026"
print(result["vendor"]["name"])      # "Acme Supplies"

# Tables are arrays of instance objects
for row in result["line"][0]["data"]:
    print(row["item"], row["qty"])
```

---

## Output format

Fields defined with dot notation (`po.number`, `po.date`) are grouped into nested objects.
Tables always produce an array of instance objects, each containing `data`, and optionally
`header` and `footer` sections.

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

Label fields (`lbl:`) are used only for positional anchoring and are never included in output.

---

## CLI reference

### `grepxcel extract`

| Flag | Default | Purpose |
|---|---|---|
| `-p FILE` | required | Pattern file — `.xlsx` or `.csv` |
| `--format` | `nested` | Output format: `nested` (default) or `legacy` |
| `-o DIR` | — | Write JSON to directory (stdout if omitted) |
| `-l FILE` | — | Append structured log to file |
| `-v` / `-vv` | off | Verbosity: step-by-step / anchor probes |
| `-d` / `--debug` | off | Same as `-vv` |
| `--max-size MB` | 5 | Compressed file size limit |
| `--max-uncompressed MB` | 50 | Uncompressed ZIP content limit (ZIP bomb guard) |
| `--max-cell-len N` | 1000 | Max cell chars fed to regex (ReDoS guard) |
| `--sheet NAME_OR_INDEX` | active | Sheet name or 0-based index to process |
| `--all-sheets` | off | Process every sheet; output is a dict keyed by sheet name (mutually exclusive with `--sheet`) |

### `grepxcel docs`

| Flag | Default | Purpose |
|---|---|---|
| `-o FILE` | `pattern-reference.xlsx` | Output path for the reference file |

### `grepxcel draft`

Drafts a starter pattern file for an unseen Excel file using an LLM. The output is
a *starting point* — review and refine the generated regexes before use.
`grepxcel suggest` is a backward-compatible alias for this command.

| Flag | Default | Purpose |
|---|---|---|
| `-o FILE` | `draft_pattern.xlsx` | Write draft pattern to this path |
| `-v` | off | Print the Excel analysis sent to the model + update status |
| `--dry-run` | off | Print the analysis that would be sent to the model, then exit (no inference) |
| `--backend local\|claude\|gemini` | `local` | Inference backend (see below) |
| `--sheet NAME_OR_INDEX` | active | Sheet to analyse |
| `--max-size MB` | 5 | Compressed file size limit |
| `--max-uncompressed MB` | 50 | Uncompressed ZIP content limit (ZIP bomb guard) |

#### Backends

| Backend | Install | Notes |
|---|---|---|
| `local` *(default)* | `python3 scripts/install_llm_deps.py` | Runs Qwen2.5-Coder-7B (GGUF) in-process. **No data leaves your machine.** Model is pinned to a revision and downloaded once (~4.7 GB); set `GREPXCEL_MODEL_AUTOUPDATE=1` to track upstream, `GREPXCEL_MODEL_DIR` to relocate the cache. |
| `claude` | `pip install -e '.[draft-cloud]'` | Anthropic API. Requires `ANTHROPIC_API_KEY`. Prints a one-line privacy notice and per-call token cost. |
| `gemini` | — | **Planned for a future release** — not yet available. Selecting it prints a notice and exits. |

> **Faster first download (`local` backend):** the model is fetched once from
> HuggingFace. Anonymous downloads work but are rate-limited and can be slow or
> stall on the ~4.7 GB file. For a faster, more reliable first run, set a free
> [HuggingFace token](https://huggingface.co/settings/tokens) (read scope):
> `HF_TOKEN=hf_... grepxcel draft data.xlsx`. It's optional and one-time — the
> model is cached afterwards. If a download stalls, grepxcel prints this tip too.

> **Privacy:** the `claude` cloud backend sends only the *structure description*
> of your sheet (column types, sample values, labels) — never the raw file. The
> `local` backend sends nothing over the network during inference.

```bash
grepxcel draft data.xlsx                       # local model (default)
grepxcel draft data.xlsx --dry-run             # inspect the analysis, no inference
ANTHROPIC_API_KEY=sk-... grepxcel draft data.xlsx --backend claude
```

#### Model cache & offline / alternative downloads

The local model is stored once in the platform-appropriate per-user cache
(resolved with [`platformdirs`](https://pypi.org/project/platformdirs/), the same
convention pip uses):

| OS | Default cache |
|---|---|
| Linux | `$XDG_CACHE_HOME/grepxcel/models/` (default `~/.cache/grepxcel/models/`) |
| macOS | `~/Library/Caches/grepxcel/models/` |
| Windows | `%LOCALAPPDATA%\grepxcel\Cache\models\` |

Override the location with `GREPXCEL_MODEL_DIR` (handy for Docker volumes or a
shared model dir) — it takes precedence over the defaults above.

If the HuggingFace download is slow or blocked, you don't have to let grepxcel
fetch it — **grepxcel only downloads when the file isn't already in the cache**,
so you can supply it yourself from any source:

```bash
mkdir -p ~/.cache/grepxcel/models            # or your $GREPXCEL_MODEL_DIR
# download the GGUF anywhere (HF website, a mirror, ModelScope, …), then place it
# with this EXACT name so grepxcel finds it and skips the download:
mv Qwen2.5-Coder-7B-Instruct-Q4_K_M.gguf  ~/.cache/grepxcel/models/
grepxcel draft data.xlsx                     # uses the local file — no download
```

Other options:

- **Mirror:** `huggingface_hub` honors `HF_ENDPOINT`, e.g.
  `HF_ENDPOINT=https://hf-mirror.com grepxcel draft data.xlsx` (third-party mirror).
- **Token:** `HF_TOKEN=hf_...` removes the anonymous rate limit (fastest fix).

> ⚠️ The normal HuggingFace download is **hash-verified** against the pinned
> revision. A **manually-placed file is not verified** by grepxcel — it trusts
> whatever is in the cache, so make sure your source is trustworthy.

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
grepxcel/        ← importable Python package (engine, parser, models, security, cli, drafter)
scripts/         ← reusable utility scripts for contributors
tests/
  fixtures/      ← pattern + data xlsx pairs (one folder per scenario)
  unit/          ← pytest unit tests
  integration/   ← pytest integration tests
docs/            ← additional documentation
tmp.local/       ← throwaway scripts, never synced  (gitignored)
samples.local/   ← local-only Excel files, never synced  (gitignored)
logs/            ← log file output                 (gitignored)
output/          ← JSON extraction results         (gitignored)
```

---

## Running the tests

```bash
.venv/bin/pytest tests/ -q
```

750+ unit and integration tests across 16 fixture scenarios — all green.

---

## Requirements

- Python 3.11+
- `openpyxl >= 3.1`
- `defusedxml >= 0.7`
- `llama-cpp-python >= 0.2.90` and `huggingface_hub >= 0.23` — only for `grepxcel draft`

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). All PRs target the `dev` branch.

## Security

See [SECURITY.md](SECURITY.md) for how to report vulnerabilities privately.

## License

MIT — see [LICENSE](LICENSE).
