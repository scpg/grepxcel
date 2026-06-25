# CLI Reference

Full reference for all `grepxcel` commands and flags. For a quick overview,
see the [README](../README.md).

---

## `grepxcel extract`

Extract data from Excel files using a pattern file.

```bash
grepxcel extract -p pattern.xlsx data.xlsx
grepxcel extract -p pattern.xlsx data.xlsx -o output/
grepxcel extract -p pattern.xlsx jan.xlsx feb.xlsx mar.xlsx
grepxcel extract -p pattern.xlsx data_dir/
grepxcel extract -p pattern.xlsx data_dir/ -r
grepxcel extract -p pattern.xlsx data.xlsx --strict
grepxcel extract -p pattern.xlsx data.xlsx --sheet Sheet2
grepxcel extract -p pattern.xlsx data.xlsx --all-sheets
grepxcel extract -p pattern.xlsx data.xlsx -v
grepxcel extract -p pattern.xlsx data.xlsx --format legacy
grepxcel extract -p pattern.xlsx data.xlsx --format csv -o out/
grepxcel extract -p pattern.xlsx data.xlsx --format xlsx -o out/
grepxcel extract -p pattern.xlsx data.xlsx --meta
grepxcel extract -p pattern.xlsx data.xlsx --log logs/run.log
grepxcel extract -p pattern.xlsx data.xlsx --log run.log --log-format json
```

| Flag | Default | Purpose |
|---|---|---|
| `-p, --pattern FILE` | required | Pattern file (`.xlsx` or `.csv`) |
| `FILE_OR_DIR...` (positional) | required | Data `.xlsx` files or directories |
| `-r, --recursive` | off | Recurse into subdirectories when a directory is given |
| `-o, --output DIR` | — | Write JSON to directory (stdout if omitted) |
| `--strict` | off | Exit with code 2 if any field has issues (missing, validation mismatch); lists affected fields on stderr |
| `-v` / `-vv` | off | Verbosity: per-field trace / anchor probes |
| `-d, --debug` | off | Same as `-vv` |
| `-l, --log FILE` | — | Append structured log to file |
| `--log-format` | `text` | Log format: `text` (human-readable) or `json` (NDJSON for SIEM/cloud — never contains cell values) |
| `--format {nested,legacy,csv,xlsx}` | nested | Output shape — see [Output formats](#output-formats) |
| `--meta` | off | Add `_meta` block (run_id, stats, issues) for pipeline auto-verification |
| `--sheet NAME_OR_INDEX` | active | Sheet name or 0-based index to process |
| `--all-sheets` | off | Process every sheet; output keyed by sheet name |
| `--max-files N` | 10,000 | Safety cap on total files to process |
| `--max-rows N` | 2048 | Max data-sheet rows (Excel max 1,048,576) |
| `--max-columns N` | 1024 | Max data-sheet columns (Excel max 16,384) |
| `--max-cell-len N` | 1000 | Max cell chars fed to regex matching (ReDoS guard) |
| `--max-size MB` | 5 | Compressed file size limit |
| `--max-uncompressed MB` | 50 | Uncompressed ZIP content limit (prevents ZIP bomb attacks) |

### Exit codes

| Code | Meaning |
|---|---|
| 0 | Extraction succeeded, all fields valid |
| 1 | Extraction completed with warnings or errors |
| 2 | `--strict`: fields had issues · or an output-format guard rejected the request (see below) |

### Output formats

`--format` controls the shape of the extracted output:

| Value | Output | Notes |
|---|---|---|
| `nested` (default) | JSON, dot-notation grouped into nested objects; tables as arrays | stdout or `-o` directory |
| `legacy` | JSON `{"cells": {}, "tables": []}` (flat internal shape) | stdout or `-o` directory |
| `csv` | Flat CSV: one row per table data row; scalars denormalized as repeated columns | **single-table patterns only** |
| `xlsx` | Colored Excel report for human review (scalars, then each table top-down) | **requires `-o`** |

**`csv` rules**
- Refused (exit 2) when the pattern has more than one `table:` block — a flat CSV can't represent multiple tables. Use `nested`.
- Only table `data` rows are exported; per-table `header`/`footer` fields (e.g. subtotals) are omitted. Use `nested` or `grepxcel.extract_df()` to reach them.
- The `_meta` block (`--meta`) is excluded from CSV columns.

**`xlsx` rules**
- Requires `-o DIR` (binary Excel can't go to stdout).
- Never overwrites a source file: if the computed output path equals an input path, grepxcel exits 2. Use a separate `-o` directory.

**Both `csv` and `xlsx`**
- Do not support `--all-sheets` (a flat file can't hold multiple sheets) — exit 2. Use `--sheet` or `--format nested`.
- Every cell value is neutralized against formula/CSV injection (CWE-1236): a value extracted from an untrusted file that begins with `=`, `+`, `-`, `@`, tab, or carriage-return is written as literal text, not an executable formula.

### Verbosity levels

| Level | What you see |
|---|---|
| *(default)* | Summary + validation warnings with hints |
| `-v` | + every cell and table match step by step |
| `-vv` or `-d` | + every anchor probe and rejection reason |

---

## `grepxcel validate-pattern`

Check that a pattern file is valid — without running an extraction.

```bash
grepxcel validate-pattern pattern.xlsx
grepxcel validate-pattern pattern.csv -v
grepxcel validate-pattern a.xlsx b.csv
```

| Flag | Purpose |
|---|---|
| `FILE...` | One or more pattern files to validate |
| `-v, --verbose` | Print parsed config, fields, and extraction sequence |

---

## `grepxcel quickstart`

Print a guided tutorial in your terminal — what grepxcel does, how to create your
first pattern, and how to run your first extraction.

```bash
grepxcel quickstart
```

---

## `grepxcel generate-examples`

Create 4 ready-to-run example sets (pattern + data + README) in a local directory.

```bash
grepxcel generate-examples
grepxcel generate-examples -o my-examples
```

| Flag | Default | Purpose |
|---|---|---|
| `-o DIR` | `./grepxcel-examples/` | Directory to create |

---

## `grepxcel docs`

Generate a colour-coded pattern reference spreadsheet.

```bash
grepxcel docs
grepxcel docs -o reference/pattern-reference.xlsx
```

| Flag | Default | Purpose |
|---|---|---|
| `-o FILE` | `pattern-reference.xlsx` | Output path |

---

## `grepxcel lint`

Inspect an Excel data file before extraction.

```bash
grepxcel lint data.xlsx
grepxcel lint jan.xlsx feb.xlsx
```

Checks: file format, ZIP integrity, encryption/IRM, Microsoft Information
Protection labels, sheet dimensions (declared vs real), merged cells, formula
cells, empty sheets.

---

## `grepxcel schema`

Generate a JSON Schema (draft 2020-12) describing the extraction output for a pattern.

```bash
grepxcel schema pattern.xlsx
grepxcel schema pattern.xlsx -o schema.json
```

| Flag | Default | Purpose |
|---|---|---|
| `-o FILE` | — | Write schema to file (stdout if omitted) |

---

## `grepxcel draft`

Draft a starter pattern file for an unseen Excel file using an AI model.

```bash
grepxcel draft data.xlsx
grepxcel draft data.xlsx -o my-pattern.xlsx
grepxcel draft data.xlsx --backend claude
grepxcel draft data.xlsx --dry-run
```

| Flag | Default | Purpose |
|---|---|---|
| `-o FILE` | `draft_pattern.xlsx` | Output path |
| `-v` | off | Show the analysis sent to the model |
| `--dry-run` | off | Print analysis only, no inference |
| `--backend` | `local` | Backend: `local`, `claude`, `github`, `server`, `gemini` |
| `--sheet` | active | Sheet to analyse |

### Backends

| Backend | Install | Key needed | Notes |
|---|---|---|---|
| `local` | `pip install 'grepxcel[suggest]'` | — | Runs offline. ~5 GB model download on first use. No data leaves your machine. |
| `github` | `pip install 'grepxcel[draft-cloud]'` | `GITHUB_TOKEN` | Free with GitHub. Highest quality in our eval. |
| `claude` | `pip install 'grepxcel[draft-cloud]'` | `ANTHROPIC_API_KEY` | ~$0.003–0.04 per draft. |
| `server` | `pip install openai` | — | Any OpenAI-compatible server (LM Studio, Ollama, vLLM). |
| `gemini` | — | — | Planned, not yet available. |

Keys can be stored in a `.env` file — grepxcel finds it automatically.

For model quality comparisons, see [EVALUATION.md](EVALUATION.md).

### Model cache

The local model is stored once per machine:

| OS | Location |
|---|---|
| Linux | `~/.cache/grepxcel/models/` |
| macOS | `~/Library/Caches/grepxcel/models/` |
| Windows | `%LOCALAPPDATA%\grepxcel\Cache\models\` |

Override with `GREPXCEL_MODEL_DIR`. For faster downloads, set `HF_TOKEN`.

### Corporate proxy / TLS inspection

If your network uses TLS inspection (NetSkope, Zscaler):

```bash
grepxcel draft data.xlsx --ca-bundle /path/to/corporate-ca.pem
```

Or set `GREPXCEL_CA_BUNDLE` / `REQUESTS_CA_BUNDLE` / `SSL_CERT_FILE`.
Run `grepxcel doctor` to verify the setup.

---

## `grepxcel doctor`

Preflight check for your environment.

```bash
grepxcel doctor
grepxcel doctor extract
grepxcel doctor draft
```

| Flag | Purpose |
|---|---|
| `area` | `extract`, `draft`, or `all` (default) |
| `--no-probe` | Skip live TLS handshake |

---

## `grepxcel mcp` / `mcp-config`

Run grepxcel as an MCP server for AI agents.

```bash
pip install 'grepxcel[mcp]'
grepxcel mcp-config               # print config for your agent
grepxcel mcp                      # start the server
```

| Flag (mcp-config) | Default | Purpose |
|---|---|---|
| `--target` | `claude-code` | Format: `claude-code`, `claude-desktop`, `cursor` |

---

## `grepxcel sbom`

Generate a CycloneDX 1.6 Software Bill of Materials.

```bash
grepxcel sbom
grepxcel sbom -o sbom.cdx.json
```

---

## `grepxcel generate-skill`

Write an AI-agent skill document.

```bash
grepxcel generate-skill
grepxcel generate-skill --target agents-md -o AGENTS.md
```

---

## Logging and data safety

grepxcel's structured logs (`--log-format json`) never contain extracted cell
values — by construction, not by redaction. Only an allow-list of safe keys
(`ts`, `level`, `event`, `cell`, `field`, `field_type`, `value_len`,
`value_sha8`) reaches the log file.

This means JSON logs can be shipped to a SIEM or cloud aggregator without risk
of leaking personal data from the Excel files being processed.

The text log (stderr, default) shows full values and is intended for a human
operator at the terminal.
