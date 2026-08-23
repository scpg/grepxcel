# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Features

- **`wizard` command** — `grepxcel wizard data.xlsx` walks the spreadsheet
  cell by cell and asks whether each value is a label, variable, or skip.
  Produces a ready-to-run `.csv` pattern file without requiring knowledge of
  the pattern syntax. Supports `--sheet` and `-o` to control which sheet is
  walked and where the pattern is written.
- **Full-screen TUI wizard** — `pip install 'grepxcel[wizard]'` enables a
  Textual-powered interactive interface with four panel zones (CELL · CLASSIFY ·
  NAVIGATE · LEGEND+STATS), colour-coded classifications (green=label,
  yellow=value, blue=header, magenta=table, dim=ignore), per-column type/match
  modals, undo (Ctrl+Z), clipboard export, session log (F12), SPLITTER row type,
  match-preset cycling, and a column legend step before per-column modals.
  Requires `textual>=8.0` (optional extra — base install unchanged).
- **Improved `_propose_type` heuristic** — the wizard's automatic cell-type
  suggestion now uses structural rules instead of a character-length threshold:
  colon-suffix → label; `@` / `://` → `var:string`; alphanumeric codes (e.g.
  `ELC001`) → `var:string`; multi-word strings → `var:string`; single-word
  pure-alpha → label, overridden to `var:string` when the left neighbour ends
  with `:` (e.g. `"Department:" → "Engineering"`).

### Documentation

- **`docs/wizard-guide.md`** — user-facing reference for the wizard: workflow,
  key bindings, row types, column modals, output format, and known limitations.

## [0.2.0] — 2026-06-25

### Features

- **`extract_df()`** — DataFrame API returning one frame per table key plus a
  `_scalars` frame. Works with pandas or polars (`pip install 'grepxcel[pandas]'`
  / `'grepxcel[polars]'`), auto-detected at runtime. Per-table `header`/`footer`
  fields (e.g. subtotals) are exposed as `{key}__header` / `{key}__footer` frames.
- **`--format csv`** — flat CSV output for single-table patterns. Table data rows
  become CSV rows; scalars are denormalized as repeated columns. Refused (exit 2)
  for multi-table patterns.
- **`--format xlsx`** — colored Excel report for human review: scalar fields, then
  each table top-down with headers, alternating data rows, and footers. Requires
  `-o`; never overwrites a source file.

### Security

- **Formula/CSV injection neutralization** (CWE-1236) — cell values written to
  `--format csv` / `xlsx` that begin with `=`, `+`, `-`, `@`, tab, or carriage
  return are emitted as literal text, so an untrusted source file cannot inject an
  executable formula into the output a human opens.
- **Drafter prompt hardening** — cell values embedded in the `draft` analyser's
  LLM prompt are collapsed to a single line, stripped of control characters, and
  length-bounded, resisting prompt injection from malicious spreadsheet content.

### Changed

- `--format csv` / `xlsx` reject `--all-sheets` (exit 2) — a flat file cannot
  represent multiple sheets; use `--sheet` or `--format nested`.
- Internal: consolidated the nested-dict flattener into `utils.flatten_nested`.

## [0.1.1] — 2026-06-24

### Features

- **`--strict` mode** — `grepxcel extract --strict` exits with code 2 if any
  field has extraction issues (validation mismatch, missing anchor, empty value).
  Lists the affected field names on stderr. Use in pipelines to catch incomplete
  extractions that would otherwise exit 0 or 1 silently.
- **`quickstart` command** — `grepxcel quickstart` prints a guided tutorial in
  your terminal: what grepxcel does, how to create your first pattern, and how
  to run your first extraction. No files created or modified.

### Documentation

- **README rewritten** — shorter, friendlier, focused on non-technical users.
  Full CLI reference moved to `docs/cli-reference.md`.
- **`docs/cli-reference.md`** — comprehensive reference for all CLI commands,
  flags, backends, and logging model (moved from README).
- Removed absolute correctness claims ("never silently wrong") — replaced with
  honest language about the tool's design intent and pattern-dependent quality.
- All README links now absolute GitHub URLs so they resolve on PyPI.

## [0.1.0] — 2026-06-24

Initial public release.

### Features

- **Pattern-based extraction** — define a pattern file (`.xlsx` or `.csv`) once,
  extract structured JSON from any similarly-laid-out `.xlsx` file.
- **Batch directory extraction** — `grepxcel extract -p pat.xlsx data_dir/`
  expands directories to their `.xlsx` files automatically. Add `-r` /
  `--recursive` to recurse into subdirectories. Excel temp files (`~$*.xlsx`)
  are skipped. Empty directories produce a clear error message.
- **`sbom` command** — generates a CycloneDX 1.6 Software Bill of Materials
  listing grepxcel and every transitive dependency with PURLs, SPDX license
  identifiers, and SHA-256 hashes from pip's RECORD files. Uses only stdlib
  (`importlib.metadata`) — no external tool required. Output is accepted by
  Dependency-Track, Grype, and other SBOM consumers. The SBOM is also
  attached as a release artifact on GitHub Releases.
- **`generate-examples` command** — creates 4 ready-to-run example sets
  (pattern + data xlsx + README) in a local directory so new users can try
  grepxcel immediately without needing their own Excel files.
- **MCP server** (`grepxcel mcp`) — expose grepxcel as a Model Context Protocol
  server so AI agents can call extract, validate-pattern, lint, schema, docs,
  doctor, and generate-examples directly. Stdio transport. Install with
  `pip install 'grepxcel[mcp]'`.
- **`mcp-config` command** — prints the MCP server config block for your AI agent
  (Claude Code, Claude Desktop, Cursor). Detects the installed command path.
- **`lbl.match` mode** — `lbl:` fields now default to literal string matching
  instead of full regex. Labels like `Term (months):` or `Breaks\n(minutes)`
  work without regex escaping.
  - Global config: `config: | lbl.match | literal` (default) / `glob` / `regexp`
  - Per-field override: `lbl:literal`, `lbl:glob`, `lbl:regexp` in column A
  - `glob` mode: shell wildcards (`*` = any text including newlines, `?` = one char)
  - `regexp` mode: opt-in for the previous full Python `re.search` behaviour
  - `validate-pattern` warns when a `lbl:` value looks like a regex but mode is
    `literal` or `glob` (e.g. `\(`, `\.`, `(?`)
  - Verbose `validate-pattern -v` now shows `lbl.match` in config and per-field
    `[mode]` tags for explicit overrides
- **Cell addressing**
  - `cell:B5`       -> absolute
  - `cell:next`     -> sequential, next non empty in direction order
  - `seek:`         -> repositioning
  - `dir:`          -> direction switching (LR / TD).
- **Table matching** — repeating mini-tables with:
  - `HEADER`
  - `DATA`          -> mini-table data rows
    - `DATA:1`      -> `1` row
    - `DATA:*`      -> any number of rows
    - `DATA:{n,m}`  -> `n` to `m` rows
  - `FOOTER`
  - `SPLITTER`
  - `SKIP_IF` 
- **`grepxcel.extract()` Python API** — one-call facade with support for:
  - `sheet=`,
  - `all_sheets=True`
  - `output_format=`
- **`draft` command** — generates a starter pattern using one of following options:
  - local GGUF model  -> _offline_, the models will be downloaded.
  - GitHub Models     -> _for those still having it, 
  stand [GitHub changelog 2026-06-16](https://github.blog/changelog/2026-06-16-github-models-is-no-longer-available-to-new-customers/#:~:text=New%20organizations%20and%20enterprises%20without,using%20GitHub%20Models%20for%20now.)_
  - Claude API        -> 
- **`validate-pattern`** — validates pattern files without extracting.
- **`schema`** — generates a JSON Schema (draft 2020-12) from a pattern file.
- **`lint`** — inspects an Excel file for potential issues before extraction.
- **`doctor`** — preflight check for dependencies, API keys, model cache, proxy/TLS.
- **`docs`** — writes a colour-coded `pattern-reference.xlsx`.
- **Structured logging** — `--log-format json` writes NDJSON with an allow-list
  of safe keys only (no extracted cell values, ever). Non-reversible value
  fingerprint (`value_len` + `value_sha8`) aids diagnostics. Summary event with
  extraction statistics (counts + field names).
- **`--meta`** — opt-in `_meta` block in extracted JSON (run_id, stats, safe
  issues) for pipeline auto-verification.
- **`pattern.version`** — optional `config: | pattern.version | N` for
  forward-compatibility; absent defaults to 1.
- **Corporate proxy support** (_**experimental**_) — `truststore`, `--ca-bundle` /
  `GREPXCEL_CA_BUNDLE` for TLS-inspection proxies.
- **Scoped `.env`** — project-local and `~/.config/grepxcel/.env` credential
  discovery; `--strict-env` refuses out-of-project `.env` files.
- **`generate-skill`** — emits a skill doc (`--target claude` or `agents-md`)
  for AI-agent integration.

### Changed

- **structlog integration** — NDJSON file output now uses a structlog processor
  chain (allow-list filter → JSONRenderer) instead of manual dict building +
  `json.dumps`. The output format is unchanged; this is an internal refactor
  that makes the structured-log pipeline extensible for future output targets
  (OTLP, Datadog, etc.). `structlog>=24.1` is now a core dependency.

### Security

- **MCP path sandboxing** — all MCP server tools validate that file paths
  resolve within the server's working directory. Absolute paths pointing
  elsewhere, `..` traversal, and symlink escapes are rejected with a clear
  error. Prevents an AI agent (or a prompt-injected one) from reading or
  writing files outside the intended project directory.
- Fail-closed XXE protection (defusedxml asserted at import).
- ZIP-bomb guard measures real decompressed size (stream, not metadata).
- AST-based ReDoS detection + hard per-match regex timeout.
- SSRF/file-read guards on server URLs (`http(s)` only).
- Scoped `.env` discovery (stops at project root).
- All GitHub Actions pinned to immutable commit SHAs.
- Cached-model sha256 integrity verification on every run.
- Structured JSON logs never contain extracted cell values (allow-list
  construction, not redaction).

[Unreleased]: https://github.com/scpg/grepxcel/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/scpg/grepxcel/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/scpg/grepxcel/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/scpg/grepxcel/releases/tag/v0.1.0
