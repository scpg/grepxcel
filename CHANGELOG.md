# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.5.0] — 2026-09-27

A minor rather than a patch bump: temporal coercion **changes extraction output** for files
whose date and time values are stored as text. A field that previously held the raw string
plus a type-mismatch warning now holds a real temporal value. Patterns do not change, but
anything consuming the JSON downstream may.

### Features

- **Date and time text is converted, not refused** — a cell declared `date`, `datetime`,
  `timestamp`, `time` or `duration` in the pattern is now converted to a real temporal value
  when it holds unambiguous text. Dates often arrive as text (exported from another system,
  typed with a leading apostrophe, or written without a date number format) and were
  previously reported as type mismatches even though the pattern had already declared what
  they were. `'2024-01-15'` → a datetime; `'09:30'` → a time; `'2:00'` and `'30:00'` → 2 and
  30 hours elapsed. **This changes extraction output** for such files: the field now holds a
  timestamp where it previously held the raw string plus a warning.
- **`config: | date.format |` and `config: | time.format |`** — declare a strptime format
  (e.g. `%d/%m/%Y`) for text that ISO 8601 cannot resolve. Without one, ambiguous text is
  refused rather than guessed: `01/02/2024` is 1 February in most of the world and 2 January
  in the United States, and picking one silently is how a tool returns confidently wrong data.
  A declared format is tried first and ISO still converts afterwards, so a sheet mixing
  `31/12/2024` and `2024-12-31` reads correctly either way. An unusable format string is
  rejected when the pattern is parsed, so the error names the pattern rather than arriving as
  a wall of per-cell mismatches pointing at the data.
  No new dependency: a lenient parser was evaluated and declined, because it reads `'09:30'`
  as *today's* date at 09:30 (making the same file extract differently tomorrow), returns a
  datetime rather than a duration for `'2:00'`, and cannot parse `'30:00'` at all.

### Fixed

- **A rejected value now says why it was rejected.** `validate_type` has always computed a
  precise reason — `is not a whole number`, `boolean is not integer`,
  `javascript: URLs are not permitted`, and the two that matter most: a cell over
  `--max-cell-len` and a regex timeout, where the pattern was never evaluated at all. The
  engine discarded it, so every one of those causes reached the user as the same sentence:
  *"Value does not match the expected pattern … Expected: matches /…/"*. A cell that could
  not possibly have matched was reported identically to one that simply did not. The reason
  now appears in the warning message and as a `Reason:` line in the rendered output.
- **`schema` and `extract` disagreed about `boolean`.** `validate_type` accepts `1`/`0`,
  `"1"`/`"0"` and `"yes"`/`"no"`, but the generated JSON Schema declared only
  `{"type": ["boolean", "null"]}`, so a schema generated from a pattern rejected 3 of the 4
  boolean forms that same pattern's extraction accepts. The declared types are widened to
  match. *Known gap:* this makes the boolean schema accept values extraction rejects (`-1`,
  `'maybe'`); tightening it is tracked in
  `docs/superpowers/specs/2026-09-27-schema-strictness-decisions.md`.
- **A bare number is never read as a time.** Python 3.11+ widened `time.fromisoformat` to
  accept bare-hour and compact forms, which would have read `'12'` as 12:00, `'1230'` as
  12:30 and `'2024'` as **20:24** — so a year, an ID or a quantity in a time-typed column
  would have become a plausible-looking time. A separator is now required; compact times
  remain readable via `time.format`.
- **`grepxcel docs` output is reproducible.** The generated `pattern-reference.xlsx` embedded
  a wall-clock timestamp, the absolute path of the invoking interpreter, and a ZIP host byte
  that differed between Linux and Windows — so two runs of the same version produced
  byte-different files and the document shipped a fragment of the machine that built it. The
  timestamp now derives from a single fixed source, the path is gone, and the archive
  metadata is pinned.

### Security

- **The `[mcp]` extra now requires `mcp>=1.28.1`** (was `>=1.23`). The older range permitted
  two HIGH-severity advisories: CVE-2026-52869 (CVSS 7.1 — the SSE and Streamable HTTP
  transports routed requests to a session by id without checking that the caller was the
  principal who created it; fixed in 1.27.2) and CVE-2026-59950 (CVSS 8.1 — the deprecated
  websocket transport accepted handshakes without Host or Origin validation; fixed in 1.28.1).
  grepxcel's server runs on stdio, so neither transport was reachable through it; the floor is
  raised so installing the extra cannot place a version with known holes into an environment
  that exposes them another way. The ceiling stays below 2: mcp 2.x removed
  `mcp.server.fastmcp`, which this server imports.
- **The `[mcp]` install check actually exercises the SDK.** `grepxcel mcp-config` only prints
  a JSON blob, and `mcp_server.py` deliberately swallows a failed SDK import
  (`except ImportError: FastMCP = None`) so the CLI can show an install hint — between them, an
  incompatible `mcp` release produced a green check and a broken `grepxcel mcp`. CI now asserts
  the import bound and builds a server instance.
- **The wizard's CDN scripts are pinned with Subresource Integrity.** The web wizard loads
  three libraries from cdnjs; there is no `package.json`, so neither Dependabot nor Snyk sees
  those versions. Each `<script>` now carries a `sha512` `integrity` hash and
  `crossorigin="anonymous"`, so a modified file is refused rather than executed against
  whatever spreadsheet the user pointed the wizard at. A test asserts every external script
  is hashed, strongly hashed and version-pinned — SRI's own failure mode is bumping a version
  without recomputing the hash, which leaves the page loading and the library silently absent.
- **Static analysis is enforced rather than advisory.** `bandit` runs in CI and fails the
  build at MEDIUM severity and above (configured in `pyproject.toml`; tests are excluded,
  since a security-conscious suite builds hostile input on purpose). The source already
  carried seven justified `# nosec` annotations from a hand-run that was never wired up — by
  the time it was enforced, a second un-triaged `urlopen` had appeared. CodeQL
  (`security-extended`) runs alongside it on pull requests and weekly, and adds the view
  bandit cannot give: whether a dangerous call is *reachable from untrusted input*. Secret
  scanning and push protection are enabled on the repository.

### CI / Infrastructure

- Checks now run on **every** pull request. Both workflows filtered `pull_request` to
  `[main, dev]`, so a pull request targeting a feature branch ran nothing at all — and the
  gap was invisible, because no checks appeared to fail rather than appearing red. All
  third-party actions are pinned to a commit SHA.
- The oracle job runs under `pytest-xdist` (`-n auto`). The fast suite stays serial by
  design: `proxy_support` installs a process-global SSL truststore, which a shared worker
  would carry into tests that mock it.

### Tests

- **Oracle type-matrix suite** (`tests/oracle/`, ~26,900 generated cases) — for every
  combination of value type × file shape × OOXML storage type × number format × read
  direction, the suite generates a manifest of known ground truth and forces
  `validate-pattern`, `lint`, `profile`, `extract` and `schema` to agree with it. Excluded
  from `pytest tests/` by the `oracle` marker; run as a pre-PR gate via
  `scripts/oracle_check.py`, and on pull requests by `.github/workflows/oracle.yml`. It found
  both `Fixed` entries above on its first full run.
- **Excel storage-model boundaries pinned** — the 1900 leap-year bug (serials 59 and 60 both
  read as 1900-02-28, and serials 1–59 sit one day ahead of naive epoch arithmetic); the
  two-day hole where 1899-12-30 and 1899-12-31 come back as `00:00:00` instead of a date; the
  2⁵³ integer precision cliff, above which odd integers snap to an even neighbour and a
  19-digit identifier comes back as a different number while still validating as an integer.

### Documentation

- `docs/pattern-file.md`: `date.format` / `time.format`, why ambiguous text is refused, a
  note on numbers in duration cells (Excel's unit is one day, so a bare `12` is twelve days),
  and **"Long numeric identifiers: declare them `string`"** — above 2⁵³ a numeric cell cannot
  hold a long ID exactly, and the rounding is undetectable after the fact.

## [0.4.1] — 2026-09-21

### Fixed

- **`not-null` / `not-empty` round-trip fidelity** — `FieldDef` now stores `required_token`
  (the original parser token) so the web wizard no longer silently rewrites `not-empty` fields
  as `not-null` when a pattern is saved after preload.
- **Single source of truth for token sets** — `VALID_MODE_TOKENS` / `VALID_CONSTRAINT_TOKENS`
  exported from `pattern_parser.py` and imported by `wizard_core.py`; local duplicate
  frozensets in `_col_a_extra_to_parts` eliminated.
- **`var:(default)` guard** — the pattern writer now strips `(default)` from `var:` col-A
  cells, symmetric with the existing `lbl:(default)` guard.
- **Stale `lbl:(default)` fixture** — `01_simple_invoice_data_pattern-from-web.csv` had
  8 rows written with the old `lbl:(default)` syntax that `PatternParser` rejects; fixed to
  `lbl:`.
- **Silent test skip** — `test_pattern_tester.py` referenced a non-existent fixture filename,
  causing 10 tests to silently skip; corrected.

### Improved

- **AI-debug logging in web-wizard session log** — three new structured event types make
  session logs machine-readable for bug analysis and architecture tracing:
  - `PRELOAD_FIELD` — one entry per cell loaded from a pattern file (role, name, type,
    match pattern, col_a_extra modifiers).
  - `SCHEMA_SUMMARY` — emitted immediately before each SAVE; reports `lbl=N var=N total=N`
    and a modifier distribution (`not-null×3 trim-whitespace×1 …`).
  - Session log header updated to document all event types.
- **Unrecognized-token warning** — `_col_a_extra_to_parts` now emits a `logging.WARNING`
  when it encounters tokens not in the parser's valid sets, naming the offending tokens and
  the complete valid token lists.

## [0.4.0] — 2026-09-20

### Features

- **`assert:` rules** — declare cross-field validation expressions directly in the pattern
  file. Evaluated after each extraction; a failed assertion logs a `WARNING` alongside the
  field-level issues. Syntax: column A = `assert:`, column B = expression, column C =
  optional message. Expressions support arithmetic, comparisons, boolean operators, and
  numeric coercion from Excel string values. Uses a safe AST evaluator — no `eval`/`exec`.
  Example: `assert:  | total == net + vat | totals must balance`.
- **`grepxcel watch`** (`pip install 'grepxcel[watch]'`) — persistent directory monitor that
  runs extraction automatically as new `.xlsx` files arrive. Uses platform-native file-system
  events (inotify / FSEvents / ReadDirectoryChangesW). Supports `--recursive`, `-o output/`,
  and `--on-error stop`. Results are newline-delimited JSON to stdout or written to the output
  directory. Skips Excel temp files (`~$*.xlsx`). Ctrl+C to stop.
- **`grepxcel test`** — pattern reliability test suite. Run `grepxcel test -p pattern.xlsx
  samples/` against a directory of representative files to get a pass/warn/fail breakdown per
  field. Field reliability scores (`47/47 ████`) help identify brittle anchors before
  going to production.
- **`flat_tables` / `--no-source`** — `extract(…, flat_tables=True)` and
  `grepxcel extract --no-source` collapse the `[{"_source": …, "data": […]}]` per-instance
  wrappers from nested output into flat `[row, …]` lists, making the result easier to pass
  directly to pandas or similar pipelines.
- **Visual audit sheet** — `--format xlsx` output now includes a `_Audit` sheet that
  colour-codes every extracted cell in the source data: green for matched fields, amber for
  warnings, red for errors. Makes QA of large reports practical without scripting.
- **Web wizard: range selection & mini-table editor** — Shift+Click and Shift+Arrow now
  batch-classify a rectangular range in one operation. The mini-table editor lets you assign
  column roles (Label / Variable / Ignore), types, and match expressions for an entire table
  in a single modal, with preloaded footer row detection.
- **Web wizard: in-grid badge system** — every cell now shows a compact two-strip badge
  indicating its classification (L/V/I/C) and, for table cells, its positional role
  (T-HEAD header / footer, T-DATA row) and anchor marker (▶/◀). Badges reserve fixed
  width so the cell values stay aligned across rows.
- **Web wizard: cascade CLEAR** — clearing a cell that anchors a mini-table propagates
  the clear to all dependent column-role cells automatically, preventing orphaned state.
- **Web wizard: SKIP_EMPTY_ROW preload** — `SKIP_EMPTY_ROW:N` rows in an existing
  pattern are now correctly restored when the wizard preloads a pattern file.
- **Web wizard: Header(C) removed** — the combined `Header(C)` classification (header +
  CONFIG) has been removed; classify headers and CONFIG rows separately.

### Changed

- **`--format legacy` deprecated** — using `--format legacy` now emits a `DeprecationWarning`
  (Python API) and a visible `⚠ DEPRECATED` message on stderr (CLI). The legacy format leaks
  internal `lbl:` keys and `_source`/`_anchor` metadata; migrate to `--format nested`
  (the default).
- **TUI wizard removed** — the Textual-based TUI (`grepxcel wizard`) has been removed; use
  `grepxcel web-wizard` instead.

### Documentation

- **Merged cells guide** — new `## Merged cells` section in `docs/pattern-file.md` covers
  top-left cell reading, labels and values spanning merged ranges, tables with merged headers,
  and `grepxcel lint` merged-cell reporting.
- **Web wizard guide** — new `docs/web-wizard-guide.md` with install, interface overview,
  cell classification keys, keyboard shortcuts, saving, regex presets, merged cells, and
  troubleshooting.
- **CLI reference** — `--no-source` flag added to extract flags table; `legacy` output format
  marked ⚠️ Deprecated.
- **README** — pre-built binaries section (download table, macOS/Windows first-run notes);
  `flat_tables=True` Python API example; `--no-source` in Output formats; merged cells link.

### CI / Infrastructure

- **Pre-built binaries** — PyInstaller-based release pipeline builds single-file executables
  for Linux, macOS, and Windows on every version tag push and attaches them (with SHA-256
  checksums) to the GitHub Release. Requires `pip install 'grepxcel[build]'`. Uses
  PyInstaller ≥ 6.10 (fixes CVE-2025-59042).

## [0.3.1] — 2026-09-15

### Changed

- **PyPI listing overhaul** — improved short description, added PyPI/Python/license/downloads
  badges, added "Capabilities at a glance" feature bullet list for faster scanning by humans
  and AI agents, added "What's new in v0.3.0" section with CHANGELOG link
- **Reduced dead whitespace** — removed 10 of 11 `---` horizontal-rule dividers from
  README (each rendered as ~70 px of blank space on PyPI)
- **Metadata** — upgraded `Development Status` classifier from `4 - Beta` to
  `5 - Production/Stable`; added Python 3.15 classifier; added `Environment :: Console`,
  `Intended Audience :: Science/Research`, `Topic :: Scientific/Engineering :: Information
  Analysis`, `Topic :: Utilities`, `Typing :: Typed` classifiers; expanded keywords with
  `automation`, `invoice`, `structured-data`, `data-pipeline`, `mcp`, `ai-agent`, and others

## [0.3.0] — 2026-09-15

### Features

- **Column A modifiers for `lbl:` and `var:`** — order-independent colon-separated
  tokens in column A extend both anchor and variable rows with optional quality gates:
  - **`not-null` / `not-empty`** (synonyms) — fatal error if the extracted value is
    empty or null, always, regardless of `--strict`. Example: `var:not-null`.
  - **`var:glob`** — column D is matched as a shell-style glob (`*`, `?`) against the
    string representation of the value. Type check still runs first. Example:
    `var:glob | sku | string | PROD-*`.
  - **`var:literal`** — column D is matched as an exact string. Special regex characters
    are treated as literals. Example: `var:literal | status | string | Active`.
  - **`var:re`** — explicit alias for the default regex mode.
  - **`lbl:re`** — short alias for `lbl:regexp`.
  - Modifiers are order-independent and composable: `var:not-null:glob`,
    `lbl:not-null`, `var:literal:not-empty`, etc.
  - Backward-compatible: bare `var:` with a non-empty column D continues to use regex.
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

[Unreleased]: https://github.com/scpg/grepxcel/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/scpg/grepxcel/compare/v0.4.1...v0.5.0
[0.4.1]: https://github.com/scpg/grepxcel/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/scpg/grepxcel/compare/v0.3.1...v0.4.0
[0.3.1]: https://github.com/scpg/grepxcel/compare/v0.3.0...v0.3.1
[0.3.0]: https://github.com/scpg/grepxcel/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/scpg/grepxcel/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/scpg/grepxcel/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/scpg/grepxcel/releases/tag/v0.1.0
