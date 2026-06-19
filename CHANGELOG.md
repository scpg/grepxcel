# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **`lint` command** — `grepxcel lint data.xlsx` inspects an Excel file before
  extraction, reporting potential issues with a ✓/⚠/✗/ℹ checklist: file format
  and accessibility, encryption/IRM detection (OLE Compound Documents from
  Microsoft Information Protection), sheet dimensions with declared-vs-real
  extent inflation, merged cells, formula cells (stale cached values), empty
  sheets, and multi-sheet inventory. Advisory notes explain known corporate
  environment issues (MIP/IRM sensitivity labels, password protection,
  conditional formatting) even when not yet auto-detected.

- **`--backend server` (OpenAI-compatible server)** — new `draft` backend that talks
  to any server exposing `/v1/chat/completions` (LM Studio, Ollama, vLLM,
  text-generation-inference). Default URL: `http://localhost:1234/v1`. Enables
  GPU/NPU-accelerated local inference on Windows (via LM Studio) while grepxcel
  runs inside WSL2. Model is auto-discovered unless `--server-model` is set.
  `grepxcel doctor` now probes the server and reports loaded models.

- **`dir:` instruction** — new extraction-sequence instruction that switches the
  scalar scan direction partway through a pattern: `dir:LR` (left-to-right) or
  `dir:TD` (top-down). It reads no cell — it only changes the direction for
  subsequent `cell:next` scanning, and resets the cursor so already-read cells are
  skipped. Lets a single pattern read one region top-down and another
  left-to-right without resorting to absolute addressing for everything.

- **`seek:` instruction** — new extraction-sequence instruction that repositions the
  scanner cursor to a target cell (A1-notation) **without reading it**. Enables
  backward repositioning after reading scattered absolute cells, so the next
  `cell:next` starts from the seek position. Example: `seek:I4` followed by
  `cell:I4 | employee.name` is now valid even after reading a cell on a later row.

- **`schema` command** — `grepxcel schema pattern.xlsx` generates a JSON Schema
  (draft 2020-12) from a pattern file, describing the extraction output structure.
  All field types, dot-notation nesting, and table arrays (with header/data/footer
  sub-objects) are mapped to their JSON Schema equivalents. Fields are nullable
  (`[type, "null"]`) since extraction can return `null` for empty cells. Use the
  schema with any standard JSON Schema validator to verify extracted JSON.

- **`duration` field type** — accepts both clock times (`datetime.time`) and
  elapsed durations (`datetime.timedelta`). The existing `time` type also now
  accepts `timedelta` values (permissive), so Excel `[h]:mm` duration cells no
  longer fail when typed as `time`.

- **Coloured pattern files** — generated `.xlsx` pattern files (from `draft`) are
  now colour-coded by row type for readability: `config:` orange, `lbl:` blue,
  `var:` green, `#` comments grey italic, `START:`/`END:` grey, `table:` purple,
  and `HEADER`/`DATA`/`FOOTER`/`SKIP_IF` tinted. Column widths auto-fit. Content
  is never changed. The reusable `scripts/colorize_pattern.py` applies the same
  formatting to any existing pattern file (CSV patterns are unaffected).

### Changed

- **`SKIP_IF` now works with `DATA:*`** — previously restricted to `DATA:{n,m}`
  only. With `DATA:*`, `SKIP_IF` rows are filtered from output while scanning
  continues forward. `DATA:1` remains restricted (ambiguous semantics).

- **Extent guard uses real data extent** — the `--max-rows` / `--max-columns`
  guard now measures the actual used extent (non-empty cells) instead of
  openpyxl's declared `max_row`/`max_column`, which can be inflated by
  styled-but-empty cells. Sheets with formatting beyond the data boundary
  now warn and proceed instead of being rejected.

- **Quieter default output** — per-cell warning blocks (Found / Expected / →)
  now require `-v`; the default shows only the extraction summary and a compact
  one-line-per-issue recap. No information is lost — `-v` restores the full
  detail alongside the per-field trace.

- **`validate-pattern -v` improvements** — config settings print one per line
  instead of all on one line; table extraction sequences display as a columnar
  grid (one column-position per row, all row types side by side) instead of
  bracket-delimited lists.

### Security

- **ZIP-bomb guard now measures real decompressed size** — `_check_zip_safety`
  previously summed the central-directory `file_size` field, which is
  attacker-controlled metadata (a crafted `.xlsx` could declare 0-byte entries
  while DEFLATE-expanding to gigabytes). It now stream-decompresses every member
  in bounded chunks and counts the bytes actually emitted, aborting the moment the
  running total crosses the `--max-uncompressed` cap — so a bomb is rejected after
  reading at most one chunk past the limit.
- **SSRF / file-read guards on server URLs** — `grepxcel doctor`'s server probe
  and the `--server-url` (OpenAI-compatible) backend now refuse any non-`http(s)`
  URL via a shared `is_http_url` check, so `file://…`, `ftp://…`, or a cloud
  metadata IP can't be reached through a `GREPXCEL_SERVER_URL` / `--server-url`.
- **Scoped `.env` discovery** — `.env` lookup stops at the project root (`.git` /
  `pyproject.toml`) instead of walking to the filesystem root, so an unrelated
  ancestor project's `.env` can no longer silently inject its API keys; a `.env`
  loaded from an ancestor directory is now announced on stderr.
- **`--log` appends instead of truncating** — the log file is opened in append
  mode (matching its documented behaviour); pointing `--log` at an existing file
  no longer silently destroys its contents.
- **Clarified ZIP expansion-ratio constant** — `MAX_EXPANSION_RATIO = 50` now
  names the actual enforced ceiling (was a confusing `DEFAULT_MAX_EXPANSION_RATIO
  * 10`); the old name remains as an alias.
- **Doc note** — recommend `GREPXCEL_VERIFY_MODEL=full` on shared/server
  deployments where the model cache is writable by others.

### Fixed

- **Case-insensitive instruction keywords** — `CELL:`, `Cell:`, `cell:` (and the
  same for `SEEK:`/`DIR:`/`TABLE:`/`CONFIG:`/`VAR:`/`LBL:`/`HEADER:`/`DATA:`/…) are
  now all accepted. Previously only the lowercase form matched, so an uppercased
  keyword was silently dropped, leaving its fields defined-but-unused. Field names
  and cell addresses keep their original case.

- **No more silently-ignored pattern rows** — a non-empty row inside `START:` (or
  before it) that matches no known keyword is now a clear parse error instead of
  being skipped. Catches typos like `cel:J5` or `tabel:*` that used to produce
  wrong/empty output with no warning. `doc:`/`info:` comments and blank rows are
  still allowed.

- **Coloured pattern files: `dir:` rows** — the pattern colorizer now tints `dir:`
  instruction rows (and recognises `def:`/`info:` aliases); previously they were
  left uncoloured.

- **`seek:` ordering reset** — `seek:X` followed immediately by `cell:X` no longer
  raises a "before or equal to previous reference" parse error. The absolute-ref
  ordering constraint is now fully reset (to `None`) after every `seek:`, so the
  first abs ref after a seek is unchecked at parse time. Backward refs after seek
  are still caught at runtime by the engine's "already passed" check.

## [0.1.0] — Unreleased

Initial public release.

### Added

- **`grepxcel.extract()` Python API** — a one-call facade
  (`grepxcel.extract("pattern.xlsx", "data.xlsx") -> dict`) that wraps `Engine`,
  is silent by default (library-friendly), and supports `sheet=`,
  `all_sheets=True`, and `output_format=`. `Engine` remains available for finer
  control.
- **`extract` command** — pattern-based data extraction from `.xlsx` files into
  nested or legacy JSON. Pattern files describe layout with `config:` / `lbl:` /
  `var:` / `doc:` rows and a `START:`…`END:` extraction sequence.
- **Cell addressing** — `cell:next` (sequential) and `cell:B5` (absolute
  A1-notation), plus `IGNORE` / `EMPTY` column keywords.
- **Table matching** — repeating mini-table blocks with `HEADER` / `DATA` /
  `FOOTER` / `SPLITTER` rows; `DATA:1`, `DATA:*`, and bounded `DATA:{n,m}` with
  `SKIP_IF` for fixed-slot templates.
- **Pattern file formats** — patterns can be authored as `.xlsx` **or** `.csv`;
  both are read into one common grid representation and behave identically.
- **`--sheet` / `--all-sheets`** sheet selection.
- **`draft` command** — drafts a starter pattern for an unseen file using a
  local GGUF model (**Gemma-4-E4B** by default, fully offline), **GitHub Models**
  (`--backend github`, free with a GitHub subscription — the highest-quality
  option in our eval), or the Claude API (`--backend claude`), with token-usage /
  cost reporting and a privacy notice for cloud use. Cloud backends read keys
  from a `.env` file automatically. (A Gemini backend is scaffolded but disabled
  — planned for a future release.)
- **Portable model-trust registry** — `KNOWN_MODEL_HASHES` pins the sha256 of
  the bundled model in code, so a cached or manually-placed file matching the
  known-good build is trusted on any machine (and a stale local trust record is
  healed automatically), without re-downloading.
- **Platform-appropriate model cache** — the local model is stored in the
  per-user cache resolved via `platformdirs` (`~/.cache/grepxcel/models` on Linux,
  `~/Library/Caches/grepxcel/models` on macOS, `%LOCALAPPDATA%\grepxcel\Cache\models`
  on Windows; honors `$XDG_CACHE_HOME`). `GREPXCEL_MODEL_DIR` overrides it.
- **`config: | ignore.case | yes`** — opt-in case-insensitive regex matching for
  every `lbl:`/`var:` pattern (global, inherited by tables). Default stays
  case-sensitive.
- **Consolidated issue recap** — the extraction summary now ends with an
  `ISSUES (cell — reason)` block listing every failing cell and why, so problems
  aren't lost in the scrollback of a long run.
- **`docs` command** — writes a colour-coded `pattern-reference.xlsx`.
- **`validate-pattern` command** — `grepxcel validate-pattern FILE...` checks a
  pattern file (`.xlsx` or `.csv`) is valid to use *without* extracting: it runs
  the full parser (structure, types, multiplicities, regex safety, comments) and
  additionally flags an empty extraction sequence and references to undefined
  fields. `-v` prints the parsed config, fields, and steps; non-zero exit if any
  file is invalid.
- **`doctor` command** — `grepxcel doctor [extract|draft|all]` preflight check:
  Python version, core/optional deps, API keys, the local-model cache + disk
  space, and corporate-proxy / TLS trust with a live handshake probe. Prints a
  `✓/⚠/✗` checklist and exits non-zero on a blocking problem.
- **Corporate TLS-inspection proxy support** (experimental) — grepxcel can run
  behind an intercept proxy (e.g. NetSkope/Zscaler) that re-signs HTTPS with a
  company CA. Optional `truststore` uses the OS trust store automatically; or set
  `--ca-bundle` / `GREPXCEL_CA_BUNDLE` (also honors `REQUESTS_CA_BUNDLE` /
  `SSL_CERT_FILE`). Covers the local-model download and the cloud `draft`
  backends. TLS verification is never disabled. *Not yet tested against a real
  intercept proxy — `enable_corporate_tls` prints a one-time caveat.*
- **`--version`** flag.
- **Pattern-file sanity validation** — every pattern file is now fully
  validated up front; malformed patterns fail with a clear message instead of
  silently producing wrong/empty output or crashing. New checks: unknown field
  types, unnamed `var:`/`lbl:` fields, patterns with no `START:` section, empty
  `START: … END:` blocks, `table:` blocks with no `DATA` row, `HEADER` after
  `DATA` / `FOOTER` before `DATA`, unknown table row keywords, and invalid
  `table:`/`DATA:`/`HEADER:`/`FOOTER:` multiplicities.
- **More field types** — added `number`/`float`/`decimal` (plain numerics),
  `text` (alias for `string`), and `boolean` (TRUE/FALSE) alongside the existing
  `string`, `integer`, `currency`, `percentage`, `date`/`datetime`/`timestamp`.
- **Verbose extraction trace** — `-v` now prints a per-field trace
  (`field ← B1 = value ✓`/`✗`) for both scalar cells and table DATA fields, with
  the failing regex shown on a `✗`, so it's easy to see what was extracted from
  where and why a pattern didn't match.
- **Data-sheet size limits** — a guard against oversized sheets with sensible
  defaults (`--max-rows` 2048, `--max-columns` 1024). Larger sheets fail with a
  message explaining how to raise the limit (up to Excel's maximum) and that the
  tool is untested at that scale.
- **Security hardening** — fail-closed XXE protection (defusedxml asserted),
  ZIP-bomb guards (size + expansion ratio), AST-based ReDoS detection plus a hard
  per-match timeout on the `regex` engine (default 0.25 s,
  `GREPXCEL_REGEX_TIMEOUT`) so even backtracking patterns the static guard can't
  see are bounded, formula rejection in pattern files, and `.xlsm`/`.xlsb`/`.xls`
  refusal.
- **Supply-chain hardening** — all GitHub Actions pinned to immutable commit
  SHAs (Dependabot keeps them current), least-privilege `permissions:` on CI, and
  cached-model integrity verification: the downloaded model's sha256 is recorded
  and re-checked on every run (re-hashing only when size/mtime changed by default,
  or always with `GREPXCEL_VERIFY_MODEL=full`); a mismatch aborts unless
  `--allow-unverified-model` / `GREPXCEL_ALLOW_UNVERIFIED_MODEL`.

[Unreleased]: https://github.com/scpg/grepxcel/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/scpg/grepxcel/releases/tag/v0.1.0
