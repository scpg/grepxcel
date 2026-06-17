# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **`--backend server` (OpenAI-compatible server)** — new `draft` backend that talks
  to any server exposing `/v1/chat/completions` (LM Studio, Ollama, vLLM,
  text-generation-inference). Default URL: `http://localhost:1234/v1`. Enables
  GPU/NPU-accelerated local inference on Windows (via LM Studio) while grepxcel
  runs inside WSL2. Model is auto-discovered unless `--server-model` is set.
  `grepxcel doctor` now probes the server and reports loaded models.

- **`seek:` instruction** — new extraction-sequence instruction that repositions the
  scanner cursor to a target cell (A1-notation) **without reading it**. Enables
  backward repositioning after reading scattered absolute cells, so the next
  `cell:next` starts from the seek position. Example: `seek:I4` followed by
  `cell:I4 | employee.name` is now valid even after reading a cell on a later row.

### Changed

- **`SKIP_IF` now works with `DATA:*`** — previously restricted to `DATA:{n,m}`
  only. With `DATA:*`, `SKIP_IF` rows are filtered from output while scanning
  continues forward. `DATA:1` remains restricted (ambiguous semantics).

### Fixed

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
