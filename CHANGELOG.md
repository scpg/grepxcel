# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
- **`--version`** flag.
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
