# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] — Unreleased

Initial public release.

### Added

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
  local GGUF model (default, fully offline) or the Claude API (`--backend claude`),
  with token-cost reporting and a privacy notice for cloud use. (A Gemini backend
  is scaffolded but disabled — planned for a future release.)
- **`docs` command** — writes a colour-coded `pattern-reference.xlsx`.
- **`--version`** flag.
- **Security hardening** — fail-closed XXE protection (defusedxml asserted),
  ZIP-bomb guards (size + expansion ratio), AST-based ReDoS detection on every
  user regex, formula rejection in pattern files, and `.xlsm`/`.xlsb`/`.xls`
  refusal.

[Unreleased]: https://github.com/scpg/grepxcel/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/scpg/grepxcel/releases/tag/v0.1.0
