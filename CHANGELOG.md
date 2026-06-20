# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] — YYYY-MM-DD

Initial public release.

### Features

- **Pattern-based extraction** — define a pattern file (`.xlsx` or `.csv`) once,
  extract structured JSON from any similarly-laid-out `.xlsx` file.
- **Cell addressing** — `cell:next` (sequential) and `cell:B5` (absolute), plus
  `seek:` repositioning and `dir:` direction switching (LR / TD).
- **Table matching** — repeating mini-tables with `HEADER` / `DATA` / `FOOTER` /
  `SPLITTER` / `SKIP_IF` rows; `DATA:1`, `DATA:*`, and bounded `DATA:{n,m}`.
- **`grepxcel.extract()` Python API** — one-call facade with `sheet=`,
  `all_sheets=True`, and `output_format=` support.
- **`draft` command** — generates a starter pattern using a local GGUF model
  (offline), GitHub Models, or the Claude API, with token/cost reporting.
- **`validate-pattern`** — validates pattern files without extracting.
- **`schema`** — generates a JSON Schema (draft 2020-12) from a pattern file.
- **`lint`** — inspects an Excel file for potential issues before extraction.
- **`doctor`** — preflight check for dependencies, API keys, model cache,
  proxy/TLS.
- **`docs`** — writes a colour-coded `pattern-reference.xlsx`.
- **Structured logging** — `--log-format json` writes NDJSON with an allow-list
  of safe keys only (no extracted cell values, ever). Non-reversible value
  fingerprint (`value_len` + `value_sha8`) aids diagnostics. Summary event with
  extraction statistics (counts + field names).
- **`--meta`** — opt-in `_meta` block in extracted JSON (run_id, stats, safe
  issues) for pipeline auto-verification.
- **`pattern.version`** — optional `config: | pattern.version | N` for
  forward-compatibility; absent defaults to 1.
- **Corporate proxy support** (experimental) — `truststore`, `--ca-bundle` /
  `GREPXCEL_CA_BUNDLE` for TLS-inspection proxies.
- **Scoped `.env`** — project-local and `~/.config/grepxcel/.env` credential
  discovery; `--strict-env` refuses out-of-project `.env` files.
- **`generate-skill`** — emits a skill doc (`--target claude` or `agents-md`)
  for AI-agent integration.

### Security

- Fail-closed XXE protection (defusedxml asserted at import).
- ZIP-bomb guard measures real decompressed size (stream, not metadata).
- AST-based ReDoS detection + hard per-match regex timeout.
- SSRF/file-read guards on server URLs (`http(s)` only).
- Scoped `.env` discovery (stops at project root).
- All GitHub Actions pinned to immutable commit SHAs.
- Cached-model sha256 integrity verification on every run.
- Structured JSON logs never contain extracted cell values (allow-list
  construction, not redaction).

[Unreleased]: https://github.com/scpg/grepxcel/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/scpg/grepxcel/releases/tag/v0.1.0
