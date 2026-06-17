# Design: OpenAI-compatible server backend for `grepxcel draft`

**Date:** 2026-06-17
**Status:** Proposed

## Problem

The `grepxcel draft` command supports local (llama-cpp-python), Claude, and
GitHub Models backends. On Windows machines with GPU/NPU hardware, the local
backend runs inside WSL2 where GPU acceleration is limited (NVIDIA CUDA only)
and NPU access is unavailable. Applications like LM Studio, Ollama, and vLLM
run natively on Windows with full hardware access and expose an
OpenAI-compatible HTTP API that WSL2 can reach via localhost.

## Solution

Add an `OpenAICompatBackend` class that talks to any server exposing the
standard `/v1/chat/completions` endpoint. Wire it as `--backend server` in the
CLI.

## Design

### Backend class: `OpenAICompatBackend` (in `grepxcel/drafter.py`)

Follows the existing backend pattern (`LLMBackend` protocol + optional
`last_cost()`).

- **Constructor args:** `base_url` (default `http://localhost:1234/v1`),
  `model` (default `None` = auto-discover), `api_key` (default `'not-needed'`).
- **`chat(system, user) -> str`:** uses the `openai` SDK (`OpenAI(base_url=...,
  api_key=...)`). Temperature 0.1, max_tokens 2048 (same as all other
  backends).
- **Model auto-discovery:** when no model is specified, queries `GET /v1/models`
  and picks the first loaded model. Raises `RuntimeError` if no models are
  loaded.
- **`last_cost() -> CostRecord`:** reports token counts with `$0.00` cost (local
  server, no metering). No rate-limit headers.
- **No proxy support:** this backend targets localhost/LAN servers; corporate
  TLS-inspection is irrelevant.
- **Dependency:** `openai` SDK (already required by GitHub Models backend, in
  the `draft-cloud` extra).

### CLI changes (`grepxcel/cli.py`)

- `--backend` choices: add `'server'`.
- New flags (only relevant when `--backend server`):
  - `--server-url URL` (default from `GREPXCEL_SERVER_URL` or
    `http://localhost:1234/v1`)
  - `--server-model MODEL` (default from `GREPXCEL_SERVER_MODEL` or `None`)
  - `--server-api-key KEY` (default from `GREPXCEL_SERVER_API_KEY` or
    `'not-needed'`)
- In `_run_draft()`: new `elif selected == 'server':` block. No privacy notice
  (data stays local unless the user explicitly points at a remote URL).
- Epilog help text: document the `server` backend with LM Studio / Ollama /
  vLLM as examples.

### Doctor check (`grepxcel/doctor.py`)

New `check_server()` function:
- Reads `GREPXCEL_SERVER_URL` (default `http://localhost:1234/v1`).
- `GET {base_url}/models` with 3-second timeout via `urllib.request`.
- Reports loaded model name on success (OK), warns on connection refused / no
  models loaded (WARN), or fails on unexpected error (FAIL).
- Wired into `run_doctor(area='draft')` as an optional check (only runs when
  `GREPXCEL_SERVER_URL` is set or the default port is reachable).

### Environment variables

| Variable | Purpose | Default |
|---|---|---|
| `GREPXCEL_SERVER_URL` | Base URL for the OpenAI-compatible server | `http://localhost:1234/v1` |
| `GREPXCEL_SERVER_MODEL` | Model id to use | `None` (auto-discover) |
| `GREPXCEL_SERVER_API_KEY` | API key (if server requires auth) | `not-needed` |

CLI flags take precedence over env vars.

### Testing

- **`tests/unit/test_drafter.py`**: `TestOpenAICompatBackend` — mock the
  `openai.OpenAI` client; verify `chat()` returns content, `_resolve_model()`
  picks the first model, error on empty model list.
- **`tests/unit/test_cli_args.py`** (or existing CLI test file): verify
  `--backend server --server-url X` parses correctly.
- **`tests/unit/test_doctor.py`**: mock HTTP probe; verify OK/WARN/FAIL output.
- No live integration test (would need a running server).

### Documentation

- `README.md`: add `server` to the backend list.
- CLI epilog: document the backend with examples.
- `CHANGELOG.md`: entry under `[Unreleased]`.

## Files to modify

| File | Change |
|---|---|
| `grepxcel/drafter.py` | `OpenAICompatBackend` class (~45 lines) |
| `grepxcel/cli.py` | `--backend server`, 3 flags, `elif`, epilog |
| `grepxcel/doctor.py` | `check_server()` + wire into `run_doctor` |
| `tests/unit/test_drafter.py` | ~3 unit tests |
| `tests/unit/test_cli_args.py` | CLI parsing test |
| `tests/unit/test_doctor.py` | Doctor server check test |
| `README.md` | Backend list |
| `CHANGELOG.md` | Unreleased entry |

## Out of scope

- No new pip dependency (reuses `openai` from `draft-cloud` extra).
- No backend registry / plugin system (if/elif chain is fine at 5 entries).
- No streaming support (matches existing backends).
- No Gemini enablement (separate effort, blocked on GCP billing).
