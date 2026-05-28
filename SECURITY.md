# Security Policy

## Supported versions

Only the latest release on `main` receives security fixes.

## Reporting a vulnerability

**Please do not open a public GitHub issue for security vulnerabilities.**

Use GitHub's built-in private reporting:

1. Go to the **Security** tab of this repository
2. Click **Report a vulnerability**
3. Fill in the description, steps to reproduce, and impact

You can expect an acknowledgement within 72 hours and a fix or mitigation plan within 14 days for confirmed vulnerabilities.

## Scope

In scope:

- `grepxcel` CLI and Python API (`engine/` package)
- Malicious Excel file handling (formula injection, ZIP bombs, oversized files, XXE)
- Any path traversal or code execution issue triggered by a crafted `.xlsx` file

Out of scope:

- Issues in `tmp.local/` (development-only, gitignored, never installed)
- Vulnerabilities in the local LLM used by `grepxcel suggest` (third-party model)
- Social engineering or phishing

## Security design notes

grepxcel applies several defences against malicious Excel files:

- `defusedxml` for XML parsing (XXE protection). This is asserted fail-closed
  before every file load — the tool refuses to parse input if defusedxml is not
  active, rather than parsing it unsafely.
- Compressed and uncompressed file size limits, plus an expansion-ratio ceiling
  (ZIP bomb guard).
- Per-cell character length limit and AST-based nested-quantifier rejection
  (ReDoS guard) on every user-supplied regex.
- Formula detection: pattern files must contain plain text only.
- `data_only=True` when loading data files (formulas are never executed).
- `.xlsm` / `.xlsb` / `.xls` are rejected; only macro-free `.xlsx` is accepted.

### `grepxcel suggest` (optional extra) supply chain

The optional `suggest` command downloads a GGUF model from Hugging Face:

- The model is **pinned to an immutable commit revision**; `huggingface_hub`
  verifies the file hash against the Hub for that revision.
- **Auto-update is off by default.** To opt in to tracking upstream `main`, set
  `GREPXCEL_MODEL_AUTOUPDATE=1` — be aware this fetches updated weights from a
  third-party repository.
- Inference runs entirely in-process; **no extracted data ever leaves the
  machine**.
