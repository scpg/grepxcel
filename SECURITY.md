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

- Issues in `tmp-scripts/` (development-only, never installed)
- Vulnerabilities in the local LLM used by `grepxcel suggest` (third-party model)
- Social engineering or phishing

## Security design notes

grepxcel applies several defences against malicious Excel files:

- `defusedxml` for XML parsing (XXE protection)
- Compressed and uncompressed file size limits (ZIP bomb guard)
- Per-cell character length limit (ReDoS guard)
- Formula detection: pattern files must contain plain text only
- `data_only=True` when loading data files (formulas never executed)
