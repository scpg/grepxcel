# Current status

A living snapshot of where the project stands, written for whoever picks it up
next — human or agent. More than one assistant works on this repository
(Claude Code and Devin, at least), and neither sees the other's notes. This file
is the shared one. **Update it when you change any of the facts below.**

Last updated: **2026-10-05**

---

## Release state

| | |
|---|---|
| Version in `grepxcel/__init__.py` | `0.5.0` |
| Published to PyPI | **no** — `v0.5.0` is not tagged yet |
| Latest TestPyPI build | `0.5.0rc46` (verified: build, publish, install-and-extract smoke test) |
| `main` | still on v0.4.1 |
| `dev` | ahead of `main` by the whole 0.5.0 line plus the wizard security work |

**0.5.0 carries two breaking changes.** Both are in `CHANGELOG.md`, and anyone
given an earlier RC needs telling:

1. Temporal coercion changes extraction output — a date/time stored as text now
   yields a real temporal value instead of the raw string plus a warning.
2. Macro-bearing workbooks are refused by *content*, so a renamed `.xlsm` that
   used to extract now fails.

---

## Open pull requests

| PR | Direction | Status |
|---|---|---|
| **#79** | `dev` → `main` | **This is the release PR.** Its title ("fix: 10 code-review bugs") is from 2026-09-20 and describes a small fraction of what it now ships — ~100 commits. Retitle before merging; the release notes are generated from it. |
| **#87** | `dev` | `python-multipart` in `[web]` + a `test-web` CI job. Green. Closes the gap described under "CI blind spot" below. Merge. |

---

## Web wizard security — a two-part finding, now fully closed

Recorded in full because it was fixed in two separate passes by two different
agents, and the second half was easy to miss.

**Part 1 — CSRF and DNS rebinding (PR #86, merged).** Only `/api/shutdown`
checked `Origin`. Every other endpoint parsed its body with `request.json()`,
which ignores `Content-Type`, so a page on another origin could submit a
`<form enctype="text/plain">` containing valid JSON — a "simple" request needing
no CORS preflight. `/api/save` that way wrote or overwrote a file of any name in
the data file's folder. Any `Host` was accepted, so the session was also
readable through DNS rebinding. Now: loopback `Host` required on every request,
non-loopback `Origin` rejected on state-changing requests, and bodies must be
`application/json` (multipart only for the upload endpoint).

**Part 2 — stored XSS (fixed here).** #86 closed the delivery vector but not the
sink. `/api/classify`, `/api/classify-batch` and `/api/note` validated `action`
against an allow-list while accepting *any* non-empty string as a cell ref, and
`wizard.js` interpolated that value into `onclick="jumpToCell('<ref>')"`. An
apostrophe closed the JS string literal. `escHtml()` would not have helped — it
escaped `& < > "` but not `'`, and the value sat inside single quotes.

Three layers now:

1. `_require_cell_ref()` validates against `^[A-Z]{1,3}[1-9][0-9]{0,6}\Z` at
   every entry point, so a malformed ref never reaches session state or the log.
2. The chip carries its ref in a `data-ref` attribute with a delegated click
   listener — no JavaScript is built from data at all.
3. `escHtml()` now escapes the apostrophe, for every other caller.

**Do not rely on `.upper()`.** Upper-casing the ref incidentally mangled most
script payloads, because JS identifiers are case-sensitive. That is an accident,
not a control; nothing documented it and nothing tested it.

`tests/unit/test_wizard_cell_ref_validation.py` pins all three layers, including
a test that bans inline `onclick` handlers built by interpolation anywhere in
`wizard.js`.

---

## CI blind spot — important if you are adding wizard tests

`requirements-dev.txt` does **not** include FastAPI, and the `test` job installs
only `requirements.txt -r requirements-dev.txt`. Every wizard test therefore
**skips in CI**. PR #86's security tests were merged without ever having run
there.

**PR #87 fixes this** by adding a `test-web` job that installs `.[web]`. Until it
lands, a green CI run says nothing about wizard behaviour — run
`.venv/bin/pytest tests/unit/test_wizard*.py` locally.

---

## Known environment traps

- **A pytest run with no `N passed` summary line is never a pass**, whatever the
  exit code says. A hard exit in a daemon thread once killed the process after
  the last test and before pytest's epilogue, so the suite exited 0 while a test
  was failing. Guarded now by a session fixture plus an AST test, but read the
  summary line rather than trusting `$?`.
- **Starlette's TestClient cannot drive a bracketed IPv6 host.** It splits the
  netloc on `:` and calls `int()`, so `base_url='http://[::1]:8765'` raises
  `ValueError` before the request reaches the app. Version-dependent — it passed
  in CI and failed locally on starlette 1.6.0 / httpx2 2.12.0. Assert
  `_is_loopback_host()` directly instead of routing IPv6 through the client.
- **Temp directories must clean themselves up however the run ends** —
  `mkdtemp` + `atexit` in Python, `mktemp -d` + `trap … EXIT` in shell. Two
  leaks have already exhausted the inodes on a tmpfs `/tmp` (with `df -h` still
  showing plenty of free space; check `df -i`). Logs that must be kept do not
  belong in `/tmp` — they go in `logs/`.
- `$GITHUB_TOKEN` in the maintainer's shell is invalid and shadows a working
  keyring credential; `git push` and `gh` need `env -u GITHUB_TOKEN`.

---

## Open work, roughly by value

1. **Parser fuzzing.** For a file-format parser this buys more real security than
   any scanner integration. Not started.
2. **Schema strictness decisions** —
   `docs/superpowers/specs/2026-09-27-schema-strictness-decisions.md` needs a
   per-category call, including the boolean widening that currently makes a
   generated schema accept `-1` and `'maybe'` while extraction rejects them.
3. **Tag `v0.5.0`** once #79 is retitled and merged.
4. **Static tutorial site** — Astro on S3 + CloudFront, three environments, CDK.
   Setup checklist in `infra/AWS-SETUP.md`; blocked on a domain name. The
   screenshot generator (`scripts/capture_screenshots.py`) is written and
   verified byte-reproducible across runs.
5. **Two GitHub secret-scanning toggles** are UI-only:
   `secret_scanning_non_provider_patterns` and `..._validity_checks`.

### Deliberately rejected

**VirusTotal integration.** Its public API terms forbid use "in commercial
products or services" and "in business workflows that do not contribute new
files", so compliance would require *uploading* customer spreadsheets — which
are exactly the sensitive data. Hash lookups are also uninformative here, since
the files are unique internal reports. If a scan hook is wanted, make it
scanner-agnostic (a command the operator configures) rather than binding the
project to one vendor.
