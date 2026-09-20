# Copilot Instructions — grepxcel

These instructions are used by GitHub Copilot when reviewing pull requests and
suggesting code. They encode project conventions that every contributor (human or AI)
must follow.

---

## Project layout (know before you review)

| Directory | Purpose |
|---|---|
| `grepxcel/` | Core package — engine, parser, logger, models, security, utils |
| `grepxcel/examples/` | **User-facing** playground files shipped inside the package via `grepxcel generate-examples`. Must stay in sync with `find_pattern_xlsx()`. |
| `tests/fixtures/` | **Developer-facing** test data — 22 fixtures, never shipped to users |
| `tests/` | pytest suite (unit + integration) |
| `scripts/` | Reusable contributor utilities (tracked, bootstrapped) |
| `tmp.local/` | Throwaway scripts — gitignored, never committed |
| `samples.local/` | Real Excel samples — gitignored, never committed |

---

## Hard rules — flag any violation as a required change

### No inline scripts
`python3 -c "..."`, shell here-docs, and multi-step pipelines that embed logic are
**forbidden**. Any script logic must live in a `.py` file under `scripts/` or
`tmp.local/`. Flag inline scripts in PRs as blocking issues.

### Do not touch user-curated fixture patterns
Files matching `pattern-manual*.xlsx` or `pattern-manual*.csv` anywhere under
`tests/fixtures/` are hand-crafted and **read-only**. A PR that writes, regenerates,
or overwrites these files must be blocked.

### Bootstrap boilerplate is mandatory in every new script
Every `.py` file added to `scripts/` must start with:
```python
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()
```
Every `.py` file added to `tmp.local/` must start with:
```python
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
from _bootstrap import ensure_venv; ensure_venv()
```

### Branch workflow
All feature work lives on `dev`. Direct commits to `main` are not allowed.
PRs must target `main` from `dev` (or a feature branch that merges into `dev` first).

---

## Test conventions

- Tests use `pytest`. Run with `.venv/bin/pytest tests/`.
- CI does **not** install grepxcel as a package. Tests must use `cli.main()` and
  `grepxcel.extract()` directly — no `subprocess` calls to the CLI binary.
- Regression snapshots live in `tests/fixtures/<name>/<name>_snapshot-from-*.json`.
  Do not update them without understanding why the output changed.
- The `examples/` pattern files must stay in sync with fixture patterns.
  Run `scripts/sync_examples.py` after adding or changing a backend pattern.

---

## Security conventions

- All untrusted cell values (from user-supplied Excel files) must be neutralised
  before being written to CSV or any output format (prefix injection characters with
  a tab or similar guard).
- Regex patterns applied to cell content must be guarded by `--max-cell-len` to
  prevent ReDoS.
- File size limits (`--max-size`, `--max-uncompressed`) must be respected on every
  new code path that opens an xlsx file.

---

## Output conventions

Output routing depends on the command:

- **`grepxcel extract`**: `stdout` — JSON only (pipe-friendly); `stderr` — warnings, errors, summary, verbose/debug output. `-q / --quiet` suppresses header and summary but must never suppress warnings or errors.
- **`grepxcel test`, `doctor`, `validate-pattern`, `quickstart`**: primary output → `stdout`; error/usage messages → `stderr`.
- **Side-effect confirmations** (file written, path saved) → `stderr` across all commands.

A PR that sends the test report, doctor output, or validation result to stderr is a routing bug.

---

## Examples vs. fixtures distinction (common source of bugs)

`grepxcel/examples/` patterns are **what new users see first**. Editing them changes
the out-of-box experience. `tests/fixtures/` patterns are CI-only. These are two
completely separate things — do not conflate them in PR descriptions or code comments.
