# Claude Instructions — grepxcel

## MANDATORY: No inline scripts

**Never** run `python3 -c "..."` or any other inline script technique.
All test, exploration, and one-off scripts **must** be written as files inside `tmp.local/`.
Before creating a new script, check `scripts/` (tracked utilities) and `tmp.local/` (local throwaway) — reuse what already exists.

This rule has no exceptions.

## Project layout

- `grepxcel/`     — core package (engine, parser, logger, models, security, utils)
- `tests/`        — pytest suite; fixtures are generated xlsx files in `tests/fixtures/`
- `scripts/`      — reusable utility scripts for contributors (tracked, documented)
- `tmp.local/`    — throwaway scripts (gitignored); put all ad-hoc code here
- `samples.local/`— real Excel samples (gitignored, never committed)
- `logs/`         — runtime log output (gitignored)
- `output/`       — JSON extraction output (gitignored)

## Before writing a new script

Check `scripts/` first — a suitable utility may already exist there.
If the script is reusable and useful to contributors, put it in `scripts/`.
If it is exploratory or one-off, put it in `tmp.local/`.

## Bootstrap boilerplate (mandatory for every new script)

Every Python script in `scripts/` or `tmp.local/` **must** start with the following venv
bootstrap block so users can run it as plain `python3 <script>` without activating the venv:

**For `scripts/` scripts:**
```python
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()
```

**For `tmp.local/` scripts:**
```python
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
from _bootstrap import ensure_venv; ensure_venv()
```

`ensure_venv()` re-execs the script with `.venv/bin/python3` transparently.
If the venv does not exist it prints a clear error with setup instructions and exits.

## Running tests

```bash
.venv/bin/pytest tests/
```

## Running the CLI

The CLI is subcommand-based: `extract`, `validate-pattern`, `draft`, `docs`, `doctor` (run `grepxcel <cmd> --help`).

```bash
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx -v            # per-field trace: field ← B1 = value ✓/✗
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx -vv           # debug: anchor probes
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx -d            # same as -vv
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx -o output/    # write JSON to a directory
.venv/bin/grepxcel extract -p pattern.xlsx data.xlsx --sheet Sheet2
.venv/bin/grepxcel extract -p pattern.xlsx jan.xlsx feb.xlsx       # multiple data files
.venv/bin/grepxcel validate-pattern pattern.xlsx                   # check a pattern is valid (no extraction)
.venv/bin/grepxcel validate-pattern pattern.csv -v                # + parsed fields & extraction steps
.venv/bin/grepxcel draft data.xlsx                                 # draft a starter pattern (local LLM)
.venv/bin/grepxcel draft data.xlsx --ca-bundle corp-ca.pem         # behind a corporate TLS-inspection proxy
.venv/bin/grepxcel docs                                            # write pattern-reference.xlsx
.venv/bin/grepxcel doctor [extract|draft|all]                      # preflight: deps, keys, model, proxy/TLS
```

> The pattern file may be `.xlsx` **or** `.csv`. `suggest` is a hidden
> backward-compatible alias for `draft`.

> **Note on venv activation:** `source .venv/bin/activate` is not needed.
> Both `grepxcel` and `pytest` are installed with a shebang pointing directly to
> `.venv/bin/python3`, so calling them by their full path is fully self-contained.
> Claude Code allowlists `Bash(.venv/bin/pytest *)` and `Bash(.venv/bin/grepxcel *)`
> in `.claude/settings.local.json` — `source` cannot be allowlisted because it is a
> shell builtin that modifies shell state and Claude Code intentionally blocks
> compound commands (`&&`, `||`, `;`) from matching permission patterns.

### CLI parameters (`grepxcel extract`)

| Flag | Default | Purpose |
|---|---|---|
| `-p, --pattern FILE` | required | Pattern file (`.xlsx` or `.csv`) |
| `FILE...` (positional) | required | One or more data `.xlsx` files |
| `-o, --output DIR` | — | Write JSON to directory (stdout if omitted) |
| `-l, --log FILE` | — | Append structured log to file |
| `-v` / `-vv` | off | Verbosity: per-field trace / anchor probes |
| `-d, --debug` | off | Same as `-vv` |
| `--sheet NAME_OR_INDEX` | active | Sheet name or 0-based index to process |
| `--all-sheets` | off | Process every sheet; output keyed by sheet name |
| `--format {nested,legacy}` | nested | Output shape |
| `--max-rows N` | 2048 | Max data-sheet rows (Excel max 1,048,576) |
| `--max-columns N` | 1024 | Max data-sheet columns (Excel max 16,384) |
| `--max-cell-len N` | 1000 | Max cell chars fed to regex (ReDoS guard) |
| `--max-size MB` | 5 | Compressed file size limit |
| `--max-uncompressed MB` | 50 | Uncompressed ZIP content limit (ZIP-bomb guard) |

### Output convention

- **stdout** — extracted JSON (pipe-friendly)
- **stderr** — warnings, errors, summary, verbose/debug output

## Branch workflow

Always work on `dev`. Never commit or push directly to `main`.

```bash
# 1. Make changes on dev, commit, push
git add <files>
git commit -m "..."
git push origin dev

# 2. Open a PR from dev → main
gh pr create --base main --head dev --title "Short title" --body "What and why."

# 3. Check CI status
gh run list --limit 5

# 4. Once CI is green, merge the PR
gh pr merge <PR-number> --squash

# 5. Sync local main (optional)
git fetch origin main
```

`gh pr *` and `gh run *` are allowlisted in `.claude/settings.local.json`.
