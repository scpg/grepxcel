# Claude Instructions — grepxcel

## MANDATORY: No inline scripts

**Never** run `python3 -c "..."` or any other inline script technique.
All test, exploration, and one-off scripts **must** be written as files inside `tmp.local/`.
Before creating a new script, check if a suitable one already exists in `tmp.local/` and reuse it.

This rule has no exceptions.

## Project layout

- `engine/`       — core package (engine, parser, logger, models, security, utils)
- `tests/`        — pytest suite; fixtures are generated xlsx files in `tests/fixtures/`
- `tmp.local/`  — throwaway scripts (gitignored); put all ad-hoc code here
- `samples.local/`— real Excel samples (gitignored, never committed)
- `logs/`         — runtime log output (gitignored)
- `output/`       — JSON extraction output (gitignored)

## Running tests

```bash
.venv/bin/pytest tests/
```

## Running the CLI

```bash
.venv/bin/grepxcel -p pattern.xlsx data.xlsx
.venv/bin/grepxcel -p pattern.xlsx data.xlsx -v --output output/
.venv/bin/grepxcel -p pattern.xlsx data.xlsx -vv            # debug: anchor probes
.venv/bin/grepxcel -p pattern.xlsx data.xlsx -d             # same as -vv
.venv/bin/grepxcel -p pattern.xlsx data.xlsx --sheet Sheet2
```

> **Note on venv activation:** `source .venv/bin/activate` is not needed.
> Both `grepxcel` and `pytest` are installed with a shebang pointing directly to
> `.venv/bin/python3`, so calling them by their full path is fully self-contained.
> Claude Code allowlists `Bash(.venv/bin/pytest *)` and `Bash(.venv/bin/grepxcel *)`
> in `.claude/settings.local.json` — `source` cannot be allowlisted because it is a
> shell builtin that modifies shell state and Claude Code intentionally blocks
> compound commands (`&&`, `||`, `;`) from matching permission patterns.

### CLI parameters

| Flag | Default | Purpose |
|---|---|---|
| `-p FILE` | required | Pattern xlsx file |
| `-o DIR` | — | Write JSON to directory (stdout if omitted) |
| `-l FILE` | — | Append structured log to file |
| `-v` / `-vv` | off | Verbosity: step-by-step / anchor probes |
| `-d` / `--debug` | off | Same as `-vv` |
| `--max-size MB` | 5 | Compressed file size limit |
| `--max-uncompressed MB` | 50 | Uncompressed ZIP content limit (ZIP bomb guard) |
| `--max-cell-len N` | 1000 | Max cell chars fed to regex (ReDoS guard) |
| `--sheet NAME_OR_INDEX` | active | Sheet name or 0-based index to process |

### Output convention

- **stdout** — extracted JSON (pipe-friendly)
- **stderr** — warnings, errors, summary, verbose/debug output
