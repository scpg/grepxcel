# Claude Instructions — grepxcel

## MANDATORY: No inline scripts

**Never** run `python3 -c "..."` or any other inline script technique.
All test, exploration, and one-off scripts **must** be written as files inside `tmp-scripts/`.
Before creating a new script, check if a suitable one already exists in `tmp-scripts/` and reuse it.

This rule has no exceptions.

## Project layout

- `engine/`       — core package (engine, parser, logger, models, security, utils)
- `tests/`        — pytest suite; fixtures are generated xlsx files in `tests/fixtures/`
- `tmp-scripts/`  — throwaway scripts (gitignored); put all ad-hoc code here
- `samples.local/`— real Excel samples (gitignored, never committed)
- `logs/`         — runtime log output (gitignored)
- `output/`       — JSON extraction output (gitignored)

## Running tests

```bash
source .venv/bin/activate
pytest tests/
```

## Running the CLI

```bash
source .venv/bin/activate
grepxcel -p pattern.xlsx data.xlsx
grepxcel -p pattern.xlsx data.xlsx -v --output output/
grepxcel -p pattern.xlsx data.xlsx -vv            # debug: anchor probes
grepxcel -p pattern.xlsx data.xlsx -d             # same as -vv
grepxcel -p pattern.xlsx data.xlsx --sheet Sheet2
```

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
