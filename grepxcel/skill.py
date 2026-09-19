"""`grepxcel generate-skill` — emit a "skill" doc that teaches an AI agent how
and *when* to use grepxcel.

One curated content core is rendered for a chosen target ecosystem. v1 targets:

  - ``claude``     — a SKILL.md (Agent Skill) with YAML frontmatter
  - ``agents-md``  — an AGENTS.md (cross-tool convergence file), plain markdown

The command reference is introspected from the live CLI so the flag/command list
never goes stale; the surrounding prose (when to use, workflow, gotchas) is
curated — a good skill is judgment, not a ``--help`` dump.
"""
from __future__ import annotations

import argparse
import sys

_NAME = 'grepxcel'
_DESCRIPTION = (
    'Extract structured data from Excel (.xlsx) files into JSON using a small, '
    'reusable pattern file. Use when pulling the same fields from many '
    'similarly-structured spreadsheets.'
)

_WHEN_TO_USE = """\
## When to use grepxcel

Use it when you need to pull the **same fields** out of **many** similarly
structured `.xlsx` files (invoices, statements, timesheets, reports) into clean
JSON for a database or pipeline. You write a *pattern* once; it applies to every
file with that layout.

Reach for something else when: it's a one-off read of a single sheet, the files
have no consistent structure, or you just need raw cell values (use a plain
spreadsheet library)."""

_WORKFLOW = """\
## Typical workflow

1. `grepxcel draft data.xlsx -o pattern.xlsx` — get a starter pattern (optional;
   you can also hand-write one).
2. `grepxcel validate-pattern pattern.xlsx -v` — check it parses and see the
   fields/steps before extracting.
3. `grepxcel extract -p pattern.xlsx data.xlsx` — extract → JSON on stdout.
   For many files: `grepxcel extract -p pattern.xlsx data/ -r` (recurse a directory).
4. `grepxcel schema pattern.xlsx -o schema.json` — (optional) a JSON Schema to
   validate the extracted JSON across many files.
5. `grepxcel lint data.xlsx` — inspect a file for issues before extraction.
   Default: compact one-line summary per file. Add `-v` for the full checklist.
   `grepxcel lint data/ -r` scans a whole directory.
6. `grepxcel test -p pattern.xlsx samples/` — run a pattern against many files and
   report which pass/fail. Useful for validating a pattern across a corpus before
   deploying it to a pipeline."""

_MCP_INTEGRATION = """\
## MCP integration (preferred for AI agents)

If grepxcel is configured as an MCP server in your agent, you can call its tools
directly — no shell commands needed. Run `grepxcel mcp-config` to print the config
block to add to your agent. Available MCP tools mirror the CLI:
`extract`, `validate_pattern`, `lint`, `schema`, `doctor`, `generate_examples`.

When MCP is available, prefer it over CLI calls: the agent receives structured
return values rather than parsing terminal output."""

_GOTCHAS = """\
## Gotchas worth knowing

- **Output convention:** `extract` JSON → **stdout** (pipe to `jq` / a file);
  warnings, summary, verbose traces → **stderr**. `lint` output → **stdout**
  (so `grepxcel lint data/ -r | wc -l` counts files correctly).
- A pattern file is `.xlsx` **or** `.csv`. Keywords (`cell:`, `table:`, `var:`…)
  are case-insensitive; an unknown instruction is a hard error (not ignored).
- Empty regex cell ⇒ accept any value of the declared type.
- `--log-format json` writes NDJSON for SIEM/cloud ingestion; extracted cell
  values are **never included** (allow-list construction, not redaction).
- Cloud `draft` backends read keys from `.env` (project-local or
  `~/.config/grepxcel/.env`); `--strict-env` refuses a stray out-of-project `.env`.
- Patterns may declare `config: | pattern.version | N`; absent ⇒ 1."""


def _command_reference(parser: argparse.ArgumentParser | None = None) -> str:
    """Introspect the live CLI for the subcommand list + one-line help, so the
    reference can never drift from the actual commands."""
    if parser is None:
        from .cli import _build_parser
        parser = _build_parser()
    lines = ['## Commands', '']
    for action in parser._actions:
        if isinstance(action, argparse._SubParsersAction):
            for choice in action._choices_actions:
                lines.append(f'- `{_NAME} {choice.dest}` — {choice.help}')
            break
    lines.append('')
    lines.append(f'Run `{_NAME} <command> --help` for full options.')
    return '\n'.join(lines)


def _body(parser: argparse.ArgumentParser | None = None) -> str:
    return '\n\n'.join([
        f'# {_NAME}',
        _DESCRIPTION,
        _WHEN_TO_USE,
        _WORKFLOW,
        _MCP_INTEGRATION,
        _command_reference(parser),
        _GOTCHAS,
    ])


def generate_skill(target: str = 'claude',
                   parser: argparse.ArgumentParser | None = None) -> str:
    """Render the skill doc for *target* ('claude' or 'agents-md')."""
    body = _body(parser)
    if target == 'claude':
        frontmatter = (
            '---\n'
            f'name: {_NAME}\n'
            f'description: {_DESCRIPTION}\n'
            '---\n\n'
        )
        return frontmatter + body + '\n'
    if target == 'agents-md':
        return body + '\n'
    raise ValueError(f"unknown skill target {target!r} (use 'claude' or 'agents-md')")


_DEFAULT_FILENAME = {'claude': 'SKILL.md', 'agents-md': 'AGENTS.md'}


def run_skill(target: str, out_path: str | None = None, out=None) -> int:
    """CLI entry point: write the skill doc to a file or stdout."""
    out = out or sys.stdout
    try:
        text = generate_skill(target)
    except ValueError as exc:
        print(f'Error: {exc}', file=sys.stderr)
        return 1
    if out_path:
        with open(out_path, 'w', encoding='utf-8') as fh:
            fh.write(text)
        print(f'Skill written to: {out_path}', file=sys.stderr)
    else:
        out.write(text)
    return 0
