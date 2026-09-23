# PYTHON_ARGCOMPLETE_OK
"""
Command-line interface for grepxcel.

Usage:
    grepxcel extract -p pattern.xlsx data.xlsx
    grepxcel extract -p pattern.xlsx data1.xlsx data2.xlsx data3.xlsx
    grepxcel extract -p pattern.xlsx data.xlsx -v --output results/
    grepxcel extract -p pattern.xlsx data.xlsx --log logs/run.log
    grepxcel draft data.xlsx -o draft_pattern.xlsx
    grepxcel draft data.xlsx -v
"""

import argparse
import datetime
import json
import os
import sys

from .color import MARK_FAIL, MARK_OK, MARK_WARN, colorize_marks, paint, should_color
from .engine import Engine
from .logger import Logger, VerbosityLevel
from .utils import flatten_table_instances as _flatten_table_instances


# ── JSON serialisation ────────────────────────────────────────────────────────

def _json_default(obj):
    """Serialise types that json.dump does not handle natively.

    Excel cells surface as several datetime flavours: dates/datetimes and
    time-of-day all have ``.isoformat()``; durations ([h]:mm cells) come through
    as ``timedelta``, which has no isoformat, so render it as ``str`` (e.g.
    ``"8:30:00"``).
    """
    if isinstance(obj, (datetime.datetime, datetime.date, datetime.time)):
        return obj.isoformat()
    if isinstance(obj, datetime.timedelta):
        return str(obj)
    raise TypeError(f'Type {type(obj).__name__} is not JSON serialisable')


# ── Shared argument helpers ───────────────────────────────────────────────────

def _add_security_args(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        '--max-size',
        type=float, default=5, metavar='MB',
        help='Maximum compressed file size in MB (default: 5)',
    )
    p.add_argument(
        '--max-uncompressed',
        type=float, default=50, metavar='MB',
        help='Maximum uncompressed ZIP content in MB (default: 50)',
    )


def _add_strict_arg(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        '--strict',
        action='store_true',
        help='Exit with code 2 if any defined field is missing (null) in the '
             'output. Use in pipelines to catch incomplete extractions.',
    )


def _add_sheet_arg(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        '--sheet',
        metavar='NAME_OR_INDEX',
        help='Sheet to use: name (e.g. Sheet2) or 0-based index (default: active sheet)',
    )


def _resolve_sheet(args) -> str | None:
    """Return the --sheet value unchanged (string or None).

    Resolution happens downstream: a string is matched as a sheet NAME first and
    falls back to a 0-based index only when no sheet has that name. This lets a
    numerically-named sheet (e.g. '2025') be selected by name, while '--sheet 0'
    still works as an index. Converting to int here would wrongly turn the name
    '2025' into index 2025.
    """
    return getattr(args, 'sheet', None)


# ── Argument parser ───────────────────────────────────────────────────────────

class _GroupedHelpFormatter(argparse.RawDescriptionHelpFormatter):
    """Like RawDescriptionHelpFormatter but hides the bare subparsers metavar entry."""

    def _format_action(self, action):
        if isinstance(action, argparse._SubParsersAction):
            return ''
        return super()._format_action(action)

def _commands_help() -> str:
    """Grouped command listing with color when stdout is a TTY."""
    c = should_color(sys.stdout)

    def _sec(s):
        return paint(s, 'cyan', c)

    return (
        f"{_sec('extract & validate:')}\n"
        "  extract            Extract data from Excel files using a pattern file\n"
        "  validate-pattern   Validate a pattern file before running extraction\n"
        "\n"
        f"{_sec('onboarding:')}\n"
        "  quickstart         Guided tutorial — learn grepxcel in your terminal\n"
        "  web-wizard         Build a pattern file visually in your browser\n"
        "  generate-examples  Create ready-to-run example files in a local directory\n"
        "  docs               Write pattern-reference.xlsx + grepxcel-guide.docx\n"
        "\n"
        f"{_sec('automation:')}\n"
        "  watch              Monitor a directory and extract new .xlsx files automatically\n"
        "  test               Test a pattern's reliability against a sample directory\n"
        "\n"
        f"{_sec('AI & MCP:')}\n"
        "  draft              Draft a starter pattern file using a local LLM\n"
        "  generate-skill     Write an AI-agent skill doc (Claude, Cursor, Copilot, Windsurf…)\n"
        "  mcp                Start the grepxcel MCP server (stdio transport)\n"
        "  mcp-config         Print the MCP server config for your AI agent\n"
        "\n"
        f"{_sec('inspection:')}\n"
        "  lint               Inspect an Excel file for potential extraction issues\n"
        "  schema             Generate a JSON Schema from a pattern file\n"
        "\n"
        f"{_sec('compliance & ops:')}\n"
        "  sbom               Generate a CycloneDX 1.6 SBOM for this installation\n"
        "  doctor             Check the environment is ready (deps, API keys, model)\n"
        "  self-test          Post-install check: parse and extract the 4 bundled examples\n"
        "\n"
        "Run 'grepxcel <command> --help' for per-command options.\n"
        "Run 'grepxcel -h -h' for a synopsis of every command's options."
    )


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='grepxcel',
        description='Extract structured data from Excel files using a pattern.',
        formatter_class=_GroupedHelpFormatter,
        epilog='Shell TAB completion: eval "$(register-python-argcomplete grepxcel)"',
    )
    from . import __version__
    p.add_argument(
        '-v', '--version',
        action='version',
        version=f'%(prog)s {__version__}',
    )
    sub = p.add_subparsers(
        dest='command',
        metavar='COMMAND',
        title='commands',
        description=_commands_help(),
    )
    sub.required = True

    _add_extract_subparser(sub)
    _add_validate_subparser(sub)
    _add_watch_subparser(sub)
    _add_draft_subparser(sub)
    _add_web_wizard_subparser(sub)
    _add_docs_subparser(sub)
    _add_lint_subparser(sub)
    _add_schema_subparser(sub)
    _add_skill_subparser(sub)
    _add_examples_subparser(sub)
    _add_sbom_subparser(sub)
    _add_mcp_subparser(sub)
    _add_mcp_config_subparser(sub)
    _add_self_test_subparser(sub)
    _add_doctor_subparser(sub)
    _add_quickstart_subparser(sub)
    _add_test_subparser(sub)
    return p


def _add_validate_subparser(sub) -> None:
    p = sub.add_parser(
        'validate-pattern',
        help='Check a pattern file (.xlsx or .csv) is valid to use',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Parses the pattern with the same rules extraction uses (structure, types,
multiplicities, regex safety, comments) and also flags an empty extraction
sequence and references to undefined fields. Exits non-zero if any file is
invalid.

verbosity:
  (none)  print ✓ VALID / ✗ INVALID + any warnings or errors
  -q      silent on success; only print warnings / errors (useful in CI)
  -v      + parsed config, fields, and extraction sequence

examples:
  grepxcel validate-pattern pattern.xlsx
  grepxcel validate-pattern pattern.csv -v        # + parsed fields & steps
  grepxcel validate-pattern a.xlsx b.csv          # validate several
  grepxcel validate-pattern pattern.xlsx -q       # silent success, CI-friendly
        """,
    )
    p.add_argument('files', nargs='+', metavar='FILE',
                   help='Pattern file(s) to validate (.xlsx or .csv)')
    p.add_argument('-v', '--verbose', action='store_true',
                   help='Print the parsed config, fields, and extraction sequence')
    p.add_argument('-q', '--quiet', action='store_true',
                   help='Silent on success — only print warnings or errors; '
                        'exit code is unchanged')


def _add_lint_subparser(sub) -> None:
    p = sub.add_parser(
        'lint',
        help='Inspect an Excel file for potential extraction issues',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Checks an Excel data file before extraction: format, encryption/IRM,
sheet dimensions (declared vs real extent), merged cells, formula cells,
and known corporate-environment issues. Reports ✓/⚠/✗/ℹ.

examples:
  grepxcel lint data.xlsx
  grepxcel lint jan.xlsx feb.xlsx        # lint several files
  grepxcel lint data/                    # lint all Excel files in a directory
  grepxcel lint data/ -r                 # recurse into subdirectories
  grepxcel lint data/ -r -v              # full checklist detail per file
        """,
    )
    p.add_argument('files', nargs='+', metavar='FILE_OR_DIR',
                   help='Excel file(s) or director(ies) to inspect')
    p.add_argument('-r', '--recursive', action='store_true',
                   help='Recurse into subdirectories when a directory is given')
    p.add_argument('-v', '--verbose', action='store_true',
                   help='Show full checklist detail (default: one summary line per file)')


def _add_schema_subparser(sub) -> None:
    p = sub.add_parser(
        'schema',
        help='Generate a JSON Schema from a pattern file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Generates a JSON Schema (draft 2020-12) that describes the extraction
output for the given pattern. Use it to validate extracted JSON with
any standard JSON Schema validator.

examples:
  grepxcel schema pattern.xlsx
  grepxcel schema pattern.xlsx -o schema.json
  grepxcel schema pattern.csv
        """,
    )
    p.add_argument('files', nargs='+', metavar='PATTERN',
                   help='Pattern file(s) to generate schema for (.xlsx or .csv)')
    p.add_argument('-o', '--output', metavar='FILE',
                   help='Write schema to file (default: stdout)')
    p.add_argument(
        '--force',
        action='store_true',
        help='Overwrite an existing output file without prompting',
    )


def _add_skill_subparser(sub) -> None:
    p = sub.add_parser(
        'generate-skill',
        help='Write an AI-agent skill doc for Claude, Cursor, Copilot, Windsurf…',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Writes a skill/rules doc that teaches an AI agent how and when to use grepxcel.
The command list is introspected from the live CLI; the prose is curated.

targets:
  claude     SKILL.md                      — Claude Code / Claude Desktop
  cursor     .cursor/rules/grepxcel.mdc   — Cursor IDE (glob-triggered MDC rule)
  agents-md  AGENTS.md                    — OpenAI Codex + any tool that reads AGENTS.md

examples:
  grepxcel generate-skill                              # Claude SKILL.md to stdout
  grepxcel generate-skill -o SKILL.md
  grepxcel generate-skill --target cursor -o .cursor/rules/grepxcel.mdc
  grepxcel generate-skill --target agents-md -o AGENTS.md
        """,
    )
    p.add_argument('--target',
                   choices=['claude', 'cursor', 'agents-md'],
                   default='claude',
                   help='AI engine to target (default: claude)')
    p.add_argument('-o', '--output', metavar='FILE',
                   help='Write the skill doc to FILE (default: stdout)')
    p.add_argument(
        '--force',
        action='store_true',
        help='Overwrite an existing output file without prompting',
    )


def _add_examples_subparser(sub) -> None:
    p = sub.add_parser(
        'generate-examples',
        help='Create ready-to-run example files in a local directory',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Copies 4 bundled examples (pattern + data xlsx) into a local directory so
you can immediately try grepxcel without needing your own Excel files.
Each example includes a README.txt with the exact commands to run.

examples:
  grepxcel generate-examples                         # → ./grepxcel-examples/
  grepxcel generate-examples -o my-examples
        """,
    )
    p.add_argument(
        '-o', '--output',
        metavar='DIR',
        default='grepxcel-examples',
        help='Directory to create (default: ./grepxcel-examples/)',
    )
    p.add_argument(
        '--force',
        action='store_true',
        help='Overwrite an existing non-empty directory without prompting',
    )


def _add_sbom_subparser(sub) -> None:
    p = sub.add_parser(
        'sbom',
        help='Generate a CycloneDX 1.6 SBOM for this installation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Generate a CycloneDX 1.6 Software Bill of Materials (SBOM) listing grepxcel
and every transitive dependency with PURLs, licenses, and hashes.

Uses only stdlib — no external tool required. The output is valid CycloneDX
JSON accepted by Dependency-Track, Grype, and other SBOM consumers.

examples:
  grepxcel sbom                        # print to stdout
  grepxcel sbom -o sbom.cdx.json       # write to file
        """,
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        help='Write SBOM to file (default: stdout)',
    )
    p.add_argument(
        '--force',
        action='store_true',
        help='Overwrite an existing output file without prompting',
    )


def _add_mcp_subparser(sub) -> None:
    p = sub.add_parser(
        'mcp',
        help='Start the MCP server (stdio transport)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Starts grepxcel as a Model Context Protocol (MCP) server using stdio
transport. AI agents (Claude Code, Claude Desktop, Cursor, etc.) can
call grepxcel tools directly: extract, validate-pattern, lint, schema,
docs, doctor, and generate-examples.

Requires: pip install 'grepxcel[mcp]'

To see the config to add to your AI agent, run:
  grepxcel mcp-config
        """,
    )
    p.add_argument('-v', '--verbose', action='count', default=0,
                   help='-v: log each tool call to stderr; -vv: also log return values')


def _add_mcp_config_subparser(sub) -> None:
    p = sub.add_parser(
        'mcp-config',
        help='Print the MCP server config for your AI agent',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Detects how grepxcel is installed and prints the JSON config block
to add to your AI agent's MCP configuration file.

examples:
  grepxcel mcp-config                          # Claude Code (default)
  grepxcel mcp-config --target claude-desktop
  grepxcel mcp-config --target cursor
        """,
    )
    p.add_argument(
        '--target',
        choices=['claude-code', 'claude-desktop', 'cursor'],
        default='claude-code',
        help='Config format for your AI agent (default: claude-code)',
    )


def _add_self_test_subparser(sub) -> None:
    p = sub.add_parser(
        'self-test',
        help='Post-install check: parse and extract the 4 bundled examples via the Python API',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Runs 10 checks (validate + extract for each bundled example) using only
the installed package — no external files, no network, no pytest needed.

For a full CLI and integration check, see the CI smoke test suite
(tests/smoke/) which runs automatically after every publish.

examples:
  grepxcel self-test        # run all checks, exit 0 on pass
  grepxcel self-test -v     # show full traceback on failure

Exits 0 when every check passes; exits 1 on any failure.
        """,
    )
    p.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Show full tracebacks on failure',
    )


def _add_doctor_subparser(sub) -> None:
    p = sub.add_parser(
        'doctor',
        help='Check the environment is ready (deps, API keys, model, proxy/TLS)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  grepxcel doctor              # check everything (extract + draft + proxy/TLS)
  grepxcel doctor extract      # only what 'extract' needs (offline)
  grepxcel doctor draft        # only what 'draft' needs (local + cloud + proxy)

Exits non-zero if the selected area has a blocking (✗) problem.
        """,
    )
    p.add_argument(
        'area',
        nargs='?',
        choices=['extract', 'draft', 'all'],
        default='all',
        help="Which area to check: extract, draft, or all (default: all)",
    )
    p.add_argument(
        '--no-probe',
        action='store_true',
        help='Skip the live TLS handshake probe (offline / faster)',
    )
    p.add_argument(
        '--strict-env',
        action='store_true',
        help='Refuse a .env from outside the current project (also '
             'GREPXCEL_STRICT_ENV); the per-user config-dir .env stays allowed',
    )


def _add_web_wizard_subparser(sub) -> None:
    p = sub.add_parser(
        'web-wizard',
        help='Build a pattern file visually in your browser (mouse-friendly)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Opens a local web server and launches your default browser.  Click cells in
the spreadsheet grid to classify them as Label / Value / Header / Table /
Ignore.  Works with a mouse — no keyboard shortcuts required.

Requires: pip install "grepxcel[web]"

examples:
  grepxcel web-wizard data.xlsx
  grepxcel web-wizard data.xlsx -p existing-pattern.xlsx
  grepxcel web-wizard data.xlsx --port 9000
  grepxcel web-wizard data.xlsx --no-browser
        """,
    )
    p.add_argument('file', metavar='FILE',
                   help='Excel data file to inspect')
    p.add_argument('-p', '--pattern', metavar='FILE', default=None,
                   help='Pre-populate from an existing pattern file (.xlsx or .csv)')
    p.add_argument('--port', metavar='PORT', type=int, default=8765,
                   help='Local port to listen on (default: 8765)')
    p.add_argument('--no-browser', action='store_true',
                   help='Do not automatically open a browser window')
    p.add_argument('--sheet', metavar='NAME_OR_INDEX', default=None,
                   help='Sheet to open on startup — name or 0-based index (default: active sheet)')
    p.add_argument('--max-rows', metavar='N', type=int, default=150,
                   help='Maximum rows to display in the grid (default: 150)')
    p.add_argument('--max-cols', metavar='N', type=int, default=40,
                   help='Maximum columns to display in the grid (default: 40)')
    p.add_argument('--max-size', type=float, default=5, metavar='MB',
                   help='Compressed file size limit in MB (default: 5)')
    p.add_argument('--max-uncompressed', type=float, default=50, metavar='MB',
                   help='Uncompressed content size limit in MB (default: 50)')


def _add_quickstart_subparser(sub) -> None:
    sub.add_parser(
        'quickstart',
        help='Guided tutorial — learn grepxcel in your terminal',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Prints a step-by-step guide covering what grepxcel does, how to create
your first pattern, and how to run your first extraction. No files are
created or modified.

examples:
  grepxcel quickstart
        """,
    )


def _add_test_subparser(sub) -> None:
    p = sub.add_parser(
        'test',
        help='Run a pattern against a directory of .xlsx files and report reliability',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Run the pattern against every .xlsx in DIRECTORY and report how many files
each field was extracted from. Useful for CI and pattern development.

exit codes:
  0  — all files passed (every field extracted cleanly)
  1  — some files had warnings or partial extraction
  2  — one or more files failed completely (extraction error)

examples:
  grepxcel test -p pattern.xlsx samples/
  grepxcel test -p pattern.xlsx samples/ --recursive
  grepxcel test -p pattern.xlsx samples/ --format json
  grepxcel test -p pattern.xlsx samples/ --strict
        """,
    )
    p.add_argument(
        '-p', '--pattern',
        metavar='PATTERN',
        required=True,
        help='Pattern file (.xlsx or .csv)',
    )
    p.add_argument(
        'directory',
        metavar='DIRECTORY',
        help='Directory containing .xlsx test files',
    )
    p.add_argument(
        '-r', '--recursive',
        action='store_true',
        default=False,
        help='Recurse into subdirectories',
    )
    p.add_argument(
        '--format',
        choices=['human', 'json'],
        default='human',
        help='Output format: human (default) or json',
    )
    p.add_argument(
        '--strict',
        action='store_true',
        default=False,
        help='Treat any missing field as a failure (exit 2)',
    )
    p.add_argument(
        '-v', '--verbose',
        action='count',
        default=0,
        help='-v: add field reliability table  -vv: full per-file detail',
    )
    p.add_argument(
        '--no-color',
        action='store_true',
        default=False,
        help='Disable emoji/color in human output',
    )
    _add_sheet_arg(p)
    _add_security_args(p)


def _add_extract_subparser(sub) -> None:
    p = sub.add_parser(
        'extract',
        help='Extract data from Excel files using a pattern file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
verbosity:
  -q      warnings + errors only — no ENGINE START header or summary stats
  (none)  warnings + summary on stderr, extracted JSON on stdout
  -v      + per-field trace: 'field ← B1 = value ✓/✗' for cells and table data
  -vv     + every anchor probe and rejection reason
  -d      same as -vv (debug mode)

examples:
  grepxcel extract -p pattern.xlsx report.xlsx
  grepxcel extract -p pattern.xlsx jan.xlsx feb.xlsx mar.xlsx
  grepxcel extract -p pattern.xlsx data.xlsx -q          # clean output
  grepxcel extract -p pattern.xlsx data.xlsx --format csv -q
  grepxcel extract -p pattern.xlsx data.xlsx -v
  grepxcel extract -p pattern.xlsx data.xlsx --output results/
  grepxcel extract -p pattern.xlsx data.xlsx --log logs/run.log --output results/
        """,
    )
    p.add_argument(
        '-p', '--pattern',
        required=True,
        metavar='FILE',
        help='Pattern Excel file that describes the layout to extract',
    )
    p.add_argument(
        'files',
        nargs='+',
        metavar='FILE_OR_DIR',
        help='Data Excel files or directories to process (.xlsx)',
    )
    p.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Recurse into subdirectories when a directory is given',
    )
    p.add_argument(
        '--max-files',
        type=int, default=_DEFAULT_MAX_FILES, metavar='N',
        help=f'Safety cap on total files to process (default: {_DEFAULT_MAX_FILES})',
    )
    p.add_argument(
        '-v', '--verbose',
        action='count', default=0,
        help='Increase verbosity (-v: step-by-step, -vv: anchor probes)',
    )
    p.add_argument(
        '-d', '--debug',
        action='store_true',
        help='Debug output (anchor probes and rejections); equivalent to -vv',
    )
    p.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress the ENGINE START header and summary stats; '
             'only warnings and errors are printed to stderr. '
             'Useful when piping output or calling from scripts.',
    )
    p.add_argument(
        '-o', '--output',
        metavar='DIR',
        help='Write extracted data as JSON into DIR  (<datafile-stem>.json per file)',
    )
    p.add_argument(
        '-l', '--log',
        metavar='FILE',
        help='Append a log to FILE (text by default; NDJSON with --log-format json)',
    )
    p.add_argument(
        '--log-format',
        choices=['text', 'json'], default='text',
        help='Log file format: text (human, shows values) or json (NDJSON for '
             'SIEM/cloud — never contains extracted cell values). Default: text',
    )
    p.add_argument(
        '--max-cell-len',
        type=int, default=1000, metavar='CHARS',
        help='Maximum cell character length passed to regex matching (default: 1000)',
    )
    p.add_argument(
        '--max-rows',
        type=int, default=2048, metavar='N',
        help='Maximum rows in the data sheet (default: 2048; Excel max is 1048576, '
             'untested above the default)',
    )
    p.add_argument(
        '--max-columns',
        type=int, default=1024, metavar='N',
        help='Maximum columns in the data sheet (default: 1024; Excel max is 16384, '
             'untested above the default)',
    )
    p.add_argument(
        '--format',
        choices=['nested', 'json', 'legacy', 'csv', 'xlsx'], default='nested',
        help='Output format: nested / json (default, same format; json is the canonical name), '
             'legacy (deprecated), csv, or xlsx (colored Excel report, requires -o)',
    )
    p.add_argument(
        '--csv-delimiter',
        default=',',
        metavar='CHAR',
        help='Field delimiter for --format csv (default: comma)',
    )
    p.add_argument(
        '--csv-mode',
        choices=['extended', 'table'],
        default='extended',
        help='--format csv output mode: extended (default, scalars repeated per row) '
             'or table (table columns only, scalars omitted)',
    )
    p.add_argument(
        '--csv-table',
        type=int,
        metavar='N',
        default=None,
        help='--format csv: export table number N (1-based) when the pattern has '
             'multiple tables. Without this flag, the first table is exported and '
             'extra tables produce a warning.',
    )
    p.add_argument(
        '--no-source',
        action='store_true',
        dest='no_source',
        help='Flatten table instance wrappers in JSON output: each table key maps '
             'directly to a list of row dicts instead of [{"_source": ..., "data": [...]}]. '
             'Only applies to --format nested.',
    )
    p.add_argument(
        '--meta',
        action='store_true',
        help='Add a _meta block to the JSON output (run_id, stats, issues) '
             'for pipeline auto-verification',
    )
    sheet_group = p.add_mutually_exclusive_group()
    sheet_group.add_argument(
        '--all-sheets',
        action='store_true',
        help='Process every sheet in the workbook; output is a dict keyed by sheet name',
    )
    sheet_group.add_argument(
        '--sheet',
        metavar='NAME_OR_INDEX',
        help='Sheet to use: name (e.g. Sheet2) or 0-based index (default: active sheet)',
    )
    p.add_argument(
        '--include-images',
        action='store_true',
        dest='include_images',
        help='Extract embedded images from the data file and save them to --images-dir. '
             'Adds an _images key to the JSON output mapping cell references to image paths.',
    )
    p.add_argument(
        '--images-dir',
        metavar='DIR',
        dest='images_dir',
        default=None,
        help='Directory to save extracted images into (default: next to --output, or cwd). '
             'Only used with --include-images.',
    )
    _add_strict_arg(p)
    _add_security_args(p)


def _add_watch_subparser(sub) -> None:
    p = sub.add_parser(
        'watch',
        help='Monitor a directory and extract new .xlsx files automatically',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Watch *directory* for new or moved-in .xlsx files and extract each one with the
given pattern as soon as it appears.  Results go to stdout (newline-delimited JSON)
or, with -o, to individual .json files in the output directory.

Requires:  pip install 'grepxcel[watch]'

examples:
  grepxcel watch -p pattern.xlsx inbox/
  grepxcel watch -p pattern.xlsx inbox/ -o output/
  grepxcel watch -p pattern.xlsx inbox/ --recursive -o output/
  grepxcel watch -p pattern.xlsx inbox/ --on-error stop
        """,
    )
    p.add_argument(
        '-p', '--pattern',
        required=True,
        metavar='FILE',
        help='Pattern file (.xlsx or .csv)',
    )
    p.add_argument(
        'directory',
        metavar='DIRECTORY',
        help='Directory to watch for new .xlsx files',
    )
    p.add_argument(
        '-o', '--output',
        metavar='DIR',
        help='Write extracted JSON files here instead of stdout',
    )
    p.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Also watch subdirectories',
    )
    p.add_argument(
        '--on-error',
        choices=['continue', 'stop'],
        default='continue',
        dest='on_error',
        help='What to do when a file fails to extract: '
             'continue (default) or stop the watcher',
    )
    p.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress progress messages; only print extracted JSON to stdout',
    )
    _add_sheet_arg(p)
    _add_security_args(p)


def _add_docs_subparser(sub) -> None:
    p = sub.add_parser(
        'docs',
        help='Write a pattern-format guide (pattern-reference.xlsx + grepxcel-guide.docx)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  grepxcel docs
  grepxcel docs -o reference/
  grepxcel docs -o reference/ -v   # show created directories
        """,
    )
    p.add_argument(
        '-o', '--output',
        metavar='DIR',
        default='.',
        help='Output directory (default: current directory)',
    )
    p.add_argument(
        '--force',
        action='store_true',
        help='Overwrite existing output files without prompting',
    )
    p.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Show created directories',
    )


def _add_draft_subparser(sub) -> None:
    _draft_args(sub.add_parser(
        'draft',
        help='Draft a starter pattern file for an Excel file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=r"""
'draft' is an OPTIONAL feature. 'extract' and 'docs' work out of the box; the
local 'draft' model needs an optional install:
    pip install 'grepxcel[suggest]'      (local model, runs offline)
    pip install 'grepxcel[draft-cloud]'  (cloud backends, e.g. --backend claude)
If a dependency is missing, grepxcel tells you exactly what to install.

backends:
  local   (default) Run a local GGUF model (Qwen3-8B) via llama-cpp-python.
          Needs:  pip install 'grepxcel[suggest]'
          Model is downloaded automatically on first run (~5 GB).
          No data leaves your machine during inference.
  claude  Send the Excel structure description to the Claude API.
          Needs ANTHROPIC_API_KEY and pip install 'grepxcel[draft-cloud]'.
          Uses claude-sonnet-5 by default — high quality at reasonable cost.
          The raw file is NOT transmitted — only column types, sample
          values, and labels are sent.
  nvidia  Send the Excel structure description to NVIDIA NIM (free-tier cloud).
          Needs NVIDIA_API_KEY (free key at build.nvidia.com) and pip install openai.
          Default model: mistralai/mistral-nemotron (override with --nvidia-model).
          Other examples: nvidia/nemotron-3-super-120b-a12b, z-ai/glm-5.3-flash
          The raw file is NOT transmitted — only column types, sample
          values, and labels are sent.
  server  Send the Excel structure description to any OpenAI-compatible API
          server (LM Studio, Ollama, vLLM, text-generation-inference, etc.).
          Needs pip install openai. Default URL: http://localhost:1234/v1
          (override with --server-url or GREPXCEL_SERVER_URL). Model is
          auto-discovered unless --server-model is set. Data stays local
          unless you point --server-url at a remote host.
  gemini  Planned for a future release — not yet available.
  github  Currently unavailable — GitHub retired the free-tier Models endpoint.
          The implementation is preserved for when GitHub provides a replacement.

Keys are read from a .env file (current dir or any parent) if present.
Cloud backends print per-call token usage; claude also prints the dollar cost.

model cache (local backend):
  Stored once in the per-user cache (platform-appropriate, via platformdirs):
    Linux    ~/.cache/grepxcel/models/   (honors $XDG_CACHE_HOME)
    macOS    ~/Library/Caches/grepxcel/models/
    Windows  %LOCALAPPDATA%\grepxcel\Cache\models\
  Override the location with GREPXCEL_MODEL_DIR (takes precedence).
  grepxcel only downloads when the file is missing — to avoid the HuggingFace
  download you can place the GGUF there yourself (exact filename), from any
  source. Faster/alternative downloads:
    HF_TOKEN=hf_...                 remove the anonymous rate limit (fastest fix)
    HF_ENDPOINT=https://hf-mirror.com   use a mirror
  Integrity: a downloaded model is checksummed (sha256) and re-verified on every
  run; a checksum mismatch (tampering/corruption) aborts. The check is cheap by
  default — it re-hashes only when the file's size/mtime changed. For high
  assurance, GREPXCEL_VERIFY_MODEL=full forces a full re-hash every run. A
  manually-placed file has no recorded checksum and can't be verified — grepxcel
  warns and proceeds on trust. Use --allow-unverified-model (or
  GREPXCEL_ALLOW_UNVERIFIED_MODEL=1) to silence the warning / override a mismatch.

examples:
  grepxcel draft report.xlsx
  grepxcel draft report.xlsx -o my_pattern.xlsx
  grepxcel draft report.xlsx --backend claude
  grepxcel draft report.xlsx --dry-run
        """,
    ))


def _draft_args(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        'file',
        metavar='FILE',
        help='Excel file to analyse',
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        default='draft_pattern.xlsx',
        help='Write draft pattern to FILE (default: draft_pattern.xlsx)',
    )
    p.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Print the Excel analysis sent to the LLM and model update status',
    )
    p.add_argument(
        '--dry-run',
        action='store_true',
        help='Print the Excel analysis that would be sent to the model, then exit without running inference',
    )
    p.add_argument(
        '--force',
        action='store_true',
        help='Overwrite an existing output file without prompting',
    )
    p.add_argument(
        '--backend',
        choices=['local', 'claude', 'gemini', 'github', 'nvidia', 'server'],
        default='local',
        help='Inference backend: local (default, GGUF model), claude (requires '
             'ANTHROPIC_API_KEY), nvidia (free-tier NVIDIA NIM, requires NVIDIA_API_KEY), '
             'server (any OpenAI-compatible server, e.g. LM Studio / Ollama / '
             'vLLM; use --server-url). '
             'gemini is planned for a future release. '
             'github is currently unavailable (GitHub retired the free-tier endpoint).',
    )
    p.add_argument(
        '--nvidia-model',
        default=os.environ.get('GREPXCEL_NVIDIA_MODEL', 'mistralai/mistral-nemotron'),
        help="NVIDIA NIM model id when --backend nvidia, e.g. "
             "'nvidia/nemotron-3-super-120b-a12b', 'z-ai/glm-5.3-flash' "
             "(default: mistralai/mistral-nemotron — benchmark winner). "
             "Available models vary by account tier — run `grepxcel doctor draft` to check. "
             "Also via GREPXCEL_NVIDIA_MODEL.",
    )
    p.add_argument(
        '--github-model',
        default='openai/gpt-4o-mini',
        help="GitHub Models model id when --backend github, e.g. 'openai/gpt-4o', "
             "'meta/llama-3.3-70b-instruct' (default: openai/gpt-4o-mini).",
    )
    p.add_argument(
        '--server-url',
        default=os.environ.get('GREPXCEL_SERVER_URL', 'http://localhost:1234/v1'),
        help='Base URL for --backend server (default: http://localhost:1234/v1). '
             'Also settable via GREPXCEL_SERVER_URL.',
    )
    p.add_argument(
        '--server-model',
        default=os.environ.get('GREPXCEL_SERVER_MODEL'),
        help='Model id for --backend server; if omitted, auto-discovers the '
             'first loaded model via /v1/models. Also via GREPXCEL_SERVER_MODEL.',
    )
    p.add_argument(
        '--server-api-key',
        default=os.environ.get('GREPXCEL_SERVER_API_KEY', 'not-needed'),
        help='API key for --backend server (most local servers ignore this). '
             'Also via GREPXCEL_SERVER_API_KEY.',
    )
    p.add_argument(
        '--allow-unverified-model',
        action='store_true',
        help='Proceed even if the cached local model fails its integrity check '
             '(checksum mismatch, or a manually-placed file with no recorded '
             'checksum). Also settable via GREPXCEL_ALLOW_UNVERIFIED_MODEL=1.',
    )
    p.add_argument(
        '--ca-bundle',
        metavar='FILE',
        default=None,
        help='Path to a corporate CA bundle (.pem) to trust behind a '
             'TLS-inspection proxy (NetSkope/Zscaler). Also via GREPXCEL_CA_BUNDLE '
             '/ REQUESTS_CA_BUNDLE / SSL_CERT_FILE. TLS verification stays on. '
             '(Experimental — not tested against a real intercept proxy.)',
    )
    p.add_argument(
        '--strict-env',
        action='store_true',
        help='Refuse a .env from outside the current project (also '
             'GREPXCEL_STRICT_ENV); the per-user config-dir .env stays allowed.',
    )
    _add_security_args(p)
    _add_sheet_arg(p)


# ── overwrite guard ──────────────────────────────────────────────────────────

def _refuse_overwrite(paths: list[str], force: bool) -> None:
    """Fail with a clear message if any path already exists and --force is not set."""
    if force:
        return
    existing = [p for p in paths if os.path.exists(p)]
    if not existing:
        return
    lines = ['Error: the following output file(s) already exist:']
    for p in existing:
        lines.append(f'  {p}')
    lines.append('Use --force to overwrite.')
    print('\n'.join(lines), file=sys.stderr)
    raise SystemExit(1)


# ── extract helpers ───────────────────────────────────────────────────────────

def _resolve_level(args) -> VerbosityLevel:
    if args.debug or args.verbose >= 2:
        return VerbosityLevel.DEBUG
    if args.verbose == 1:
        return VerbosityLevel.VERBOSE
    if getattr(args, 'quiet', False):
        return VerbosityLevel.QUIET
    return VerbosityLevel.NORMAL


def _process_file(pattern: str, data_file: str, args,
                  stem: str = None) -> tuple[bool, list[str]]:
    """
    Run the engine on one data file.
    Returns (ok, issue_fields) where ok is True if no errors/warnings occurred
    and issue_fields is a list of field names that had issues (for --strict).
    """
    level = _resolve_level(args)

    quiet = getattr(args, 'quiet', False)
    if len(args.files) > 1 and not quiet:
        print(f'\n{"─" * 62}', file=sys.stderr)
        print(f'  File: {data_file}', file=sys.stderr)
        print(f'{"─" * 62}', file=sys.stderr)

    logger = Logger(
        level=level,
        log_file=args.log,
        log_format=getattr(args, 'log_format', 'text'),
        source=data_file,
    )
    output_format = getattr(args, 'format', 'nested')
    if output_format == 'json':
        output_format = 'nested'
    is_csv = output_format == 'csv'
    is_xlsx = output_format == 'xlsx'
    if output_format == 'legacy':
        print(colorize_marks(
            f'  {MARK_WARN}  --format legacy is deprecated and will be removed in a future '
            'version. Switch to --format nested (the default). The legacy format exposes '
            'internal lbl: keys and _source/_anchor metadata.',
            should_color(sys.stderr)), file=sys.stderr)
    all_sheets = getattr(args, 'all_sheets', False)

    engine_format = 'nested' if (is_csv or is_xlsx) else output_format

    try:
        engine = Engine()
        if all_sheets:
            result = engine.process_all(
                pattern, data_file, logger=logger,
                max_file_mb=args.max_size,
                max_uncompressed_mb=args.max_uncompressed,
                max_cell_len=args.max_cell_len,
                max_rows=args.max_rows,
                max_cols=args.max_columns,
                output_format=engine_format,
            )
        else:
            result = engine.process(
                pattern, data_file, logger=logger,
                max_file_mb=args.max_size,
                max_uncompressed_mb=args.max_uncompressed,
                max_cell_len=args.max_cell_len,
                max_rows=args.max_rows,
                max_cols=args.max_columns,
                sheet=_resolve_sheet(args),
                output_format=engine_format,
            )
    except Exception as exc:
        print(colorize_marks(
            f'\n  {MARK_FAIL}  Unexpected error processing {data_file}: {exc}',
            should_color(sys.stderr)), file=sys.stderr)
        return False, []
    finally:
        logger.close()

    if getattr(args, 'meta', False):
        result['_meta'] = logger.build_meta()

    if getattr(args, 'include_images', False):
        from .engine import extract_images
        _images_dir = getattr(args, 'images_dir', None) or args.output or '.'
        _images, _img_warns = extract_images(
            data_file, _images_dir,
            stem or os.path.splitext(os.path.basename(data_file))[0],
        )
        for w in _img_warns:
            print(f'⚠️  [image] {w}', file=sys.stderr)
        if _images:
            result['_images'] = _images

    if getattr(args, 'no_source', False) and engine_format == 'nested':
        if all_sheets:
            result = {k: _flatten_table_instances(v) for k, v in result.items()}
        else:
            result = _flatten_table_instances(result)

    if is_xlsx:
        from .xlsx_writer import nested_to_xlsx
        os.makedirs(args.output, exist_ok=True)
        out_path = os.path.join(args.output, f'{stem}.xlsx')
        nested_to_xlsx(result, out_path, logger=logger)
        print(f'\n  Excel report written to: {out_path}', file=sys.stderr)
    elif is_csv:
        from .csv_writer import nested_to_csv
        _csv_table_arg = getattr(args, 'csv_table', None)
        _table_idx = (_csv_table_arg - 1) if _csv_table_arg is not None else None
        try:
            csv_text, csv_dropped = nested_to_csv(
                result,
                delimiter=getattr(args, 'csv_delimiter', ','),
                mode=getattr(args, 'csv_mode', 'extended'),
                table_idx=_table_idx,
            )
        except ValueError as _csv_err:
            print(colorize_marks(
                f'  {MARK_FAIL}  {_csv_err}',
                should_color(sys.stderr)), file=sys.stderr)
            sys.exit(2)
        if csv_dropped:
            dropped_list = ', '.join(csv_dropped)
            print(colorize_marks(
                f'  {MARK_WARN}  --format csv dropped header/footer rows: {dropped_list}. '
                'Use --format nested (JSON) to keep them.',
                should_color(sys.stderr)), file=sys.stderr)
        if args.output:
            os.makedirs(args.output, exist_ok=True)
            out_path = os.path.join(args.output, f'{stem}.csv')
            with open(out_path, 'w', encoding='utf-8', newline='') as f:
                f.write(csv_text)
            print(f'\n  CSV written to: {out_path}', file=sys.stderr)
        else:
            sys.stdout.write(csv_text)
    elif args.output:
        os.makedirs(args.output, exist_ok=True)
        out_path = os.path.join(args.output, f'{stem}.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, default=_json_default, ensure_ascii=False)
        print(f'\n  JSON written to: {out_path}', file=sys.stderr)
    else:
        json.dump(result, sys.stdout, indent=2, default=_json_default, ensure_ascii=False)
        sys.stdout.write('\n')

    ok = not (logger.has_errors() or logger.has_warnings())
    issue_fields: list[str] = []
    if not ok:
        for rec in logger.issues():
            if rec.field:
                issue_fields.append(rec.field)
    stats = logger.last_stats
    if stats:
        issue_fields.extend(stats.get('empty_field_names', []))

    # In quiet mode the Logger emitted nothing to the console; surface any
    # warnings/errors now as concise one-liners (same text as the ISSUES recap
    # inside the normal summary block, but without the stats header).
    if quiet and not ok:
        color = should_color(sys.stderr)
        if len(args.files) > 1:
            print(f'  {data_file}:', file=sys.stderr)
        for rec in logger.issues():
            print(colorize_marks('  ' + logger.issue_line(rec), color),
                  file=sys.stderr)

    return ok, issue_fields


_DEFAULT_MAX_FILES = 10_000


def _expand_files(paths: list[str], recursive: bool = False,
                  max_files: int = _DEFAULT_MAX_FILES) -> list[str]:
    """Expand directories in *paths* to their .xlsx files.

    Raises ``SystemExit`` if more than *max_files* are collected (safety cap
    against accidentally recursing into a huge tree).  Symlinks (files and
    directories) are skipped with a warning on stderr.
    """
    result = []
    symlinks_found = []

    def _check_cap():
        if len(result) > max_files:
            print(
                f'  Exceeded {max_files} files — aborting. '
                f'Use --max-files to raise the limit.',
                file=sys.stderr,
            )
            raise SystemExit(1)

    def _collect(dirpath, filenames):
        for fn in sorted(filenames):
            full = os.path.join(dirpath, fn)
            if os.path.islink(full):
                symlinks_found.append(full)
                continue
            if fn.lower().endswith('.xlsx') and not fn.startswith('~$'):
                result.append(full)
                _check_cap()

    for p in paths:
        if os.path.islink(p):
            symlinks_found.append(p)
            continue
        if os.path.isdir(p):
            if recursive:
                for dirpath, dirs, filenames in os.walk(p):
                    # warn about symlinked subdirectories
                    for d in dirs:
                        dp = os.path.join(dirpath, d)
                        if os.path.islink(dp):
                            symlinks_found.append(dp)
                    _collect(dirpath, filenames)
            else:
                _collect(p, os.listdir(p))
        else:
            result.append(p)

    if symlinks_found:
        print(colorize_marks(
            f'  {MARK_WARN}  Skipped {len(symlinks_found)} symlink(s) '
            f'(not followed for safety):',
            should_color(sys.stderr)), file=sys.stderr)
        for s in symlinks_found[:5]:
            print(f'       {s} → {os.readlink(s)}', file=sys.stderr)
        if len(symlinks_found) > 5:
            print(f'       … and {len(symlinks_found) - 5} more',
                  file=sys.stderr)

    return result


def _output_stem(data_file: str, all_files: list[str]) -> str:
    """
    Build a collision-safe output filename stem.
    If all input files have unique basenames, use the basename alone.
    If there are duplicates, prefix with the parent directory name.
    """
    basename = os.path.splitext(os.path.basename(data_file))[0]
    all_basenames = [os.path.splitext(os.path.basename(f))[0] for f in all_files]
    if all_basenames.count(basename) > 1:
        parent = os.path.basename(os.path.dirname(os.path.abspath(data_file)))
        return f'{parent}_{basename}'
    return basename


# ── docs handler ─────────────────────────────────────────────────────────────

def _run_docs(args) -> int:
    from .docs_generator import DocsGenerator
    verbose = getattr(args, 'verbose', False)
    xlsx_path = os.path.join(args.output, 'pattern-reference.xlsx')
    docx_path = os.path.join(args.output, 'grepxcel-guide.docx')
    _refuse_overwrite([xlsx_path, docx_path], getattr(args, 'force', False))
    dir_existed = os.path.isdir(args.output)
    paths = DocsGenerator().write(args.output)
    if verbose and not dir_existed:
        print(f'Created: {os.path.abspath(args.output)}', file=sys.stderr)
    for path in paths:
        print(f'Written: {path}', file=sys.stderr)
    return 0


# ── lint handler ─────────────────────────────────────────────────────────────

def _run_lint(args) -> int:
    from .lint import run_lint
    return run_lint(args.files,
                    recursive=getattr(args, 'recursive', False),
                    verbose=getattr(args, 'verbose', False))


# ── validate-pattern handler ─────────────────────────────────────────────────

def _run_validate(args) -> int:
    from .pattern_check import run_validate
    return run_validate(args.files,
                        verbose=getattr(args, 'verbose', False),
                        quiet=getattr(args, 'quiet', False))


# ── schema handler ──────────────────────────────────────────────────────────

def _run_skill(args) -> int:
    out_path = getattr(args, 'output', None)
    if out_path:
        _refuse_overwrite([out_path], getattr(args, 'force', False))
    from .skill import run_skill
    return run_skill(args.target, out_path)


def _run_examples(args) -> int:
    output_dir = args.output
    if os.path.isdir(output_dir) and os.listdir(output_dir):
        if not getattr(args, 'force', False):
            print(
                f'Error: the following output directory already exists and is not empty:\n'
                f'  {output_dir}\n'
                f'Use --force to overwrite.',
                file=sys.stderr,
            )
            return 1
    from .examples_generator import generate_examples
    generate_examples(output_dir)
    return 0


def _run_mcp(args) -> int:
    from .mcp_server import run_server
    verbose = getattr(args, 'verbose', 0)
    try:
        run_server(verbose=verbose)
    except KeyboardInterrupt:
        print('\nshutting down MCP server\ndone', file=sys.stderr)
    return 0


def _run_mcp_config(args) -> int:
    from .mcp_config import run_mcp_config
    return run_mcp_config(target=getattr(args, 'target', 'claude-code'))


def _run_sbom(args) -> int:
    out_path = getattr(args, 'output', None)
    if out_path:
        _refuse_overwrite([out_path], getattr(args, 'force', False))
    from .sbom import run_sbom
    return run_sbom(output=out_path)


def _run_schema(args) -> int:
    from .schema import run_schema
    if not args.output:
        return run_schema(args.files)
    _refuse_overwrite([args.output], getattr(args, 'force', False))
    out = open(args.output, 'w', encoding='utf-8')
    try:
        rc = run_schema(args.files, out=out)
    finally:
        out.close()
    if rc == 0:
        print(f'Schema written to: {args.output}', file=sys.stderr)
    return rc


# ── self-test handler ────────────────────────────────────────────────────────

def _run_self_test(args) -> int:
    from .self_test import run_self_test
    return run_self_test(verbose=getattr(args, 'verbose', False))


# ── doctor handler ───────────────────────────────────────────────────────────

def _run_doctor(args) -> int:
    from .doctor import run_doctor
    _load_dotenv(strict=getattr(args, 'strict_env', False))  # reflect .env/keys
    return run_doctor(area=getattr(args, 'area', 'all'),
                      probe=not getattr(args, 'no_probe', False),
                      strict_env=_strict_env_enabled(getattr(args, 'strict_env', False)))


# ── draft handler ─────────────────────────────────────────────────────────────

def _is_project_root(directory: str) -> bool:
    """A directory that marks a project boundary (.git or pyproject.toml)."""
    return (os.path.isdir(os.path.join(directory, '.git'))
            or os.path.isfile(os.path.join(directory, 'pyproject.toml')))


def _discover_project_env(start: str):
    """Walk cwd upward for a .env, stopping at the project root.

    Returns (path_or_None, in_project). in_project is False only when the .env
    was found in an ancestor that is NOT part of a recognised project (a loose /
    global .env) — the case --strict-env refuses.
    """
    directory = start
    while True:
        candidate = os.path.join(directory, '.env')
        has_marker = _is_project_root(directory)
        if os.path.isfile(candidate):
            return candidate, (directory == start or has_marker)
        if has_marker:
            return None, True   # project root, no .env — never escape above it
        parent = os.path.dirname(directory)
        if parent == directory:
            return None, False  # filesystem root, no project
        directory = parent


def _config_dir() -> str:
    """Per-user config dir for grepxcel (XDG-correct via platformdirs)."""
    try:
        from platformdirs import user_config_dir
        return user_config_dir('grepxcel')
    except Exception:
        return os.path.expanduser('~/.config/grepxcel')


def _config_dir_env() -> "str | None":
    """The sanctioned global .env at <config-dir>/.env, if it exists."""
    p = os.path.join(_config_dir(), '.env')
    return p if os.path.isfile(p) else None


def _read_env_file(path: str) -> None:
    """Load KEY=VALUE lines into the environment. Real env vars always win
    (setdefault); the file is only read (no execution); quotes are stripped."""
    try:
        with open(path, encoding='utf-8') as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, val = line.split('=', 1)
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key:
                    os.environ.setdefault(key, val)
    except OSError:
        pass


def _strict_env_enabled(explicit: bool = False) -> bool:
    """--strict-env flag OR GREPXCEL_STRICT_ENV truthy."""
    if explicit:
        return True
    return os.environ.get('GREPXCEL_STRICT_ENV', '').strip().lower() in (
        '1', 'true', 'yes', 'on', 'y')


def _load_dotenv(strict: bool = False) -> None:
    """Load .env credentials for cloud backends. Precedence (highest wins):
    real environment variables > project .env (cwd→project root) >
    ~/.config/grepxcel/.env (sanctioned global).

    A .env found in an unrelated ancestor is announced on stderr; with
    --strict-env / GREPXCEL_STRICT_ENV it is refused instead (the config-dir
    .env is the sanctioned exception and is always allowed).
    """
    strict = _strict_env_enabled(strict)
    start = os.getcwd()
    project_env, in_project = _discover_project_env(start)

    if project_env and not in_project:
        if strict:
            print(f'grepxcel: refusing out-of-project .env {project_env} '
                  f'(--strict-env / GREPXCEL_STRICT_ENV). Put credentials in a '
                  f'project .env or {os.path.join(_config_dir(), ".env")}.',
                  file=sys.stderr)
            sys.exit(2)
        print(f'grepxcel: loaded environment from {project_env} '
              f'(ancestor — not part of your project)', file=sys.stderr)
    elif project_env and (os.path.realpath(os.path.dirname(project_env))
                          != os.path.realpath(start)):
        print(f'grepxcel: loaded environment from {project_env}', file=sys.stderr)

    # Project .env first (its keys win via setdefault), then the config-dir .env.
    if project_env:
        _read_env_file(project_env)
    config_env = _config_dir_env()
    if config_env:
        _read_env_file(config_env)


def _run_draft(args) -> int:
    if not getattr(args, 'dry_run', False):
        _refuse_overwrite([args.output], getattr(args, 'force', False))
    _load_dotenv(strict=getattr(args, 'strict_env', False))  # find cloud keys
    # Wire corporate-proxy / custom-CA TLS trust before any network call
    # (model download or cloud backend). Verification stays on.
    from .proxy_support import enable_corporate_tls
    enable_corporate_tls(getattr(args, 'ca_bundle', None))
    from .drafter import ClaudeBackend, GitHubModelsBackend, NvidiaBackend, OpenAICompatBackend, PatternDrafter
    sheet    = _resolve_sheet(args)
    selected = getattr(args, 'backend', 'local')
    backend  = None
    if selected == 'gemini':
        # Implemented but disabled — planned for a future release.
        print(
            '[!] The Gemini backend is planned for a future release and is not yet '
            'available.\n    Use --backend local or --backend claude.',
            file=sys.stderr,
        )
        return 1
    if selected == 'github':
        # Implemented but disabled — GitHub retired the free-tier Models endpoint
        # (HTTP 410 "retirement brownout").  The implementation is preserved; flip
        # _GITHUB_MODELS_ENABLED in drafter.py if GitHub provides a new endpoint.
        print(
            '[!] The GitHub Models backend is currently unavailable.\n'
            '    GitHub retired the free-tier Models endpoint '
            '(HTTP 410 retirement brownout).\n'
            '    Use --backend local or --backend claude instead.',
            file=sys.stderr,
        )
        return 1
    if selected == 'claude':
        print("[!] Excel structure description will be sent to Anthropic's API.",
              file=sys.stderr)
        backend = ClaudeBackend()
    elif selected == 'nvidia':
        nv_model = getattr(args, 'nvidia_model', 'meta/llama-3.1-8b-instruct')
        print(f"[!] Excel structure description will be sent to NVIDIA NIM ({nv_model}).",
              file=sys.stderr)
        backend = NvidiaBackend(model=nv_model)
    elif selected == 'github':
        gh_model = getattr(args, 'github_model', 'openai/gpt-4o-mini')
        print(f"[!] Excel structure description will be sent to GitHub Models ({gh_model}).",
              file=sys.stderr)
        backend = GitHubModelsBackend(model=gh_model)
    elif selected == 'server':
        srv_url = getattr(args, 'server_url', 'http://localhost:1234/v1')
        srv_model = getattr(args, 'server_model', None)
        srv_key = getattr(args, 'server_api_key', 'not-needed')
        backend = OpenAICompatBackend(
            base_url=srv_url, model=srv_model, api_key=srv_key,
        )
    drafter = PatternDrafter(
        input_path=args.file,
        output_path=args.output,
        sheet=sheet,
        max_file_mb=args.max_size,
        max_uncompressed_mb=args.max_uncompressed,
        verbose=args.verbose,
        dry_run=args.dry_run,
        backend=backend,
        allow_unverified=getattr(args, 'allow_unverified_model', False),
    )
    return drafter.run()


# ── Multi-level help ──────────────────────────────────────────────────────────

def _synopsis_args(sp: argparse.ArgumentParser) -> str:
    """Compact args string: required flags, positionals, then [options]."""
    required_parts: list[str] = []
    positional_parts: list[str] = []
    has_optional = False
    for action in sp._actions:
        if isinstance(action, (argparse._HelpAction, argparse._SubParsersAction)):
            continue
        if not action.option_strings:
            meta = action.metavar or action.dest.upper()
            if isinstance(meta, tuple):
                meta = meta[0]
            if action.nargs == '+':
                positional_parts.append(f'{meta} [...]')
            elif action.nargs == '*':
                positional_parts.append(f'[{meta} ...]')
            elif action.nargs == '?':
                if action.choices:
                    positional_parts.append(f'[{{{",".join(str(c) for c in action.choices)}}}]')
                else:
                    positional_parts.append(f'[{meta}]')
            else:
                positional_parts.append(str(meta))
        elif getattr(action, 'required', False):
            flag = min(action.option_strings, key=len)
            meta = action.metavar or ''
            required_parts.append(f'{flag} {meta}'.strip())
        else:
            has_optional = True
    parts = required_parts + positional_parts
    if has_optional:
        parts.append('[options]')
    return '  '.join(parts)


def _print_synopsis(parser: argparse.ArgumentParser) -> None:
    """Print compact one-line synopsis per subcommand (-h -h)."""
    sub_action = next(
        (a for a in parser._actions if isinstance(a, argparse._SubParsersAction)),
        None,
    )
    if sub_action is None:
        return
    color = should_color(sys.stdout)
    heading = paint('Per-command synopsis', 'bold', color) + '  (-h -h -h for full help of each):'
    print(f'\n{heading}\n')
    name_w = max(len(n) for n in sub_action.choices) + 2
    for name, sp in sub_action.choices.items():
        args_str = _synopsis_args(sp)
        padding = ' ' * (name_w - len(name))
        print(f"  {paint(name, 'cyan', color)}{padding}  {args_str}")


def _print_full_help(parser: argparse.ArgumentParser) -> None:
    """Print full --help for every subcommand (-h -h -h)."""
    sub_action = next(
        (a for a in parser._actions if isinstance(a, argparse._SubParsersAction)),
        None,
    )
    if sub_action is None:
        return
    color = should_color(sys.stdout)
    first = True
    for name, sp in sub_action.choices.items():
        if not first:
            print()
        first = False
        label = paint(name, 'bold', color)
        fill = '─' * max(0, 58 - len(name))
        bar = '─' * 4
        print(f'{bar} {label} {fill}')
        print()
        print(sp.format_help())


# ── Entry point ───────────────────────────────────────────────────────────────

def main(argv=None):
    # 'suggest' is a transparent backward-compat alias for 'draft'
    if argv is None:
        argv = sys.argv[1:]
    if argv and argv[0] == 'suggest':
        argv = ['draft'] + argv[1:]

    help_count = sum(1 for a in argv if a in ('-h', '--help'))
    if help_count >= 2:
        p = _build_parser()
        if help_count >= 3:
            _print_full_help(p)
        else:
            _print_synopsis(p)
        sys.exit(0)

    parser = _build_parser()
    try:
        import argcomplete
        argcomplete.autocomplete(parser)
    except ImportError:
        pass
    args = parser.parse_args(argv)

    if args.command == 'quickstart':
        from .quickstart import run_quickstart
        sys.exit(run_quickstart())

    if args.command == 'watch':
        from .watcher import watch
        try:
            watch(
                pattern_path=args.pattern,
                directory=args.directory,
                output_dir=getattr(args, 'output', None),
                recursive=getattr(args, 'recursive', False),
                sheet=getattr(args, 'sheet', None),
                on_error=getattr(args, 'on_error', 'continue'),
                quiet=getattr(args, 'quiet', False),
                max_size_mb=getattr(args, 'max_size', 5.0),
                max_uncompressed_mb=getattr(args, 'max_uncompressed', 50.0),
            )
        except (FileNotFoundError, ValueError) as exc:
            print(f'grepxcel watch: error: {exc}', file=sys.stderr)
            sys.exit(1)
        except ImportError as exc:
            print(f'grepxcel watch: {exc}', file=sys.stderr)
            sys.exit(1)
        sys.exit(0)

    if args.command == 'web-wizard':
        from .wizard_api import run as run_web
        run_web(
            xlsx_path=args.file,
            pattern_path=getattr(args, 'pattern', None),
            port=getattr(args, 'port', 8765),
            open_browser=not getattr(args, 'no_browser', False),
            max_rows=getattr(args, 'max_rows', 150),
            max_cols=getattr(args, 'max_cols', 40),
            sheet=getattr(args, 'sheet', None),
            max_file_mb=getattr(args, 'max_size', 5),
            max_uncompressed_mb=getattr(args, 'max_uncompressed', 50),
        )
        sys.exit(0)

    if args.command == 'draft':
        sys.exit(_run_draft(args))

    if args.command == 'docs':
        sys.exit(_run_docs(args))

    if args.command == 'self-test':
        sys.exit(_run_self_test(args))

    if args.command == 'doctor':
        sys.exit(_run_doctor(args))

    if args.command == 'lint':
        sys.exit(_run_lint(args))

    if args.command == 'validate-pattern':
        sys.exit(_run_validate(args))

    if args.command == 'schema':
        sys.exit(_run_schema(args))

    if args.command == 'generate-skill':
        sys.exit(_run_skill(args))

    if args.command == 'generate-examples':
        sys.exit(_run_examples(args))

    if args.command == 'mcp':
        sys.exit(_run_mcp(args))

    if args.command == 'mcp-config':
        sys.exit(_run_mcp_config(args))

    if args.command == 'sbom':
        sys.exit(_run_sbom(args))

    if args.command == 'test':
        from .pattern_tester import run_tests, format_human, format_json
        try:
            report = run_tests(
                pattern_path=args.pattern,
                directory=args.directory,
                recursive=getattr(args, 'recursive', False),
                sheet=getattr(args, 'sheet', None),
                strict=getattr(args, 'strict', False),
                max_size_mb=getattr(args, 'max_size', 5.0),
                max_uncompressed_mb=getattr(args, 'max_uncompressed', 50.0),
            )
        except (FileNotFoundError, ValueError) as exc:
            print(f'grepxcel test: error: {exc}', file=sys.stderr)
            sys.exit(2)

        fmt = getattr(args, 'format', 'human')
        if fmt == 'json':
            print(format_json(report))
        else:
            no_color = getattr(args, 'no_color', False)
            use_color = (not no_color) and should_color(sys.stdout)
            verbose = getattr(args, 'verbose', 0)
            print(format_human(report, color=use_color, verbose=verbose))

        if report.failed > 0:
            sys.exit(2)
        elif report.warned > 0:
            sys.exit(1)
        else:
            sys.exit(0)

    # ── Security parameter validation ────────────────────────────────────────
    parser = _build_parser()
    if hasattr(args, 'max_size') and args.max_size <= 0:
        parser.error('--max-size must be greater than 0')
    if hasattr(args, 'max_uncompressed') and args.max_uncompressed <= 0:
        parser.error('--max-uncompressed must be greater than 0')
    if hasattr(args, 'max_cell_len') and args.max_cell_len < 1:
        parser.error('--max-cell-len must be at least 1')
    if hasattr(args, 'max_rows') and args.max_rows < 1:
        parser.error('--max-rows must be at least 1')
    if hasattr(args, 'max_columns') and args.max_columns < 1:
        parser.error('--max-columns must be at least 1')

    # extract — expand directories to .xlsx files
    expanded = _expand_files(args.files,
                             recursive=getattr(args, 'recursive', False),
                             max_files=getattr(args, 'max_files',
                                               _DEFAULT_MAX_FILES))
    if not expanded:
        print('  No .xlsx files found in the specified paths.', file=sys.stderr)
        sys.exit(1)

    fmt = getattr(args, 'format', 'nested')
    if fmt == 'json':
        fmt = 'nested'

    if fmt in ('csv', 'xlsx') and getattr(args, 'all_sheets', False):
        print(colorize_marks(
            f'\n  {MARK_FAIL}  --format {fmt} does not support --all-sheets '
            f'(a flat {fmt} cannot represent multiple sheets). '
            f'Use --sheet to pick one sheet, or --format nested for all sheets.',
            should_color(sys.stderr)), file=sys.stderr)
        sys.exit(2)

    if fmt == 'csv':
        from .csv_writer import count_table_instructions
        n_tables = count_table_instructions(args.pattern)
        csv_table = getattr(args, 'csv_table', None)
        if csv_table is not None and csv_table < 1:
            print(colorize_marks(
                f'\n  {MARK_FAIL}  --csv-table must be >= 1 (got {csv_table}).',
                should_color(sys.stderr)), file=sys.stderr)
            sys.exit(2)
        if n_tables > 1:
            if csv_table is not None and csv_table > n_tables:
                print(colorize_marks(
                    f'\n  {MARK_FAIL}  --csv-table {csv_table} is out of range: '
                    f'this pattern has {n_tables} tables.',
                    should_color(sys.stderr)), file=sys.stderr)
                sys.exit(2)
            if csv_table is None:
                print(colorize_marks(
                    f'  {MARK_WARN}  Pattern has {n_tables} tables; --format csv exports '
                    f'only the first. Use --csv-table N (1-{n_tables}) to choose a specific '
                    'table, or --format nested for all tables.',
                    should_color(sys.stderr)), file=sys.stderr)

    if fmt == 'xlsx':
        if not args.output:
            print(colorize_marks(
                f'\n  {MARK_FAIL}  --format xlsx requires -o / --output (cannot write '
                'binary Excel to stdout).',
                should_color(sys.stderr)), file=sys.stderr)
            sys.exit(2)
        out_dir = os.path.abspath(args.output)
        for data_file in expanded:
            out_path = os.path.join(
                out_dir,
                os.path.splitext(os.path.basename(data_file))[0] + '.xlsx',
            )
            if os.path.abspath(data_file) == os.path.abspath(out_path):
                print(colorize_marks(
                    f'\n  {MARK_FAIL}  --format xlsx would overwrite the source file '
                    f'{data_file}. Use a different -o directory.',
                    should_color(sys.stderr)), file=sys.stderr)
                sys.exit(2)

    # ── Pattern pre-validation ───────────────────────────────────────────────
    # Run static checks once before touching any data file.
    # Errors abort immediately; warnings are printed but extraction continues.
    from .pattern_check import check_pattern
    _pv = check_pattern(args.pattern)
    if not _pv.valid:
        color = should_color(sys.stderr)
        print(colorize_marks(
            f'{MARK_FAIL}  Pattern invalid: {args.pattern}', color), file=sys.stderr)
        for err in _pv.errors:
            print(colorize_marks(f'   {MARK_FAIL} {err}', color), file=sys.stderr)
        sys.exit(1)
    if _pv.warnings:
        color = should_color(sys.stderr)
        quiet = getattr(args, 'quiet', False)
        if not quiet:
            print(colorize_marks(
                f'{MARK_WARN}  Pattern warnings: {args.pattern}', color), file=sys.stderr)
        for warn in _pv.warnings:
            print(colorize_marks(f'   {MARK_WARN} {warn}', color), file=sys.stderr)

    all_ok = True
    strict = getattr(args, 'strict', False)
    strict_failures: list[tuple[str, list[str]]] = []

    for data_file in expanded:
        ok, empty_fields = _process_file(args.pattern, data_file, args,
                                         stem=_output_stem(data_file, expanded))
        all_ok = all_ok and ok
        if strict and empty_fields:
            strict_failures.append((data_file, empty_fields))

    if strict_failures:
        color = should_color(sys.stderr)
        print(colorize_marks(
            f'\n  {MARK_FAIL}  --strict: missing fields detected',
            color), file=sys.stderr)
        for data_file, fields in strict_failures:
            if len(expanded) > 1:
                print(f'  {data_file}:', file=sys.stderr)
            for field in fields:
                print(f'    • {field}', file=sys.stderr)
        sys.exit(2)

    sys.exit(0 if all_ok else 1)
