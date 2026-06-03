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

from .engine import Engine
from .logger import Logger, VerbosityLevel


# ── JSON serialisation ────────────────────────────────────────────────────────

def _json_default(obj):
    """Serialise types that json.dump does not handle natively."""
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()
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

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='grepxcel',
        description='Extract structured data from Excel files using a pattern.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
commands:
  extract   Extract data from Excel files using a pattern file
  draft     Use a local LLM to draft a starter pattern file for an Excel file
  docs      Write a pattern-format reference xlsx (pattern-reference.xlsx)

Run 'grepxcel <command> --help' for per-command options.
        """,
    )
    from . import __version__
    p.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__}',
    )
    sub = p.add_subparsers(dest='command', metavar='COMMAND')
    sub.required = True

    _add_extract_subparser(sub)
    _add_draft_subparser(sub)
    _add_docs_subparser(sub)
    return p


def _add_extract_subparser(sub) -> None:
    p = sub.add_parser(
        'extract',
        help='Extract data from Excel files using a pattern file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
verbosity:
  (none)  warnings + summary on stderr, extracted JSON on stdout
  -v      + step-by-step: cells found, tables matched
  -vv     + every anchor probe and rejection reason
  -d      same as -vv (debug mode)

examples:
  grepxcel extract -p pattern.xlsx report.xlsx
  grepxcel extract -p pattern.xlsx jan.xlsx feb.xlsx mar.xlsx
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
        metavar='FILE',
        help='One or more data Excel files to process',
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
        '-o', '--output',
        metavar='DIR',
        help='Write extracted data as JSON into DIR  (<datafile-stem>.json per file)',
    )
    p.add_argument(
        '-l', '--log',
        metavar='FILE',
        help='Append structured log to FILE',
    )
    p.add_argument(
        '--max-cell-len',
        type=int, default=1000, metavar='CHARS',
        help='Maximum cell character length passed to regex matching (default: 1000)',
    )
    p.add_argument(
        '--format',
        choices=['nested', 'legacy'], default='nested',
        help='Output format: nested (default) or legacy ({"cells":{}, "tables":[]})',
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
    _add_security_args(p)


def _add_docs_subparser(sub) -> None:
    p = sub.add_parser(
        'docs',
        help='Write a self-documenting pattern-format reference xlsx',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  grepxcel docs
  grepxcel docs -o reference/pattern-reference.xlsx
        """,
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        default='pattern-reference.xlsx',
        help='Output path for the reference file (default: pattern-reference.xlsx)',
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
  local   (default) Run a local GGUF model via llama-cpp-python.
          Needs:  pip install 'grepxcel[suggest]'
          Model is downloaded automatically on first run (~4.7 GB).
          No data leaves your machine during inference.
  claude  Send the Excel structure description to the Claude API.
          Needs ANTHROPIC_API_KEY and pip install 'grepxcel[draft-cloud]'.
          The raw file is NOT transmitted — only column types, sample
          values, and labels are sent.
  gemini  Planned for a future release — not yet available.

The Claude backend prints a one-line privacy notice and the per-call token cost.

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
  Note: a manually-placed file is NOT hash-verified — trust your source.

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
        '--backend',
        choices=['local', 'claude', 'gemini'],
        default='local',
        help='Inference backend: local (default, GGUF model) or claude (requires '
             'ANTHROPIC_API_KEY). gemini is planned for a future release.',
    )
    _add_security_args(p)
    _add_sheet_arg(p)


# ── extract helpers ───────────────────────────────────────────────────────────

def _resolve_level(args) -> VerbosityLevel:
    if args.debug or args.verbose >= 2:
        return VerbosityLevel.DEBUG
    if args.verbose == 1:
        return VerbosityLevel.VERBOSE
    return VerbosityLevel.NORMAL


def _process_file(pattern: str, data_file: str, args, stem: str = None) -> bool:
    """
    Run the engine on one data file.
    Returns True if the file had no errors or warnings, False otherwise.
    """
    level = _resolve_level(args)

    if len(args.files) > 1:
        print(f'\n{"─" * 62}', file=sys.stderr)
        print(f'  File: {data_file}', file=sys.stderr)
        print(f'{"─" * 62}', file=sys.stderr)

    logger = Logger(level=level, log_file=args.log)
    output_format = getattr(args, 'format', 'nested')
    all_sheets = getattr(args, 'all_sheets', False)

    try:
        engine = Engine()
        if all_sheets:
            result = engine.process_all(
                pattern, data_file, logger=logger,
                max_file_mb=args.max_size,
                max_uncompressed_mb=args.max_uncompressed,
                max_cell_len=args.max_cell_len,
                output_format=output_format,
            )
        else:
            result = engine.process(
                pattern, data_file, logger=logger,
                max_file_mb=args.max_size,
                max_uncompressed_mb=args.max_uncompressed,
                max_cell_len=args.max_cell_len,
                sheet=_resolve_sheet(args),
                output_format=output_format,
            )
    except Exception as exc:
        print(f'\n  ✗  Unexpected error processing {data_file}: {exc}', file=sys.stderr)
        return False
    finally:
        logger.close()

    if args.output:
        os.makedirs(args.output, exist_ok=True)
        out_path = os.path.join(args.output, f'{stem}.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, default=_json_default)
        print(f'\n  JSON written to: {out_path}', file=sys.stderr)
    else:
        json.dump(result, sys.stdout, indent=2, default=_json_default)
        sys.stdout.write('\n')

    return not (logger.has_errors() or logger.has_warnings())


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
    DocsGenerator().write(args.output)
    print(f'Pattern reference written to: {args.output}', file=sys.stderr)
    return 0


# ── draft handler ─────────────────────────────────────────────────────────────

def _run_draft(args) -> int:
    from .drafter import ClaudeBackend, PatternDrafter
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
    if selected == 'claude':
        print("[!] Excel structure description will be sent to Anthropic's API.",
              file=sys.stderr)
        backend = ClaudeBackend()
    drafter = PatternDrafter(
        input_path=args.file,
        output_path=args.output,
        sheet=sheet,
        max_file_mb=args.max_size,
        max_uncompressed_mb=args.max_uncompressed,
        verbose=args.verbose,
        dry_run=args.dry_run,
        backend=backend,
    )
    return drafter.run()


# ── Entry point ───────────────────────────────────────────────────────────────

def main(argv=None):
    # 'suggest' is a transparent backward-compat alias for 'draft'
    if argv is None:
        argv = sys.argv[1:]
    if argv and argv[0] == 'suggest':
        argv = ['draft'] + argv[1:]

    args = _build_parser().parse_args(argv)

    if args.command == 'draft':
        sys.exit(_run_draft(args))

    if args.command == 'docs':
        sys.exit(_run_docs(args))

    # extract
    all_ok = True
    for data_file in args.files:
        ok = _process_file(args.pattern, data_file, args,
                           stem=_output_stem(data_file, args.files))
        all_ok = all_ok and ok

    sys.exit(0 if all_ok else 1)
