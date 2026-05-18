"""
Command-line interface for grepxcel.

Usage:
    grepxcel extract -p pattern.xlsx data.xlsx
    grepxcel extract -p pattern.xlsx data1.xlsx data2.xlsx data3.xlsx
    grepxcel extract -p pattern.xlsx data.xlsx -v --output results/
    grepxcel extract -p pattern.xlsx data.xlsx --log logs/run.log
    grepxcel suggest data.xlsx -o suggested_pattern.xlsx
    grepxcel suggest data.xlsx -v
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


def _resolve_sheet(args) -> str | int | None:
    sheet = getattr(args, 'sheet', None)
    if sheet is not None:
        try:
            return int(sheet)
        except ValueError:
            pass
    return sheet


# ── Argument parser ───────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='grepxcel',
        description='Extract structured data from Excel files using a pattern.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
commands:
  extract   Extract data from Excel files using a pattern file
  suggest   Use a local LLM to suggest a pattern file for an Excel file
  docs      Write a pattern-format reference xlsx (pattern-reference.xlsx)

Run 'grepxcel <command> --help' for per-command options.
        """,
    )
    sub = p.add_subparsers(dest='command', metavar='COMMAND')
    sub.required = True

    _add_extract_subparser(sub)
    _add_suggest_subparser(sub)
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
    _add_security_args(p)
    _add_sheet_arg(p)


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


def _add_suggest_subparser(sub) -> None:
    p = sub.add_parser(
        'suggest',
        help='Use a local LLM to suggest a pattern file for an Excel file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
The model is downloaded automatically on first run (~2.4 GB) and cached locally.
A lightweight update check runs once per day — no data ever leaves your machine
during inference.

Override the model cache directory with the GREPXCEL_MODEL_DIR environment variable.

examples:
  grepxcel suggest report.xlsx
  grepxcel suggest report.xlsx -o my_pattern.xlsx
  grepxcel suggest report.xlsx -v
        """,
    )
    p.add_argument(
        'file',
        metavar='FILE',
        help='Excel file to analyse',
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        default='suggested_pattern.xlsx',
        help='Write suggested pattern to FILE (default: suggested_pattern.xlsx)',
    )
    p.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Print the Excel analysis sent to the LLM and model update status',
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
    sheet = _resolve_sheet(args)
    output_format = getattr(args, 'format', 'nested')

    try:
        result = Engine().process(
            pattern, data_file, logger=logger,
            max_file_mb=args.max_size,
            max_uncompressed_mb=args.max_uncompressed,
            max_cell_len=args.max_cell_len,
            sheet=sheet,
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


# ── suggest handler ───────────────────────────────────────────────────────────

def _run_suggest(args) -> int:
    from .suggester import PatternSuggester
    sheet = _resolve_sheet(args)
    suggester = PatternSuggester(
        input_path=args.file,
        output_path=args.output,
        sheet=sheet,
        max_file_mb=args.max_size,
        max_uncompressed_mb=args.max_uncompressed,
        verbose=args.verbose,
    )
    return suggester.run()


# ── Entry point ───────────────────────────────────────────────────────────────

def main(argv=None):
    args = _build_parser().parse_args(argv)

    if args.command == 'suggest':
        sys.exit(_run_suggest(args))

    if args.command == 'docs':
        sys.exit(_run_docs(args))

    # extract
    all_ok = True
    for data_file in args.files:
        ok = _process_file(args.pattern, data_file, args,
                           stem=_output_stem(data_file, args.files))
        all_ok = all_ok and ok

    sys.exit(0 if all_ok else 1)
