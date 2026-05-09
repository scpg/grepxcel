"""
Command-line interface for grepxcel.

Usage:
    grepxcel -p pattern.xlsx data.xlsx
    grepxcel -p pattern.xlsx data1.xlsx data2.xlsx data3.xlsx
    grepxcel -p pattern.xlsx data.xlsx -v --output results/
    grepxcel -p pattern.xlsx data.xlsx --log logs/run.log
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


# ── Argument parser ───────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='grepxcel',
        description='Extract structured data from Excel files using a pattern.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
verbosity:
  (none)  warnings + summary on stderr, extracted JSON on stdout
  -v      + step-by-step: cells found, tables matched
  -vv     + every anchor probe and rejection reason
  -d      same as -vv (debug mode)

examples:
  grepxcel -p pattern.xlsx report.xlsx
  grepxcel -p pattern.xlsx jan.xlsx feb.xlsx mar.xlsx
  grepxcel -p pattern.xlsx data.xlsx -v
  grepxcel -p pattern.xlsx data.xlsx --output results/
  grepxcel -p pattern.xlsx data.xlsx --log logs/run.log --output results/
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
        '--max-size',
        type=float, default=5, metavar='MB',
        help='Maximum compressed file size in MB (default: 5)',
    )
    p.add_argument(
        '--max-uncompressed',
        type=float, default=50, metavar='MB',
        help='Maximum uncompressed ZIP content in MB (default: 50)',
    )
    p.add_argument(
        '--max-cell-len',
        type=int, default=1000, metavar='CHARS',
        help='Maximum cell character length passed to regex matching (default: 1000)',
    )
    p.add_argument(
        '--sheet',
        metavar='NAME_OR_INDEX',
        help='Sheet to process: name (e.g. Sheet2) or 0-based index (default: active sheet)',
    )

    return p


# ── Per-file processing ───────────────────────────────────────────────────────

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

    sheet = args.sheet
    if sheet is not None:
        try:
            sheet = int(sheet)
        except ValueError:
            pass  # keep as string (sheet name)

    try:
        result = Engine().process(
            pattern, data_file, logger=logger,
            max_file_mb=args.max_size,
            max_uncompressed_mb=args.max_uncompressed,
            max_cell_len=args.max_cell_len,
            sheet=sheet,
        )
    except Exception as exc:
        print(f'\n  ✗  Unexpected error processing {data_file}: {exc}', file=sys.stderr)
        return False
    finally:
        logger.close()

    payload = {
        'cells':  result.get('cells', {}),
        'tables': result.get('tables', []),
    }

    if args.output:
        os.makedirs(args.output, exist_ok=True)
        out_path = os.path.join(args.output, f'{stem}.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, indent=2, default=_json_default)
        print(f'\n  JSON written to: {out_path}', file=sys.stderr)
    else:
        json.dump(payload, sys.stdout, indent=2, default=_json_default)
        sys.stdout.write('\n')

    return not (logger.has_errors() or logger.has_warnings())


# ── Entry point ───────────────────────────────────────────────────────────────

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


def main(argv=None):
    args = _build_parser().parse_args(argv)

    all_ok = True
    for data_file in args.files:
        ok = _process_file(args.pattern, data_file, args,
                           stem=_output_stem(data_file, args.files))
        all_ok = all_ok and ok

    sys.exit(0 if all_ok else 1)
