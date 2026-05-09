"""
Command-line interface for grepxcel.

Usage:
    grepxcel -p pattern.xlsx data.xlsx
    grepxcel -p pattern.xlsx data1.xlsx data2.xlsx data3.xlsx
    grepxcel -p pattern.xlsx data.xlsx -v 2 --output results/
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
verbosity levels:
  0  quiet   — no console output during processing
  1  normal  — summary + all validation warnings  (default)
  2  verbose — + step-by-step: cells found, tables matched
  3  debug   — + every anchor probe and rejection reason

examples:
  grepxcel -p pattern.xlsx report.xlsx
  grepxcel -p pattern.xlsx jan.xlsx feb.xlsx mar.xlsx
  grepxcel -p pattern.xlsx data.xlsx -v 2
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
        '-v', '--verbosity',
        type=int, choices=[0, 1, 2, 3], default=1, metavar='LEVEL',
        help='Console verbosity: 0=quiet 1=normal 2=verbose 3=debug  (default: 1)',
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
        type=float, default=50, metavar='MB',
        help='Maximum file size in MB to accept (default: 50). '
             'Raise this for legitimate large files.',
    )

    return p


# ── Per-file processing ───────────────────────────────────────────────────────

def _process_file(pattern: str, data_file: str, args, stem: str = None) -> bool:
    """
    Run the engine on one data file.
    Returns True if the file had no errors, False otherwise.
    """
    level = VerbosityLevel(args.verbosity)

    if args.verbosity > 0 and len(args.files) > 1:
        print(f'\n{"─" * 62}')
        print(f'  File: {data_file}')
        print(f'{"─" * 62}')

    logger = Logger(level=level, log_file=args.log)

    try:
        result = Engine().process(
            pattern, data_file, logger=logger,
            max_file_mb=args.max_size,
        )
    except Exception as exc:
        print(f'\n  ✗  Unexpected error processing {data_file}: {exc}', file=sys.stderr)
        return False
    finally:
        logger.close()

    if args.output:
        os.makedirs(args.output, exist_ok=True)
        out_path = os.path.join(args.output, f'{stem}.json')
        payload = {
            'cells':  result.get('cells', {}),
            'tables': result.get('tables', []),
        }
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, indent=2, default=_json_default)
        if args.verbosity > 0:
            print(f'\n  JSON written to: {out_path}')

    return not logger.has_errors()


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
