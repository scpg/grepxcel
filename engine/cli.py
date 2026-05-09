"""
Command-line interface for grepxcel.

Usage:
    python -m engine pattern.xlsx data.xlsx
    python -m engine pattern.xlsx data.xlsx -v 2
    python -m engine pattern.xlsx data.xlsx --output results.json
    python -m engine pattern.xlsx data.xlsx --log logs/run.log
"""

import argparse
import datetime
import json
import sys

from .engine import Engine
from .logger import Logger, VerbosityLevel


def _json_default(obj):
    """Make datetime objects JSON-serialisable."""
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()
    raise TypeError(f'Object of type {type(obj).__name__} is not JSON serialisable')


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='python -m engine',
        description='Extract structured data from Excel using a pattern file.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
verbosity levels:
  0  quiet   — no console output during processing
  1  normal  — summary + all validation warnings (default)
  2  verbose — + step-by-step: cells found, tables matched
  3  debug   — + every anchor probe and rejection reason

examples:
  python -m engine pattern.xlsx data.xlsx
  python -m engine pattern.xlsx data.xlsx -v 2
  python -m engine pattern.xlsx data.xlsx --output results.json
  python -m engine pattern.xlsx data.xlsx --log logs/run.log --output out.json
        """,
    )
    p.add_argument('pattern', help='Pattern Excel file (.xlsx)')
    p.add_argument('data',    help='Data Excel file (.xlsx)')
    p.add_argument(
        '-v', '--verbosity',
        type=int, choices=[0, 1, 2, 3], default=1, metavar='LEVEL',
        help='Console verbosity: 0=quiet, 1=normal, 2=verbose, 3=debug (default: 1)',
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        help='Write extracted data as JSON to FILE',
    )
    p.add_argument(
        '-l', '--log',
        metavar='FILE',
        help='Write structured log to FILE',
    )
    return p


def main(argv=None):
    args = _build_parser().parse_args(argv)

    level = VerbosityLevel(args.verbosity)
    logger = Logger(level=level, log_file=args.log)

    try:
        result = Engine().process(args.pattern, args.data, logger=logger)
    except SystemExit:
        raise
    except Exception as exc:
        print(f'\n  ✗  Unexpected error: {exc}', file=sys.stderr)
        sys.exit(1)
    finally:
        logger.close()

    if args.output:
        payload = {
            'cells':  result.get('cells', {}),
            'tables': result.get('tables', []),
        }
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(payload, f, indent=2, default=_json_default)
        print(f'\n  JSON written to: {args.output}')

    sys.exit(1 if logger.has_errors() else 0)
