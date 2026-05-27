#!/usr/bin/env python3
"""
Run the engine against every test fixture and write the extracted JSON to
output/fixtures/<fixture_name>[.<sheet>].json.

Multi-sheet fixtures are extracted once per sheet and the sheet name is
appended to the filename.

Usage:
    python scripts/extract_fixtures.py                 # all fixtures
    python scripts/extract_fixtures.py 06 03           # selected fixtures (by prefix)

Output directory: output/fixtures/  (gitignored)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

import contextlib
import datetime
import io
import json

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

import openpyxl

from engine import Engine, Logger, VerbosityLevel

ROOT     = os.path.join(os.path.dirname(__file__), '..')
FIXTURES = os.path.join(ROOT, 'tests', 'fixtures')
OUT_DIR  = os.path.join(ROOT, 'output', 'fixtures')


def _json_default(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()
    raise TypeError(f'Type {type(obj).__name__} is not JSON serialisable')


def _sheet_names(data_path: str) -> list[str]:
    """Return all sheet names in the workbook."""
    wb = openpyxl.load_workbook(data_path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def _extract(pattern_path: str, data_path: str, sheet=None) -> tuple[dict, bool]:
    """Returns (result, had_errors). Suppresses engine stderr output."""
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    with contextlib.redirect_stderr(io.StringIO()):
        result = Engine().process(pattern_path, data_path, logger=lg, **kwargs)
    return result, lg.has_errors()


def _write(path: str, data: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, default=_json_default)


def run(filter_prefixes: list[str] | None = None) -> None:
    os.makedirs(OUT_DIR, exist_ok=True)

    fixture_dirs = sorted(
        d for d in os.listdir(FIXTURES)
        if os.path.isdir(os.path.join(FIXTURES, d))
    )

    if filter_prefixes:
        fixture_dirs = [
            d for d in fixture_dirs
            if any(d.startswith(pf) for pf in filter_prefixes)
        ]

    if not fixture_dirs:
        print('No matching fixtures found.')
        return

    for name in fixture_dirs:
        folder       = os.path.join(FIXTURES, name)
        pattern_path = os.path.join(folder, 'pattern.xlsx')
        data_path    = os.path.join(folder, 'data.xlsx')

        if not os.path.exists(pattern_path) or not os.path.exists(data_path):
            print(f'  {name}/  SKIP (missing pattern.xlsx or data.xlsx)')
            continue

        sheets = _sheet_names(data_path)

        if len(sheets) == 1:
            # Single sheet — write <fixture_name>.json
            try:
                result, had_errors = _extract(pattern_path, data_path)
                out_path = os.path.join(OUT_DIR, f'{name}.json')
                _write(out_path, result)
                suffix = '  (partial — engine errors)' if had_errors else ''
                print(f'  {name}.json{suffix}')
            except Exception as exc:
                print(f'  {name}  ERROR: {exc}')
        else:
            # Multi-sheet — write <fixture_name>.<sheet>.json for each sheet
            for sheet in sheets:
                try:
                    result, had_errors = _extract(pattern_path, data_path, sheet=sheet)
                    out_path = os.path.join(OUT_DIR, f'{name}.{sheet}.json')
                    _write(out_path, result)
                    suffix = '  (partial — engine errors)' if had_errors else ''
                    print(f'  {name}.{sheet}.json{suffix}')
                except Exception as exc:
                    print(f'  {name}.{sheet}  ERROR: {exc}')


if __name__ == '__main__':
    prefixes = sys.argv[1:] if len(sys.argv) > 1 else None
    print(f'Writing JSON to {os.path.relpath(OUT_DIR)}/')
    run(prefixes)
    print('Done.')
