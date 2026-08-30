#!/usr/bin/env python3
"""
Regenerate LLM/AI draft patterns for one or more fixtures.

Runs 'grepxcel draft --backend <backend>' against each fixture's data file and
writes the result as a versioned pattern file:

    tests/fixtures/{name}/{name}_pattern-from-{backend}.xlsx
    tests/fixtures/{name}/{name}_pattern-from-{backend}.csv   (converted from xlsx)

This script is intentionally NOT run in CI — LLM calls are slow and credit-
intensive.  Run it manually when you want to refresh a draft pattern.

Usage
-----
    # One fixture, one backend
    python3 scripts/regenerate_drafts.py 01_simple_invoice --backend claude
    python3 scripts/regenerate_drafts.py 01_simple_invoice --backend local

    # Several fixtures
    python3 scripts/regenerate_drafts.py 01_simple_invoice 02_product_catalog --backend claude

    # All fixtures (careful — this calls the LLM 22 times)
    python3 scripts/regenerate_drafts.py --all --backend claude

Backend names
-------------
    local   — local LLM (llama-cpp, requires 'grepxcel[suggest]' extras)
    claude  — Claude API  (requires ANTHROPIC_API_KEY)
    github  — GitHub Models / OpenAI-compatible (requires GITHUB_TOKEN)
    (gemini is built but disabled — not yet available here)

After running this script, commit the new pattern file(s) and regenerate the
matching snapshot with scripts/generate_snapshots.py.
"""

import os
import sys
import csv
import argparse
import subprocess
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

import openpyxl

ROOT     = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
FIXTURES = os.path.join(ROOT, 'tests', 'fixtures')

# Map from our file-suffix backend names to grepxcel CLI --backend values.
# gemini is intentionally excluded — it is built but disabled pending GCP billing
# (see memory: project_gemini_pending.md).  Add it here once enabled.
_BACKEND_MAP = {
    'claude': 'claude',
    'local':  'local',
    'github': 'github',
}


def _xlsx_to_csv(xlsx_path: str, csv_path: str) -> None:
    """Convert a pattern xlsx workbook to its CSV twin.

    Strips trailing None columns from each row so the CSV stays compact.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    with open(csv_path, 'w', newline='', encoding='utf-8') as fh:
        writer = csv.writer(fh)
        for row in ws.iter_rows(values_only=True):
            r = list(row)
            while r and r[-1] is None:
                r.pop()
            writer.writerow(r)
    wb.close()


def regenerate(fixture_name: str, backend: str, dry_run: bool = False) -> bool:
    """Draft a new pattern for *fixture_name* using *backend*.

    Returns True on success, False on any error.
    """
    folder = os.path.join(FIXTURES, fixture_name)
    if not os.path.isdir(folder):
        print(f'  [{fixture_name}] ERROR: fixture folder not found', file=sys.stderr)
        return False

    data_file = os.path.join(folder, f'{fixture_name}_data.xlsx')
    if not os.path.exists(data_file):
        print(f'  [{fixture_name}] ERROR: data file not found: {data_file}', file=sys.stderr)
        return False

    out_xlsx = os.path.join(folder, f'{fixture_name}_pattern-from-{backend}.xlsx')
    out_csv  = os.path.join(folder, f'{fixture_name}_pattern-from-{backend}.csv')
    cli_backend = _BACKEND_MAP[backend]

    print(f'  [{fixture_name}] drafting with --backend {cli_backend} …', flush=True)

    if dry_run:
        print(f'  [{fixture_name}] DRY-RUN: would write {os.path.basename(out_xlsx)}')
        return True

    # Write to a temp file first so the fixture folder is only touched on success.
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name

    grepxcel_bin = os.path.join(os.path.dirname(sys.executable), 'grepxcel')
    cmd = [grepxcel_bin, 'draft', data_file, '--backend', cli_backend, '-o', tmp_path]

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0 or not os.path.exists(tmp_path) or os.path.getsize(tmp_path) == 0:
            print(f'  [{fixture_name}] ERROR: grepxcel draft failed', file=sys.stderr)
            if proc.stderr:
                print(proc.stderr.rstrip(), file=sys.stderr)
            return False

        shutil.move(tmp_path, out_xlsx)
        _xlsx_to_csv(out_xlsx, out_csv)
        print(f'  [{fixture_name}] ✓ {os.path.basename(out_xlsx)}')
        print(f'  [{fixture_name}] ✓ {os.path.basename(out_csv)}')
        return True

    except Exception as exc:
        print(f'  [{fixture_name}] ERROR: {exc}', file=sys.stderr)
        return False
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        'fixtures', nargs='*',
        help='Fixture name(s), e.g. 01_simple_invoice 02_product_catalog',
    )
    parser.add_argument(
        '--backend', required=True, choices=list(_BACKEND_MAP),
        help='LLM/AI backend to use for draft generation',
    )
    parser.add_argument(
        '--all', dest='all_fixtures', action='store_true',
        help='Regenerate drafts for every fixture',
    )
    parser.add_argument(
        '--dry-run', action='store_true',
        help='Show what would be generated without calling the LLM',
    )
    args = parser.parse_args()

    if args.all_fixtures:
        fixtures = sorted(
            n for n in os.listdir(FIXTURES)
            if os.path.isdir(os.path.join(FIXTURES, n))
        )
    elif args.fixtures:
        fixtures = args.fixtures
    else:
        parser.error('Specify fixture name(s) or --all')

    print(f'Regenerating draft patterns (backend: {args.backend}) …')
    ok = err = 0
    for name in fixtures:
        if regenerate(name, args.backend, dry_run=args.dry_run):
            ok += 1
        else:
            err += 1

    print(f'\nDone — {ok} succeeded, {err} failed.')
    if err:
        sys.exit(1)


if __name__ == '__main__':
    main()
