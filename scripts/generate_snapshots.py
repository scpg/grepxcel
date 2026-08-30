#!/usr/bin/env python3
"""
Generate or refresh extraction snapshots for one or more fixtures.

A snapshot is the JSON output of running 'grepxcel extract' with a specific
pattern against the fixture's data file.  Snapshots are stored in the fixture
folder and used by tests/integration/test_snapshots.py for regression testing.

One snapshot file is produced per pattern variant found in the fixture folder:

    tests/fixtures/{name}/{name}_snapshot-from-{backend}.json
    tests/fixtures/{name}/{name}_snapshot-manual.json

Usage
-----
    # One fixture — all pattern variants
    python3 scripts/generate_snapshots.py 01_simple_invoice

    # Several fixtures
    python3 scripts/generate_snapshots.py 01_simple_invoice 02_product_catalog

    # All fixtures
    python3 scripts/generate_snapshots.py --all

    # One fixture, one specific source
    python3 scripts/generate_snapshots.py 01_simple_invoice --source draft
    python3 scripts/generate_snapshots.py 19_blood_pressure_tracker --source manual

When to run
-----------
Run this script after:
  - Adding or updating any pattern file (pattern-from-*.xlsx or pattern-manual.xlsx)
  - Any intentional engine change that alters the JSON output format

Commit the new/updated snapshot files alongside the pattern changes.

WARNING: Overwriting a snapshot is permanent in git once committed.  Double-check
the diff before committing — a changed snapshot means the engine output changed.
"""

import json
import os
import sys
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

from grepxcel import Engine, Logger, VerbosityLevel
from tests.conftest import find_data_file, find_pattern_for_backend, find_extraction_config, _BACKENDS

ROOT     = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
FIXTURES = os.path.join(ROOT, 'tests', 'fixtures')

_ALL_SOURCES = (*_BACKENDS, 'manual')


def _snapshot_path(folder: str, source: str) -> str:
    name = os.path.basename(folder)
    filename = (f'{name}_snapshot-manual.json' if source == 'manual'
                else f'{name}_snapshot-from-{source}.json')
    return os.path.join(folder, filename)


def generate_snapshot(fixture_name: str, source: str) -> bool:
    """Run extraction for one fixture/source combination and write the snapshot.

    Returns True if a snapshot was written, False if skipped (no pattern found).
    """
    folder = os.path.join(FIXTURES, fixture_name)
    pattern = find_pattern_for_backend(folder, source)
    if not pattern:
        return False

    data = find_data_file(folder)
    if not os.path.exists(data):
        print(f'  [{fixture_name}/{source}] ERROR: data file not found', file=sys.stderr)
        return False

    extra  = find_extraction_config(folder)
    lg     = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pattern, data_file=data, logger=lg, **extra)

    # Normalise via JSON round-trip so datetimes become strings.
    normalised = json.loads(json.dumps(result, default=str, ensure_ascii=False))

    out = _snapshot_path(folder, source)
    with open(out, 'w', encoding='utf-8') as fh:
        json.dump(normalised, fh, indent=2, ensure_ascii=False)
        fh.write('\n')

    print(f'  [{fixture_name}] ✓ {os.path.basename(out)}')
    return True


def process_fixture(fixture_name: str, sources: list[str]) -> int:
    """Generate snapshots for all requested sources of one fixture.

    Returns the number of snapshots written.
    """
    written = 0
    for source in sources:
        if generate_snapshot(fixture_name, source):
            written += 1
    if not written:
        print(f'  [{fixture_name}] no matching patterns — skipped')
    return written


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        'fixtures', nargs='*',
        help='Fixture name(s), e.g. 01_simple_invoice',
    )
    parser.add_argument(
        '--all', dest='all_fixtures', action='store_true',
        help='Generate snapshots for every fixture',
    )
    parser.add_argument(
        '--source', choices=list(_ALL_SOURCES),
        help='Only generate snapshot for this source (backend or "manual")',
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

    sources = [args.source] if args.source else list(_ALL_SOURCES)

    print('Generating snapshots …')
    total = 0
    for name in fixtures:
        total += process_fixture(name, sources)

    print(f'\nDone — {total} snapshot(s) written.')


if __name__ == '__main__':
    main()
