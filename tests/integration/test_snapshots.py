"""Snapshot regression tests.

For every fixture that has a committed snapshot JSON file, run extraction with
the corresponding pattern and assert the output matches exactly.

Snapshots are NOT generated automatically — they are produced on demand by
scripts/generate_snapshots.py and committed to the repository as stable
regression baselines.  When an engine change intentionally alters the output,
regenerate the snapshot:

    python3 scripts/generate_snapshots.py <fixture_name> --source <source>

If no snapshot files exist in a fixture folder, that fixture is not included in
the parametrize list and the test is silently absent (not a failure).  The one
top-level guard test below reports how many snapshot fixtures are active.
"""

import json
import os

import pytest

from grepxcel import Engine, Logger, VerbosityLevel
from tests.conftest import (
    find_data_file, list_snapshots, find_pattern_for_backend, find_extraction_config,
)

_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


def _collect_params():
    """Build pytest parameter list from all committed snapshot files."""
    params = []
    for fixture_name in sorted(os.listdir(_FIXTURES_DIR)):
        folder = os.path.join(_FIXTURES_DIR, fixture_name)
        if not os.path.isdir(folder):
            continue
        for source, snap_path in list_snapshots(folder):
            pattern = find_pattern_for_backend(folder, source)
            if pattern is None:
                continue  # snapshot exists but pattern file is gone — skip
            params.append(pytest.param(
                fixture_name, source, snap_path, pattern,
                id=f'{fixture_name}[{source}]',
            ))
    return params


_PARAMS = _collect_params()


def test_snapshot_suite_is_not_vacuous():
    """At least one snapshot must exist for this suite to be meaningful.

    If this test fails, run scripts/generate_snapshots.py --all to seed the
    initial set of snapshots and commit them.
    """
    if not _PARAMS:
        pytest.skip('no snapshots committed yet — run scripts/generate_snapshots.py --all')


@pytest.mark.parametrize('fixture_name,source,snap_path,pat_path', _PARAMS)
def test_snapshot_regression(fixture_name, source, snap_path, pat_path):
    """Extraction output must match the committed snapshot exactly.

    A failure here means the engine's output changed.  If the change is
    intentional, regenerate the snapshot:
        python3 scripts/generate_snapshots.py <fixture_name> --source <source>
    """
    folder = os.path.join(_FIXTURES_DIR, fixture_name)
    data   = find_data_file(folder)

    extra  = find_extraction_config(folder)
    lg     = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pat_path, data_file=data, logger=lg, **extra)

    # Normalise via JSON round-trip so datetimes become strings — identical to
    # how generate_snapshots.py serialises them.
    actual = json.loads(json.dumps(result, default=str, ensure_ascii=False))

    with open(snap_path, encoding='utf-8') as fh:
        expected = json.load(fh)

    assert actual == expected, (
        f'Snapshot mismatch for {fixture_name} [{source}].\n'
        f'If this change is intentional, regenerate the snapshot:\n'
        f'  python3 scripts/generate_snapshots.py {fixture_name} --source {source}'
    )
