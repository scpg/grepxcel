"""Post-installation smoke tests.

Run after 'pip install grepxcel' in a clean environment:

    pip install pytest
    pytest tests/smoke/ --import-mode=importlib -v

--import-mode=importlib is REQUIRED: it stops pytest from prepending the
project root to sys.path, so 'import grepxcel' resolves to the package in
site-packages — the installed artifact — not the source tree.  These tests
validate the distribution, not the source.

In CI this suite fires after each publish step (TestPyPI and PyPI proper).
The test file has no dependency on any grepxcel test infrastructure; it uses
only the installed package, the stdlib, and pytest itself.
"""
import json
import os
import subprocess
import sys
import tempfile

import pytest

import grepxcel
from grepxcel.examples_generator import EXAMPLES


# ── helper ────────────────────────────────────────────────────────────────────

def _cli(*args):
    r = subprocess.run(
        [sys.executable, '-m', 'grepxcel', *args],
        capture_output=True, text=True,
    )
    return r.returncode, r.stdout, r.stderr


# ── version & import ──────────────────────────────────────────────────────────

def test_version_is_set():
    assert grepxcel.__version__
    assert grepxcel.__version__.count('.') >= 1


def test_cli_version_flag():
    rc, out, _ = _cli('--version')
    assert rc == 0
    assert grepxcel.__version__ in out


def test_cli_help_exits_zero():
    rc, _, _ = _cli('--help')
    assert rc == 0


# ── Python API — all bundled examples extract cleanly ─────────────────────────

@pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
def test_api_extract_non_empty(ex):
    result = grepxcel.extract(ex['pattern'], ex['data'])
    assert isinstance(result, dict) and result, \
        f"{ex['name']}: extract() returned empty or non-dict"


# ── CLI — JSON output parses cleanly and is non-empty ────────────────────────

@pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
def test_cli_extract_valid_json(ex):
    rc, out, err = _cli('extract', '-p', ex['pattern'], ex['data'], '-q')
    # rc=0: clean extraction; rc=1: extraction succeeded but some fields had warnings
    # (e.g. 04_loan_schedule has a Totals footer row that triggers warnings by design).
    # rc=2 would mean a hard error (pattern invalid, --strict failure, etc.).
    assert rc in (0, 1), f"unexpected exit {rc}; stderr: {err[:300]}"
    data = json.loads(out)
    assert isinstance(data, dict) and data


# ── CLI — validate-pattern exits 0 for every bundled example ─────────────────

@pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
def test_cli_validate_pattern_ok(ex):
    rc, _, err = _cli('validate-pattern', ex['pattern'], '-q')
    assert rc == 0, f"validate-pattern failed for {ex['name']}: {err[:300]}"


# ── Expected output structure — catches silent field regressions ──────────────

# For each example: scalar groups (dict) and table groups (non-empty list).
# Scalar field assertions use the nested-format path: result[group][field].
_EXPECTED: dict[str, dict] = {
    '01_simple_invoice': {
        'scalars': {'invoice': ['number', 'date'], 'client': ['name']},
    },
    '02_product_catalog': {
        'tables': ['catalog'],
    },
    '03_expense_report': {
        'scalars': {'employee': ['name', 'department']},
        'tables': ['expense'],
    },
    '04_loan_schedule': {
        'scalars': {'loan': ['amount']},
        'tables': ['schedule'],
    },
}


@pytest.mark.parametrize('name', _EXPECTED)
def test_expected_groups_and_fields(name):
    ex = next(e for e in EXAMPLES if e['name'] == name)
    result = grepxcel.extract(ex['pattern'], ex['data'])
    spec = _EXPECTED[name]

    for grp, fields in spec.get('scalars', {}).items():
        assert grp in result, f"scalar group '{grp}' missing from {name}"
        assert isinstance(result[grp], dict), \
            f"'{grp}' expected dict, got {type(result[grp]).__name__}"
        for field in fields:
            assert field in result[grp], f"field '{grp}.{field}' missing from {name}"

    for grp in spec.get('tables', []):
        assert grp in result, f"table group '{grp}' missing from {name}"
        rows = result[grp]
        assert isinstance(rows, list) and rows, \
            f"table '{grp}' is empty in {name}"


# ── Piped output must not contain emoji ───────────────────────────────────────

def test_verbose_pipe_no_emoji():
    ex = EXAMPLES[0]
    rc, _, err = _cli('extract', '-p', ex['pattern'], ex['data'], '-v')
    assert rc == 0
    for emoji in ('🟢', '🔴', '🟡', '🔵'):
        assert emoji not in err, f"Emoji {emoji!r} leaked to piped stderr"


# ── generate-examples writes the expected directory structure ─────────────────

def test_generate_examples_structure():
    with tempfile.TemporaryDirectory() as tmp:
        rc, _, err = _cli('generate-examples', '-o', tmp)
        assert rc == 0, f"generate-examples failed: {err[:300]}"
        for ex in EXAMPLES:
            base = os.path.join(tmp, ex['name'])
            assert os.path.isfile(os.path.join(base, 'pattern.xlsx')), \
                f"missing {ex['name']}/pattern.xlsx"
            assert os.path.isfile(os.path.join(base, 'data.xlsx')), \
                f"missing {ex['name']}/data.xlsx"
            assert os.path.isfile(os.path.join(base, 'README.txt')), \
                f"missing {ex['name']}/README.txt"
