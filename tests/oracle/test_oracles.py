"""Unit tests for each oracle, on one small hand-picked case each.

Wave 3's gate: before an oracle is parametrized over thousands of matrix cases,
it gets checked here on a single case where the right answer is obvious — both
that it PASSES a good case and that it FAILS a deliberately corrupted manifest.
An oracle that cannot fail is worse than no oracle, so every check_* has a
negative test too.
"""
from __future__ import annotations

import dataclasses

import pytest

from tests.oracle import oracles
from tests.oracle.generator import CaseParams, build_case

pytestmark = pytest.mark.oracle

# An all-accepted scalar case and an all-accepted table case.
ACCEPTED_SCALAR = CaseParams('integer', 'cells2', 'n-int', 'General', 'LR', 0,
                             via='next')
ACCEPTED_ABS = CaseParams('string', 'cells2', 's-text', 'General', 'TD', 1,
                          via='abs')
ACCEPTED_TABLE = CaseParams('integer', 'table2', 'n-int', 'General', 'LR', 0,
                            header_mode='same_header')
ACCEPTED_TABLE_DISTINCT = CaseParams('integer', 'table2', 'n-int', 'General',
                                     'LR', 0, header_mode='distinct_header')
# A rejected case: a bool written where an integer is expected.
REJECTED_SCALAR = CaseParams('integer', 'cells1', 'b-bool', None, 'LR', 0,
                             via='next')

ALL = [ACCEPTED_SCALAR, ACCEPTED_ABS, ACCEPTED_TABLE, ACCEPTED_TABLE_DISTINCT]


@pytest.fixture(scope='module')
def cases(tmp_path_factory):
    root = tmp_path_factory.mktemp('oracles')
    built = {}
    for p in ALL + [REJECTED_SCALAR]:
        built[p.case_id] = build_case(p, str(root / p.case_id))
    return built


def _m(cases, params):
    return cases[params.case_id]


# ── 10.2 lint ─────────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', ALL, ids=[p.case_id for p in ALL])
def test_check_lint_passes_on_generated_files(cases, params):
    oracles.check_lint(_m(cases, params))


def test_check_lint_fails_on_a_missing_file(cases):
    m = dataclasses.replace(_m(cases, ACCEPTED_SCALAR))
    m.data_path = m.data_path + '.nope'
    with pytest.raises(AssertionError, match='lint reported'):
        oracles.check_lint(m)


# ── 10.3 profile ──────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', ALL, ids=[p.case_id for p in ALL])
def test_check_profile_passes(cases, params):
    oracles.check_profile(_m(cases, params))


def test_check_profile_fails_on_a_wrong_semantic_type(cases):
    m = dataclasses.replace(_m(cases, ACCEPTED_SCALAR))
    m.cells = [dataclasses.replace(c) for c in m.cells]
    value_cell = next(c for c in m.cells if c.role == 'value')
    value_cell.profile_semantic = 'currency'    # the file really holds an integer
    with pytest.raises(AssertionError, match='semantic'):
        oracles.check_profile(m)


def test_check_profile_fails_when_a_ref_is_missing_from_the_manifest(cases):
    m = dataclasses.replace(_m(cases, ACCEPTED_SCALAR))
    m.cells = [c for c in m.cells if c.role != 'error']   # drop the error cells
    with pytest.raises(AssertionError, match='ref set'):
        oracles.check_profile(m)


def test_check_profile_sees_the_error_cells(cases):
    """The §4 profile-only type really is being classified, not skipped."""
    m = _m(cases, ACCEPTED_SCALAR)
    errors = [c for c in m.cells if c.role == 'error']
    assert errors
    for spec in errors:
        assert (spec.profile_storage, spec.profile_semantic) == ('e', 'error')
    oracles.check_profile(m)


# ── 10.4 extract ──────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', ALL, ids=[p.case_id for p in ALL])
def test_check_extract_passes_on_accepted_cases(cases, params):
    oracles.check_extract(_m(cases, params))


def test_check_extract_fails_on_a_wrong_expected_value(cases):
    m = dataclasses.replace(_m(cases, ACCEPTED_SCALAR))
    m.expected_extract = dict(m.expected_extract)
    field = next(c.field for c in m.cells if c.role == 'value')
    m.expected_extract[field] = 'definitely-not-the-value'
    oracles._EXTRACT_CACHE.pop(m.case_id, None)
    with pytest.raises(AssertionError, match='manifest expects'):
        oracles.check_extract(m)


def test_check_extract_fails_on_a_wrong_instance_count(cases):
    m = dataclasses.replace(_m(cases, ACCEPTED_TABLE))
    m.tables = [dataclasses.replace(t) for t in m.tables]
    for spec in m.tables:
        spec.expected_instances = 99
    oracles._EXTRACT_CACHE.pop(m.case_id, None)
    with pytest.raises(AssertionError, match='instance'):
        oracles.check_extract(m)


def test_rejected_value_is_still_written_to_the_output(cases):
    """Documented current behaviour (engine._process_cell is unconditional):
    validation failure warns, it does not null the value out."""
    m = _m(cases, REJECTED_SCALAR)
    result, issues = oracles.run_extract(m)
    spec = next(c for c in m.cells if c.role == 'value')
    assert spec.extract_accepted is False
    assert spec.field in result, f'rejected field vanished from {result}'
    assert result[spec.field] == spec.written_value
    assert issues, 'a rejected value raised no warning at all'


# ── 10.5 schema ───────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', ALL, ids=[p.case_id for p in ALL])
def test_check_schema_passes(cases, params):
    oracles.check_schema(_m(cases, params))


def test_check_schema_fails_on_a_mismatched_schema(cases):
    """Corrupt the pattern's declared type and the schema must reject the
    extraction — proving the oracle is not vacuous."""
    m = dataclasses.replace(_m(cases, ACCEPTED_SCALAR))
    text = open(m.pattern_path).read().replace(',integer,', ',date,')
    bad_pattern = m.pattern_path + '.bad.csv'
    with open(bad_pattern, 'w') as fh:
        fh.write(text)
    m.pattern_path = bad_pattern
    oracles._EXTRACT_CACHE.pop(m.case_id, None)
    with pytest.raises(AssertionError, match='does not validate'):
        oracles.check_schema(m)


# ── 10.6 strict ───────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', ALL, ids=[p.case_id for p in ALL])
def test_check_strict_exits_zero_for_all_accepted_cases(cases, params):
    oracles.check_strict(_m(cases, params))


def test_check_strict_exits_two_for_a_rejected_case(cases):
    m = _m(cases, REJECTED_SCALAR)
    assert m.has_rejection
    oracles.check_strict(m)   # asserts exit code 2 internally
