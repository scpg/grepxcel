"""Phase 5 — the number format re-types the value (spec §6).

A numeric Excel cell has no intrinsic date/time-ness: the number format decides,
and openpyxl applies that decision on read. So the Python type that reaches
grepxcel can differ from the one that was written, which changes which ``var:``
types accept the cell. Each ``catalog.RETYPING_CASES`` entry pins all four
observable consequences:

  1. the exact value openpyxl returns,
  2. what ``classify_value`` (so ``profile``) calls it,
  3. which declared ``var:`` types ``validate_type`` accepts it for,
  4. and — end to end through a real pattern — that ``extract`` agrees.

(4) is the one that matters: (1)–(3) are unit-level, while (4) proves the whole
pipeline behaves the way the table says, including for the counter-intuitive
cases (``12`` with ``HH:MM`` is a *datetime*, ``General`` turns a date back into
a bare serial).
"""
from __future__ import annotations

import openpyxl
import pytest

from grepxcel.cell_taxonomy import classify_value
from grepxcel.engine import Engine
from grepxcel.logger import Logger, Severity, VerbosityLevel
from grepxcel.utils import validate_type
from tests.oracle.catalog import RETYPING_CASES, RetypingCase
from tests.oracle.manifest import normalise
from tests.unit.test_random_type_fixtures import write_data_xlsx, write_pattern_csv

pytestmark = pytest.mark.oracle

IDS = [c.case_id for c in RETYPING_CASES]

#: The declared types a re-typing case is checked against end to end. Kept to the
#: ones whose acceptance actually varies across the table.
DECLARED_TYPES = ('integer', 'number', 'currency', 'percentage', 'date',
                  'datetime', 'time', 'duration', 'boolean', 'string')


def _round_trip(tmp_path, case: RetypingCase):
    path = str(tmp_path / 'data.xlsx')
    write_data_xlsx(path, {(1, 1): (case.written, case.number_format)})
    cell = openpyxl.load_workbook(path, data_only=True).active.cell(row=1, column=1)
    return cell, path


@pytest.mark.parametrize('case', RETYPING_CASES, ids=IDS)
def test_read_back_value_is_exactly_as_recorded(tmp_path, case):
    """openpyxl's cast is the premise of everything else here, so it is asserted
    rather than assumed — a change in openpyxl fails loudly."""
    cell, _ = _round_trip(tmp_path, case)
    assert type(cell.value) is type(case.read_back), (
        f'{case.case_id}: wrote {case.written!r} with {case.number_format!r}; '
        f'catalog says it reads back as {type(case.read_back).__name__} '
        f'but openpyxl returned {type(cell.value).__name__} ({cell.value!r}). '
        f'Note: {case.note}'
    )
    assert cell.value == case.read_back, (
        f'{case.case_id}: read back {cell.value!r}, catalog says '
        f'{case.read_back!r}. Note: {case.note}'
    )


@pytest.mark.parametrize('case', RETYPING_CASES, ids=IDS)
def test_profile_classification_matches(tmp_path, case):
    cell, _ = _round_trip(tmp_path, case)
    prof = classify_value(cell.value, cell.number_format)
    assert (prof.storage_type, prof.semantic_type) == (
        case.profile_storage, case.profile_semantic), (
        f'{case.case_id}: classify_value says '
        f'{prof.storage_type}/{prof.semantic_type}, catalog says '
        f'{case.profile_storage}/{case.profile_semantic}. Note: {case.note}'
    )


@pytest.mark.parametrize('case', RETYPING_CASES, ids=IDS)
def test_validate_type_acceptance_set_matches(tmp_path, case):
    """The whole point: the format decides which var: types accept the cell."""
    cell, _ = _round_trip(tmp_path, case)
    actual = frozenset(t for t in DECLARED_TYPES
                       if validate_type(cell.value, t, '.*')[0])
    assert actual == case.accepted_for, (
        f'{case.case_id}: validate_type accepts {sorted(actual)}, catalog says '
        f'{sorted(case.accepted_for)} '
        f'(only in actual: {sorted(actual - case.accepted_for)}, '
        f'only in catalog: {sorted(case.accepted_for - actual)}). '
        f'Note: {case.note}'
    )


@pytest.mark.parametrize('declared', DECLARED_TYPES)
@pytest.mark.parametrize('case', RETYPING_CASES, ids=IDS)
def test_extract_agrees_end_to_end(tmp_path, case, declared):
    """Run a real pattern declaring *declared* against the re-typed cell: the
    value always comes out (the engine writes unconditionally) and a warning is
    raised exactly when the catalog says the type is not accepted."""
    pattern = str(tmp_path / 'pattern.csv')
    data = str(tmp_path / 'data.xlsx')
    write_pattern_csv(pattern, [
        ['lbl:', 'anchor', 'string', 'ANCHOR'],
        ['var:', 'field', declared, '.*'],
        [],
        ['START:'],
        ['cell:next', 'anchor'],
        ['cell:next', 'field'],
        ['END:'],
    ])
    write_data_xlsx(data, {
        (1, 1): ('ANCHOR', None),
        (1, 2): (case.written, case.number_format),
    })

    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pattern, data_file=data, logger=lg)

    assert 'field' in result, f'{case.case_id}/{declared}: field missing: {result}'
    assert normalise(result['field']) == normalise(case.read_back), (
        f'{case.case_id}/{declared}: extracted {result["field"]!r}, '
        f'expected the re-typed value {case.read_back!r}'
    )

    expect_accepted = declared in case.accepted_for
    issues = [r for r in lg._records
              if r.severity in (Severity.WARNING, Severity.ERROR)
              and getattr(r, 'event', '') == 'value_mismatch']
    if expect_accepted:
        assert not issues, (
            f'{case.case_id}: var: {declared} should accept the re-typed value '
            f'{case.read_back!r} but extract warned: '
            f'{[r.message for r in issues]}. Note: {case.note}'
        )
    else:
        assert issues, (
            f'{case.case_id}: var: {declared} should REJECT the re-typed value '
            f'{case.read_back!r} but extract raised no mismatch warning. '
            f'Note: {case.note}'
        )
        # The reason must name the cause, same standard as §10.4.
        ok, reason = validate_type(case.read_back, declared, '.*')
        assert not ok
        assert any(reason in str(r.message) for r in issues), (
            f'{case.case_id}/{declared}: warning does not name the reason.\n'
            f'  validate_type reason: {reason!r}\n'
            f'  messages            : {[str(r.message) for r in issues]}'
        )


def test_table_is_not_vacuous():
    assert len(RETYPING_CASES) >= 15
    assert len(set(IDS)) == len(IDS), 'duplicate retyping case_id'
