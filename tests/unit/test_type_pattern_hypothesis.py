"""Hypothesis-driven "type -> pattern -> excel file -> verify" tests.

Complements tests/unit/test_random_type_fixtures.py, which deliberately uses
a fixed-seed hand-rolled placer (reproducible without persisted fixtures,
and needed for the harder combined/label-before-table ordering-guarantee
shapes documented there). Here, hypothesis drives what it is actually good
at: exploring far more (type, grid position, valid-vs-invalid) combinations
across the two simplest, single-element shapes -- one lbl:/var: label field,
one 1-row mini-table -- than a handful of fixed seeds ever would, and, if a
failure is ever found, shrinking automatically to the minimal reproducing
(type, row, col, value) case.

Cell VALUES for the "valid" side are hypothesis-generated per type, using
strategies whose bounds are checked directly against validate_type()'s real
acceptance rules (grepxcel/utils.py) -- not guessed:
  - integer/number/currency/percentage: any non-bool int/float, so plain
    st.integers()/st.floats() ranges are safe.
  - boolean: st.booleans().
  - date: isinstance(value, (date, datetime)) -- st.dates() is safe.
  - time/duration: isinstance(value, (time, timedelta)) -- BOTH are accepted
    for BOTH field types, so st.times()/st.timedeltas() are safe for either.
  - string has no real type failure to begin with (regex='.*' always
    matches) -- its valid strategy is a plain printable-text generator.
  - url needs a real scheme + urlparse() success; a general-purpose text
    strategy would need real care to stay safely inside that contract, so
    it draws from a small set of provably-valid URL shapes instead.

"invalid" values reuse the same proven per-type failure modes already used
by test_random_type_fixtures.py's TYPE_CASES (a bool for integer, a non-ISO
string for date, etc.) -- known-good negative examples, not invented extras.
"""
from __future__ import annotations

import csv
import datetime
import os
import shutil
import tempfile

import openpyxl
from hypothesis import given, settings, strategies as st

from grepxcel.engine import Engine
from grepxcel.logger import Logger, Severity, VerbosityLevel

GRID_SIZE = 20

_VALID_STRATEGIES = {
    # Allowlist real printable categories only (Letter/Number/Punctuation/
    # Symbol/space) rather than blacklisting a few -- an unassigned-but-not-
    # Cc/Cs codepoint (category Cn) is neither empty nor a control character,
    # yet utils.is_empty()'s printable-filter treats it as invisible (its
    # job is stripping zero-width/whitespace-only cells, a different concern
    # than this test's). Excluding Cn keeps this test in its own lane
    # (genuine string content) instead of re-testing is_empty()'s own scope
    # (already covered by test_property_based.py/test_utils.py).
    'string': st.text(
        alphabet=st.characters(categories=('L', 'N', 'P', 'S', 'Zs')),
        min_size=1, max_size=40,
    ),
    'integer': st.integers(min_value=-10**6, max_value=10**6),
    'boolean': st.booleans(),
    'date': st.dates(min_value=datetime.date(1990, 1, 1), max_value=datetime.date(2100, 12, 31)),
    'url': st.sampled_from([
        'https://example.com', 'http://example.org/path',
        'https://sub.example.net:8080/a/b?q=1', 'ftp://files.example.com/x',
    ]),
    'currency': st.floats(min_value=0, max_value=1_000_000, allow_nan=False, allow_infinity=False),
    'percentage': st.floats(min_value=0, max_value=1, allow_nan=False, allow_infinity=False),
    'number': st.floats(min_value=-1_000_000, max_value=1_000_000, allow_nan=False, allow_infinity=False),
    'time': st.times(),
    'duration': st.timedeltas(min_value=datetime.timedelta(0), max_value=datetime.timedelta(days=5)),
}

_FORMATS = {
    'date': 'yyyy-mm-dd', 'currency': '$#,##0.00', 'percentage': '0.00%',
    'time': 'HH:MM', 'duration': '[h]:mm',
}

_REGEX = {t: '.*' for t in _VALID_STRATEGIES}

# Known-good negative examples, matching TYPE_CASES in test_random_type_fixtures.py.
# 'string' has no real type failure (regex='.*' always matches str(anything)),
# so it is excluded from the invalid-value test rather than given a fake one.
_INVALID_EXAMPLES = {
    'integer': True,          # bool is explicitly rejected by validate_type
    'boolean': 'banana',
    'date': 'not a date',
    'url': 'not a url',
    'currency': True,
    'percentage': 'fifty percent',
    'number': False,
    'time': '9:30am',
    'duration': '2 hours',
}

_TYPE_NAMES = sorted(_VALID_STRATEGIES)


def _write_pattern_csv(path: str, rows: list[list]) -> None:
    with open(path, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        for row in rows:
            w.writerow(row)


def _write_data_xlsx(path: str, cells: dict[tuple[int, int], tuple]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    for (row, col), (value, fmt) in cells.items():
        cell = ws.cell(row=row, column=col)
        cell.value = value
        if fmt:
            cell.number_format = fmt
    wb.save(path)


def _run(pattern_path: str, data_path: str) -> dict:
    lg = Logger(level=VerbosityLevel.QUIET)
    return Engine().process(pattern_file=pattern_path, data_file=data_path, logger=lg)


def _run_with_issues(pattern_path: str, data_path: str) -> tuple[dict, list]:
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pattern_path, data_file=data_path, logger=lg)
    issues = [r for r in lg._records if r.severity in (Severity.WARNING, Severity.ERROR)]
    return result, issues


class _TmpFiles:
    """One fresh tmpdir per hypothesis example; cleaned up on __exit__."""

    def __enter__(self):
        self.dir = tempfile.mkdtemp()
        return os.path.join(self.dir, 'pattern.csv'), os.path.join(self.dir, 'data.xlsx')

    def __exit__(self, *exc):
        shutil.rmtree(self.dir, ignore_errors=True)


# ── Shape: single var: label at a hypothesis-random grid position ──────────

class TestSingleLabelTypePattern:
    @given(
        type_name=st.sampled_from(_TYPE_NAMES),
        row=st.integers(min_value=1, max_value=GRID_SIZE),
        col=st.integers(min_value=1, max_value=GRID_SIZE - 1),  # value cell sits at col+1
        value_data=st.data(),
    )
    @settings(max_examples=60, deadline=None)
    def test_valid_value_extracts_correctly(self, type_name, row, col, value_data):
        value = value_data.draw(_VALID_STRATEGIES[type_name])
        fmt = _FORMATS.get(type_name)

        with _TmpFiles() as (pattern_path, data_path):
            _write_pattern_csv(pattern_path, [
                ['lbl:', 'anchor', 'string', 'ANCHOR'],
                ['var:', 'field', type_name, _REGEX[type_name]],
                [],
                ['START:'],
                ['cell:next', 'anchor'],
                ['cell:next', 'field'],
                ['END:'],
            ])
            _write_data_xlsx(data_path, {
                (row, col): ('ANCHOR', None),
                (row, col + 1): (value, fmt),
            })

            result = _run(pattern_path, data_path)
            assert 'field' in result, f'type={type_name} value={value!r} result={result}'

    @given(
        type_name=st.sampled_from(sorted(_INVALID_EXAMPLES)),
        row=st.integers(min_value=1, max_value=GRID_SIZE),
        col=st.integers(min_value=1, max_value=GRID_SIZE - 1),
    )
    @settings(max_examples=30, deadline=None)
    def test_invalid_value_reported_not_silently_accepted(self, type_name, row, col):
        value = _INVALID_EXAMPLES[type_name]

        with _TmpFiles() as (pattern_path, data_path):
            _write_pattern_csv(pattern_path, [
                ['lbl:', 'anchor', 'string', 'ANCHOR'],
                ['var:', 'field', type_name, _REGEX[type_name]],
                [],
                ['START:'],
                ['cell:next', 'anchor'],
                ['cell:next', 'field'],
                ['END:'],
            ])
            _write_data_xlsx(data_path, {
                (row, col): ('ANCHOR', None),
                (row, col + 1): (value, None),
            })

            result, issues = _run_with_issues(pattern_path, data_path)
            # Validation failure does NOT null out the value (engine.py's
            # _process_cell() writes it unconditionally) -- it is reported
            # as an issue instead. See test_random_type_fixtures.py.
            assert result.get('field') == value
            assert issues, f'expected a validation issue for invalid {type_name} value {value!r}'


# ── Shape: single 1-row mini-table at a hypothesis-random grid position ────

class TestSingleMiniTableTypePattern:
    @given(
        type_name=st.sampled_from(_TYPE_NAMES),
        row=st.integers(min_value=1, max_value=GRID_SIZE - 1),  # data row sits at row+1
        col=st.integers(min_value=1, max_value=GRID_SIZE),
        value_data=st.data(),
    )
    @settings(max_examples=60, deadline=None)
    def test_valid_value_extracts_correctly(self, type_name, row, col, value_data):
        value = value_data.draw(_VALID_STRATEGIES[type_name])
        fmt = _FORMATS.get(type_name)

        with _TmpFiles() as (pattern_path, data_path):
            _write_pattern_csv(pattern_path, [
                ['lbl:', 'col_header', 'string', 'COLHEAD'],
                ['var:', 'cell_value', type_name, _REGEX[type_name]],
                [],
                ['START:'],
                ['table:*'],
                ['', 'HEADER:1', 'col_header'],
                ['', 'DATA:*', 'cell_value'],
                ['END:'],
            ])
            _write_data_xlsx(data_path, {
                (row, col): ('COLHEAD', None),
                (row + 1, col): (value, fmt),
            })

            result = _run(pattern_path, data_path)
            assert 'table_0' in result, f'type={type_name} value={value!r} result={result}'
            assert len(result['table_0']) == 1
