"""
CSV ↔ XLSX pattern parity sweep.

Level 1 — Synthetic parity (parser regression)
    For every fixture that has a reference pattern xlsx, convert it to a
    temporary CSV on-the-fly and run extraction with both.  The two results
    must be byte-for-byte identical.  Catches parser bugs: if the CSV reader
    produces a different grid than the xlsx reader, extraction diverges.

Level 2 — Real authored CSV (format regression)
    For every fixture that ships an actual committed pattern csv file
    (pattern-from-draft.csv or pattern-manual.csv), run extraction with that
    real file and assert it matches the xlsx result.  Catches encoding,
    delimiter, and quoting issues in real authored CSV files that the synthetic
    conversion test cannot exercise.

Because both formats are read into the same 2D-grid IR and then parsed by the
same semantic parser, any divergence means the CSV reader produced a different
grid — exactly the regression these sweeps are designed to catch.

Correctness of the xlsx extraction itself is covered by test_engine.py; this
file only asserts CSV == XLSX.
"""

import csv
import os

import openpyxl
import pytest

from grepxcel import Engine, Logger, VerbosityLevel
from tests.conftest import find_pattern_xlsx, find_pattern_csv

_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')

_ALL_FIXTURE_NAMES: list[str] = sorted(
    name for name in os.listdir(_FIXTURES_DIR)
    if os.path.isdir(os.path.join(_FIXTURES_DIR, name))
    and not os.path.exists(os.path.join(_FIXTURES_DIR, name, 'UNDER_REVIEW'))
)

# Level 1: fixtures with a pattern xlsx + data xlsx
_LEVEL1_FIXTURES: list[str] = [
    name for name in _ALL_FIXTURE_NAMES
    if find_pattern_xlsx(os.path.join(_FIXTURES_DIR, name)) is not None
    and os.path.exists(os.path.join(_FIXTURES_DIR, name, 'data.xlsx'))
]

# Level 2: fixtures with a real committed pattern csv whose tier matches the xlsx.
# If the xlsx is pattern-manual.xlsx, we only compare against pattern-manual.csv
# (not pattern-from-draft.csv) — they are different patterns with different fields.
def _level2_fixtures():
    result = []
    for name in _ALL_FIXTURE_NAMES:
        folder = os.path.join(_FIXTURES_DIR, name)
        xlsx = find_pattern_xlsx(folder)
        csv = find_pattern_csv(folder)
        if not xlsx or not csv or not os.path.exists(os.path.join(folder, 'data.xlsx')):
            continue
        xlsx_base = os.path.basename(xlsx).replace('.xlsx', '')
        csv_base = os.path.basename(csv).replace('.csv', '')
        if xlsx_base == csv_base:
            result.append(name)
    return result

_LEVEL2_FIXTURES: list[str] = _level2_fixtures()

# Fixtures whose reference pattern targets a non-active sheet.
_SHEET_OVERRIDES: dict[str, str] = {
    '12_multi_sheet': 'Details',
    '06_merged_cells': '2025',
    '17_excel_template_invoice': 'Invoice',
}


def _xlsx_pattern_to_csv(xlsx_path: str, csv_path: str) -> None:
    """Faithfully transcribe an xlsx pattern grid to CSV.

    Each xlsx row becomes one CSV row; None cells become empty fields. csv.writer
    quotes any field containing a comma (e.g. a regex like \\d{1,3}) automatically,
    so the CSV reader round-trips it back to the identical grid.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    with open(csv_path, 'w', newline='', encoding='utf-8') as fh:
        writer = csv.writer(fh)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(['' if v is None else v for v in row])
    wb.close()


def _extract(pattern_path: str, data_path: str, sheet) -> dict:
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    return Engine().process(
        pattern_file=pattern_path, data_file=data_path, logger=lg, **kwargs,
    )


def _non_empty(result: dict) -> bool:
    return bool(result) and any(result.values())


# ---------------------------------------------------------------------------
# Level 1 — synthetic parity (parser regression)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize('fixture', _LEVEL1_FIXTURES)
def test_csv_pattern_matches_xlsx_pattern(fixture, tmp_path):
    """CSV-sourced extraction must equal xlsx-sourced extraction for every fixture."""
    folder       = os.path.join(_FIXTURES_DIR, fixture)
    xlsx_pattern = find_pattern_xlsx(folder)
    data_file    = os.path.join(folder, 'data.xlsx')
    sheet        = _SHEET_OVERRIDES.get(fixture)

    csv_pattern = str(tmp_path / 'pattern.csv')
    _xlsx_pattern_to_csv(xlsx_pattern, csv_pattern)

    xlsx_result = _extract(xlsx_pattern, data_file, sheet)
    csv_result  = _extract(csv_pattern, data_file, sheet)

    assert csv_result == xlsx_result, (
        f'{fixture}: CSV pattern produced a different result than the xlsx pattern'
    )


def test_level1_sweep_is_not_vacuous():
    """Guard against the parity sweep passing only because every result is empty."""
    import tempfile

    non_empty = 0
    for fixture in _LEVEL1_FIXTURES:
        folder    = os.path.join(_FIXTURES_DIR, fixture)
        data_file = os.path.join(folder, 'data.xlsx')
        sheet     = _SHEET_OVERRIDES.get(fixture)
        with tempfile.TemporaryDirectory() as td:
            csv_pattern = os.path.join(td, 'pattern.csv')
            _xlsx_pattern_to_csv(find_pattern_xlsx(folder), csv_pattern)
            if _non_empty(_extract(csv_pattern, data_file, sheet)):
                non_empty += 1

    assert non_empty >= len(_LEVEL1_FIXTURES) * 0.7, (
        f'Only {non_empty}/{len(_LEVEL1_FIXTURES)} CSV extractions were non-empty; '
        f'the parity sweep may be vacuous.'
    )


# ---------------------------------------------------------------------------
# Level 2 — real authored CSV files (format regression)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize('fixture', _LEVEL2_FIXTURES)
def test_real_csv_pattern_matches_xlsx_pattern(fixture):
    """A committed pattern csv must extract identically to its xlsx counterpart."""
    folder       = os.path.join(_FIXTURES_DIR, fixture)
    xlsx_pattern = find_pattern_xlsx(folder)
    csv_pattern  = find_pattern_csv(folder)
    data_file    = os.path.join(folder, 'data.xlsx')
    sheet        = _SHEET_OVERRIDES.get(fixture)

    xlsx_result = _extract(xlsx_pattern, data_file, sheet)
    csv_result  = _extract(csv_pattern,  data_file, sheet)

    assert csv_result == xlsx_result, (
        f'{fixture}: real CSV pattern ({os.path.basename(csv_pattern)}) '
        f'produced a different result than the xlsx pattern'
    )


def test_level2_sweep_coverage():
    """Confirm Level-2 fixtures list is non-empty once CSV files are committed."""
    if not _LEVEL2_FIXTURES:
        pytest.skip('No committed pattern csv files found — Level-2 sweep has no fixtures yet')
    assert len(_LEVEL2_FIXTURES) >= 1
