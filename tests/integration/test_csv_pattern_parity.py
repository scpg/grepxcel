"""
CSV ↔ XLSX pattern parity sweep.

The strongest guarantee that the new CSV pattern source breaks nothing: for every
fixture that ships a reference pattern.xlsx, convert that pattern faithfully to
CSV and run extraction against the same data.xlsx with BOTH pattern formats. The
two results must be byte-for-byte identical.

Because both formats are read into the same 2D-grid IR and then parsed by the
same semantic parser, any divergence here means the CSV reader produced a
different grid — exactly the regression this sweep is designed to catch.

Correctness of the xlsx extraction itself is covered by test_engine.py; this file
only asserts CSV == XLSX.
"""

import csv
import os

import openpyxl
import pytest

from engine import Engine, Logger, VerbosityLevel

_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')

_FULL_FIXTURES: list[str] = sorted(
    name for name in os.listdir(_FIXTURES_DIR)
    if os.path.isdir(os.path.join(_FIXTURES_DIR, name))
    and os.path.exists(os.path.join(_FIXTURES_DIR, name, 'pattern.xlsx'))
    and os.path.exists(os.path.join(_FIXTURES_DIR, name, 'data.xlsx'))
)

# Fixtures whose reference pattern targets a non-active sheet. Using the right
# sheet makes the parity comparison exercise real extracted data rather than an
# empty result. Parity holds regardless of sheet (both runs use the same one);
# these overrides just make the test meaningful.
_SHEET_OVERRIDES: dict[str, str] = {
    '12_multi_sheet': 'Details',
    '06_merged_cells': '2025',
    '04_excel_template_invoce': 'Invoice',
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


@pytest.mark.parametrize('fixture', _FULL_FIXTURES)
def test_csv_pattern_matches_xlsx_pattern(fixture, tmp_path):
    """CSV-sourced extraction must equal xlsx-sourced extraction for every fixture."""
    folder      = os.path.join(_FIXTURES_DIR, fixture)
    xlsx_pattern = os.path.join(folder, 'pattern.xlsx')
    data_file    = os.path.join(folder, 'data.xlsx')
    sheet        = _SHEET_OVERRIDES.get(fixture)

    csv_pattern = str(tmp_path / 'pattern.csv')
    _xlsx_pattern_to_csv(xlsx_pattern, csv_pattern)

    xlsx_result = _extract(xlsx_pattern, data_file, sheet)
    csv_result  = _extract(csv_pattern, data_file, sheet)

    assert csv_result == xlsx_result, (
        f'{fixture}: CSV pattern produced a different result than the xlsx pattern'
    )


def test_sweep_is_not_vacuous():
    """Guard against the parity sweep passing only because every result is empty.
    Confirms the conversion + extraction actually exercises real data."""
    import tempfile

    non_empty = 0
    for fixture in _FULL_FIXTURES:
        folder    = os.path.join(_FIXTURES_DIR, fixture)
        data_file = os.path.join(folder, 'data.xlsx')
        sheet     = _SHEET_OVERRIDES.get(fixture)
        with tempfile.TemporaryDirectory() as td:
            csv_pattern = os.path.join(td, 'pattern.csv')
            _xlsx_pattern_to_csv(os.path.join(folder, 'pattern.xlsx'), csv_pattern)
            if _non_empty(_extract(csv_pattern, data_file, sheet)):
                non_empty += 1

    # The large majority of fixtures must yield real extracted data via CSV.
    assert non_empty >= len(_FULL_FIXTURES) * 0.7, (
        f'Only {non_empty}/{len(_FULL_FIXTURES)} CSV extractions were non-empty; '
        f'the parity sweep may be vacuous.'
    )
