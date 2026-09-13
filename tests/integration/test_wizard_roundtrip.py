"""Wizard round-trip parity tests.

For every fixture with a committed manual pattern, this test verifies that:

    manual pattern  ──►  Engine().process()  ──►  JSON result  (A)
                                equals
    wizard preload  ──►  CSV export  ──►  Engine().process()  ──►  JSON result  (B)

If (A) == (B) for a fixture, the wizard's round-trip is extraction-equivalent
to the original hand-crafted pattern — same fields, same values, same structure.

This suite is the primary trust-building mechanism for the web wizard.  Each
fixture that passes proves the wizard can faithfully reproduce that pattern.

Run all:
    .venv/bin/pytest tests/integration/test_wizard_roundtrip.py -v

Run one:
    .venv/bin/pytest tests/integration/test_wizard_roundtrip.py -v -k 03_purchase_order

Show preload warnings even on passing tests:
    .venv/bin/pytest tests/integration/test_wizard_roundtrip.py -v -s

Status legend (look at the -v output):
    PASSED  — wizard round-trip is extraction-equivalent ✓
    FAILED  — wizard exports a pattern that extracts differently (a real bug)
    ERROR   — extraction raised an exception (pattern too incomplete to run)
    SKIPPED — fixture not applicable (no manual pattern, or sheet not found)
"""

import json
import os
import tempfile

import openpyxl
import pytest

from grepxcel import Engine, Logger, VerbosityLevel
from grepxcel.wizard import _build_cell_order
from grepxcel.wizard_tui import _choices_to_csv, _preload_from_pattern
from tests.conftest import find_data_file, find_extraction_config, find_pattern_for_backend

_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


# ── Fixture discovery ──────────────────────────────────────────────────────────

def _collect_params():
    """Build the parametrize list from all fixtures with a manual pattern."""
    params = []
    for name in sorted(os.listdir(_FIXTURES_DIR)):
        folder = os.path.join(_FIXTURES_DIR, name)
        if not os.path.isdir(folder):
            continue
        manual_pattern = find_pattern_for_backend(folder, 'manual')
        if manual_pattern is None:
            continue            # no manual pattern — skip silently
        data = find_data_file(folder)
        if not os.path.exists(data):
            continue            # data file missing — skip silently
        params.append(pytest.param(name, folder, data, manual_pattern, id=name))
    return params


_PARAMS = _collect_params()


# ── Helpers ────────────────────────────────────────────────────────────────────

def _normalize(obj) -> object:
    """Normalise an extraction result for comparison.

    Uses the same JSON round-trip as test_snapshots so datetimes become
    strings (``datetime.date(2024, 1, 15)`` → ``"2024-01-15"``).  Both
    sides go through the same normalisation, so any systematic difference
    is a real bug, not a representation artefact.
    """
    return json.loads(json.dumps(obj, default=str, ensure_ascii=False))


def _extract(pattern_path: str, data_path: str, extra: dict) -> object:
    """Run the engine and return the normalised result dict."""
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(
        pattern_file=pattern_path,
        data_file=data_path,
        logger=lg,
        **extra,
    )
    return _normalize(result)


def _open_sheet(data_path: str, extra: dict):
    """Open the workbook and return the correct worksheet.

    Respects the fixture-level ``sheet`` key in extraction_config so the
    wizard preloads the same sheet the engine extracts from.
    """
    wb = openpyxl.load_workbook(data_path, data_only=True)
    sheet_kwarg = extra.get('sheet')
    if sheet_kwarg is None:
        return wb.active
    if isinstance(sheet_kwarg, str):
        return wb[sheet_kwarg]       # by name
    return wb.worksheets[sheet_kwarg]   # by 0-based index


def _diff_summary(ref: object, wiz: object, max_chars: int = 600) -> str:
    """Return a human-readable diff hint for assertion messages."""
    ref_json = json.dumps(ref, indent=2, ensure_ascii=False)
    wiz_json = json.dumps(wiz, indent=2, ensure_ascii=False)

    # Find first line that differs
    ref_lines = ref_json.splitlines()
    wiz_lines = wiz_json.splitlines()
    first_diff = next(
        (i for i, (a, b) in enumerate(zip(ref_lines, wiz_lines)) if a != b),
        min(len(ref_lines), len(wiz_lines)),
    )
    context_start = max(0, first_diff - 2)
    context_end   = min(len(ref_lines), first_diff + 5)

    ref_ctx = '\n'.join(ref_lines[context_start:context_end])
    wiz_ctx = '\n'.join(wiz_lines[context_start:context_end])

    return (
        f'First difference near line {first_diff + 1}:\n'
        f'--- manual pattern (expected) ---\n{ref_ctx}\n\n'
        f'+++ wizard round-trip (got) ---\n{wiz_ctx}'
    )


# ── The test ───────────────────────────────────────────────────────────────────

@pytest.mark.parametrize('fixture_name,folder,data,manual_pattern', _PARAMS)
def test_wizard_roundtrip(fixture_name, folder, data, manual_pattern, capsys):
    """Wizard round-trip must produce extraction-equivalent output to the manual pattern.

    Failure means the wizard exports a subtly wrong pattern.  The diff
    shown in the assertion message points at the first divergence.
    """
    extra = find_extraction_config(folder)

    # ── A: reference extraction (manual pattern) ──────────────────────────────
    ref = _extract(manual_pattern, data, extra)

    # ── B: wizard round-trip ──────────────────────────────────────────────────
    ws = _open_sheet(data, extra)

    choices, preload_cfg, preload_warnings = _preload_from_pattern(ws, manual_pattern)

    # Always print preload warnings so -s shows them even on passing tests
    if preload_warnings:
        print(f'\n[{fixture_name}] preload warnings:')
        for w in preload_warnings:
            print(f'  ⚠  {w}')

    direction     = preload_cfg.get('direction', 'LR')
    currency_sign = preload_cfg.get('currency_sign', '€')
    cells         = _build_cell_order(ws, direction)
    csv_text      = _choices_to_csv(
        ws, choices, cells,
        direction=direction,
        sheet_name=ws.title,
        currency_sign=currency_sign,
    )

    # Write wizard CSV to a temp file and extract
    with tempfile.NamedTemporaryFile(
        mode='w', suffix='.csv', delete=False, encoding='utf-8'
    ) as tf:
        tf.write(csv_text)
        tmp_csv = tf.name

    try:
        wiz = _extract(tmp_csv, data, extra)
    finally:
        os.unlink(tmp_csv)

    # ── Compare ───────────────────────────────────────────────────────────────
    if ref == wiz:
        return   # ✓ pass

    preload_note = (
        'Preload warnings (may explain the diff):\n  ' +
        '\n  '.join(preload_warnings)
        if preload_warnings else 'No preload warnings.'
    )

    pytest.fail(
        f'\nWizard round-trip mismatch for [{fixture_name}].\n\n'
        f'{preload_note}\n\n'
        f'{_diff_summary(ref, wiz)}\n\n'
        f'Full manual JSON length : {len(json.dumps(ref))} chars\n'
        f'Full wizard JSON length : {len(json.dumps(wiz))} chars\n\n'
        f'Hint: load data + manual pattern in the web wizard and compare the\n'
        f'Downloaded CSV against the manual pattern-manual.xlsx to spot the bug.'
    )
