"""Lint security integration tests using generated/downloaded fixtures.

Fixtures are produced by scripts/generate_lint_fixtures.py and written to
tests/fixtures/lint/ (gitignored).  Each test skips gracefully when its
fixture is absent so the normal unit-test run is unaffected.

Run in CI via the lint-security job (pull_request only), which executes the
generator script before pytest.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from grepxcel.lint import FAIL, INFO, OK, WARN, lint_file

_FIXTURES = Path(__file__).parent.parent / 'fixtures' / 'lint'


def _f(name: str) -> str:
    p = _FIXTURES / name
    if not p.exists():
        pytest.skip(f'fixture not present (run scripts/generate_lint_fixtures.py): {name}')
    return str(p)


# ── ZIP bomb ──────────────────────────────────────────────────────────────────

def test_zip_bomb_is_rejected():
    results = lint_file(_f('zip_bomb.xlsx'))
    assert any(r[0] == FAIL and 'integrity' in r[1].lower() for r in results)
    assert any('ratio' in r[2].lower() or 'bomb' in r[2].lower()
               or 'zip' in r[2].lower()
               for r in results if r[0] == FAIL)


# ── XXE injection ─────────────────────────────────────────────────────────────
# Fixture generated using the technique from XXElixir (github.com/kljunowsky/XXElixir)
# by Milan Jovic. Attribution pending — see memory/project_xxe_attribution_pending.md.

def test_xxe_injected_file_is_rejected():
    """A file with <!DOCTYPE/<!ENTITY in workbook.xml must be caught before openpyxl."""
    results = lint_file(_f('xxe_injected.xlsx'))
    assert any(r[0] == FAIL for r in results), (
        'Expected FAIL for XXE-injected file but got: '
        + str([r for r in results if r[0] != INFO])
    )
    assert any(
        'xxe' in r[2].lower()
        or 'doctype' in r[2].lower()
        or 'entity' in r[2].lower()
        or 'injection' in r[2].lower()
        or 'xml' in r[2].lower()
        for r in results if r[0] == FAIL
    )


# ── Corrupt ZIP ───────────────────────────────────────────────────────────────

def test_corrupt_zip_is_rejected():
    results = lint_file(_f('corrupt_zip.xlsx'))
    assert any(r[0] == FAIL for r in results)


# ── Binary/macro extension ────────────────────────────────────────────────────

def test_xlsb_extension_rejected():
    results = lint_file(_f('xlsb_binary.xlsb'))
    assert any(r[0] == FAIL and 'macro' in r[2].lower() or 'binary' in r[2].lower()
               for r in results)


# ── OLE-encrypted file (real file from oletools test suite) ───────────────────

def test_ole_encrypted_detected():
    """Real OLE-encrypted .xlsx (from decalage2/oletools test data) must be detected."""
    results = lint_file(_f('ole_encrypted.xlsx'))
    assert any(
        r[0] == FAIL and (
            'encrypt' in r[2].lower()
            or 'ole' in r[2].lower()
            or 'protect' in r[2].lower()
        )
        for r in results
    )


# ── Inflated dimensions ───────────────────────────────────────────────────────

def test_inflated_dimensions_warns():
    results = lint_file(_f('inflated_dimensions.xlsx'))
    assert any(r[0] == WARN and 'inflat' in r[2].lower() for r in results)


# ── Merged cells ──────────────────────────────────────────────────────────────

def test_merged_cells_warns():
    results = lint_file(_f('merged_cells.xlsx'))
    assert any(r[0] == WARN and 'merge' in r[2].lower() for r in results)


# ── Formula cells ─────────────────────────────────────────────────────────────

def test_formula_cells_warns():
    results = lint_file(_f('formula_cells.xlsx'))
    assert any(r[0] == WARN and 'formula' in r[2].lower() for r in results)


# ── Empty sheet ───────────────────────────────────────────────────────────────

def test_empty_sheet_warns():
    results = lint_file(_f('empty_sheet.xlsx'))
    assert any(r[0] == WARN and 'empty' in r[2].lower() for r in results)
