"""
Direct tests for the security module's load-bearing guarantees:
  - XXE protection (defusedxml) is live and fail-closed
  - the ReDoS detector is wired to the single re._parser import path
"""

import pytest

from engine.security import (
    SecurityError,
    assert_xxe_protection,
    check_regex_safety,
    validate_file,
)


# ─── XXE / defusedxml guard ───────────────────────────────────────────────────

def test_xxe_protection_is_enabled():
    """openpyxl must be using defusedxml in this environment."""
    import openpyxl.xml
    assert openpyxl.xml.DEFUSEDXML is True
    # assert_xxe_protection() must not raise when protection is on
    assert_xxe_protection() is None


def test_xxe_guard_fails_closed_when_disabled(monkeypatch):
    """If defusedxml is somehow disabled, the guard must raise SecurityError."""
    import openpyxl.xml
    monkeypatch.setattr(openpyxl.xml, 'DEFUSEDXML', False, raising=True)
    with pytest.raises(SecurityError, match='XXE'):
        assert_xxe_protection()


def test_validate_file_runs_xxe_guard_first(monkeypatch, tmp_path):
    """validate_file must refuse to proceed when XXE protection is off,
    before it even looks at the path."""
    import openpyxl.xml
    monkeypatch.setattr(openpyxl.xml, 'DEFUSEDXML', False, raising=True)
    missing = tmp_path / 'nope.xlsx'
    with pytest.raises(SecurityError, match='XXE'):
        validate_file(str(missing))


# ─── ReDoS detector single import path ────────────────────────────────────────

def test_regex_parser_uses_re_internal():
    """We require Python >=3.11, so the parser must come from re._parser
    (no fallback to the deprecated top-level sre_parse module)."""
    from engine import security
    import re._parser as expected
    assert security._sre_parse is expected


def test_safe_regex_accepted():
    # A normal anchored quantifier is fine.
    check_regex_safety(r'PO-[0-9]{4}', 'po_number')


def test_nested_unbounded_quantifier_rejected():
    with pytest.raises(SecurityError, match='ReDoS|backtracking'):
        check_regex_safety(r'(a+)+', 'evil')


def test_invalid_regex_rejected():
    with pytest.raises(SecurityError, match='Invalid regex'):
        check_regex_safety(r'(unclosed', 'bad')
