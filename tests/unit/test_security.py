"""
Direct tests for the security module's load-bearing guarantees:
  - XXE protection (defusedxml) is live and fail-closed
  - the ReDoS detector is wired to the single re._parser import path
"""

import zipfile

import pytest

from grepxcel.security import (
    SecurityError,
    assert_xxe_protection,
    check_regex_safety,
    validate_file,
    _check_zip_safety,
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
    from grepxcel import security
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


# ─── ZIP-bomb guard measures REAL decompressed size ──────────────────────────

def _make_zip(path, name='big.xml', payload=b'A', repeat=1):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(name, payload * repeat)


def test_legit_small_zip_passes(tmp_path):
    path = tmp_path / 'ok.xlsx'
    _make_zip(path, payload=b'hello world\n', repeat=10)
    _check_zip_safety(str(path), max_uncompressed_mb=50)   # must not raise


def test_real_decompressed_size_over_cap_rejected(tmp_path):
    """A highly compressible member whose real expansion exceeds the cap is
    rejected even though the file on disk is tiny."""
    path = tmp_path / 'big.xlsx'
    _make_zip(path, payload=b'A', repeat=8 * 1024 * 1024)  # 8 MB → compresses tiny
    with pytest.raises(SecurityError, match='ZIP|uncompressed|bomb'):
        _check_zip_safety(str(path), max_uncompressed_mb=1)


def test_falsified_metadata_cannot_bypass_guard(tmp_path, monkeypatch):
    """The guard must count ACTUAL decompressed bytes, not the central-directory
    file_size — so an entry that lies about its size (file_size=0) is still
    caught by the real expansion."""
    path = tmp_path / 'liar.xlsx'
    _make_zip(path, payload=b'A', repeat=8 * 1024 * 1024)  # real 8 MB

    real_infolist = zipfile.ZipFile.infolist

    def lying_infolist(self):
        infos = real_infolist(self)
        for info in infos:
            info.file_size = 0          # forge the declared uncompressed size
        return infos

    monkeypatch.setattr(zipfile.ZipFile, 'infolist', lying_infolist)
    with pytest.raises(SecurityError, match='ZIP|uncompressed|bomb'):
        _check_zip_safety(str(path), max_uncompressed_mb=1)
