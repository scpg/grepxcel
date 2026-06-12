"""Tests for `grepxcel validate-pattern` (grepxcel.pattern_check)."""
import io

import openpyxl

from grepxcel.pattern_check import check_pattern, run_validate


def _write(rows, tmp_path, name='pattern.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    path = str(tmp_path / name)
    wb.save(path)
    return path


_VALID = [['var:', 'x.v', 'string', '.*'], ['START:'], ['cell:A1', 'x.v'], ['END:']]


def test_valid_pattern(tmp_path):
    r = check_pattern(_write(_VALID, tmp_path))
    assert r.valid and not r.errors
    assert r.n_fields == 1 and r.n_steps == 1


def test_unknown_type_is_invalid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'currncy', '.*'], ['START:'], ['cell:A1', 'x.v'], ['END:']], tmp_path))
    assert not r.valid
    assert any('Unknown field type' in e for e in r.errors)


def test_undefined_field_is_invalid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'string', '.*'], ['START:'], ['cell:A1', 'ghost'], ['END:']], tmp_path))
    assert not r.valid
    assert any("'ghost'" in e and 'never defined' in e for e in r.errors)


def test_empty_sequence_is_invalid(tmp_path):
    r = check_pattern(_write([['var:', 'x.v', 'string', '.*']], tmp_path))  # no START:
    assert not r.valid
    assert any('No extraction steps' in e for e in r.errors)


def test_unused_field_warns_but_valid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'string', '.*'], ['var:', 'y.z', 'string', '.*'],
         ['START:'], ['cell:A1', 'x.v'], ['END:']], tmp_path))
    assert r.valid                              # unused is a warning, not an error
    assert any("'y.z'" in w for w in r.warnings)


def test_redos_regex_is_invalid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'string', r'(a+)+$'], ['START:'], ['cell:A1', 'x.v'], ['END:']], tmp_path))
    assert not r.valid
    assert any('regex' in e.lower() for e in r.errors)


def test_missing_file(tmp_path):
    r = check_pattern(str(tmp_path / 'nope.xlsx'))
    assert not r.valid
    assert r.errors


def test_run_validate_exit_codes(tmp_path):
    good = _write(_VALID, tmp_path, 'good.xlsx')
    bad = _write([['var:', 'x.v', 'currncy', '.*'], ['START:'], ['cell:A1', 'x.v'], ['END:']],
                 tmp_path, 'bad.xlsx')
    assert run_validate([good], out=io.StringIO()) == 0
    assert run_validate([bad], out=io.StringIO()) == 1
    assert run_validate([good, bad], out=io.StringIO()) == 1   # any invalid → non-zero


def test_verbose_shows_fields_and_sequence(tmp_path):
    buf = io.StringIO()
    run_validate([_write(_VALID, tmp_path)], verbose=True, out=buf)
    out = buf.getvalue()
    assert 'VALID' in out and 'fields:' in out and 'extraction sequence:' in out and 'x.v' in out
