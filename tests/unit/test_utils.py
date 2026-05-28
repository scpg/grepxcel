import datetime
import pytest
from engine.utils import is_empty, validate_type


# ─── is_empty ────────────────────────────────────────────────────────────────

def test_empty_none():
    assert is_empty(None)

def test_empty_empty_string():
    assert is_empty('')

def test_empty_whitespace_only():
    assert is_empty('   ')

def test_empty_non_breaking_space():
    assert is_empty('\xa0')

def test_empty_zero_width_chars():
    assert is_empty('​‌‍')

def test_empty_mixed_invisible():
    assert is_empty(' \xa0​ ')

def test_empty_alias():
    assert is_empty('N/A', empty_aliases=['N/A'])

def test_empty_alias_not_match():
    assert not is_empty('N/A', empty_aliases=['n/a'])

def test_not_empty_regular_string():
    assert not is_empty('hello')

def test_not_empty_zero():
    assert not is_empty(0)

def test_not_empty_false():
    assert not is_empty(False)

def test_not_empty_integer():
    assert not is_empty(42)


# ─── validate_type: string ────────────────────────────────────────────────────

def test_string_match():
    ok, _ = validate_type('AB123456', 'string', r'[A-Z]{2}[0-9]{6}')
    assert ok

def test_string_no_match():
    ok, reason = validate_type('ab123456', 'string', r'[A-Z]{2}[0-9]{6}')
    assert not ok
    assert 'ab123456' in reason

def test_string_dotall_newline():
    # validate_type uses re.DOTALL for strings, so .* must match across newlines
    ok, _ = validate_type('line1\nline2', 'string', r'.*')
    assert ok


# ─── validate_type: integer ──────────────────────────────────────────────────

def test_integer_valid_int():
    ok, _ = validate_type(5, 'integer', r'[1-9][0-9]*')
    assert ok

def test_integer_float_whole():
    ok, _ = validate_type(5.0, 'integer', r'[1-9][0-9]*')
    assert ok

def test_integer_float_non_whole():
    ok, reason = validate_type(5.5, 'integer', r'[1-9][0-9]*')
    assert not ok
    assert 'whole' in reason

def test_integer_bool_rejected():
    ok, reason = validate_type(True, 'integer', r'.*')
    assert not ok
    assert 'boolean' in reason

def test_integer_string_rejected():
    ok, reason = validate_type('5', 'integer', r'[1-9][0-9]*')
    assert not ok
    assert 'integer' in reason

def test_integer_regex_mismatch():
    ok, _ = validate_type(0, 'integer', r'[1-9][0-9]*')
    assert not ok


# ─── validate_type: currency ─────────────────────────────────────────────────

def test_currency_int():
    ok, _ = validate_type(100, 'currency', r'.*')
    assert ok

def test_currency_float():
    ok, _ = validate_type(99.99, 'currency', r'.*')
    assert ok

def test_currency_bool_rejected():
    ok, reason = validate_type(True, 'currency', r'.*')
    assert not ok
    assert 'boolean' in reason

def test_currency_string_rejected():
    ok, reason = validate_type('€100', 'currency', r'.*')
    assert not ok
    assert 'numeric' in reason


# ─── validate_type: date/datetime ────────────────────────────────────────────

def test_date_object():
    ok, _ = validate_type(datetime.date(2026, 1, 1), 'date', r'.*')
    assert ok

def test_datetime_object():
    ok, _ = validate_type(datetime.datetime(2026, 1, 1, 12, 0), 'date', r'.*')
    assert ok

def test_date_string_rejected():
    ok, reason = validate_type('2026-01-01', 'date', r'.*')
    assert not ok
    assert 'date' in reason


# ─── validate_type: percentage ───────────────────────────────────────────────

def test_percentage_float():
    ok, _ = validate_type(0.6246, 'percentage', r'.*')
    assert ok

def test_percentage_int():
    ok, _ = validate_type(1, 'percentage', r'.*')
    assert ok

def test_percentage_bool_rejected():
    ok, reason = validate_type(True, 'percentage', r'.*')
    assert not ok
    assert 'boolean' in reason

def test_percentage_string_rejected():
    ok, reason = validate_type('62%', 'percentage', r'.*')
    assert not ok
    assert 'numeric' in reason


# ─── validate_type: unknown type ─────────────────────────────────────────────

def test_unknown_type():
    ok, reason = validate_type('x', 'blob', r'.*')
    assert not ok
    assert 'unknown' in reason
