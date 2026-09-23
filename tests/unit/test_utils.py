import datetime
import pytest
from grepxcel.utils import is_empty, validate_type


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

# ── is_empty: ignore_case alias matching ─────────────────────────────────────

def test_empty_alias_ignore_case_lower():
    assert is_empty('n/a', empty_aliases=['N/A'], ignore_case=True)

def test_empty_alias_ignore_case_mixed():
    assert is_empty('N/a', empty_aliases=['N/A'], ignore_case=True)

def test_empty_alias_ignore_case_still_requires_content():
    """ignore_case only affects comparison; a non-alias value is still not empty."""
    assert not is_empty('hello', empty_aliases=['N/A'], ignore_case=True)

def test_empty_alias_case_sensitive_no_match():
    """Without ignore_case, 'n/a' does NOT match alias 'N/A'."""
    assert not is_empty('n/a', empty_aliases=['N/A'], ignore_case=False)

def test_empty_alias_ignore_case_multiple_aliases():
    assert is_empty('tbd', empty_aliases=['N/A', 'TBD', '-'], ignore_case=True)

def test_empty_alias_with_surrounding_spaces():
    """Stripping happens before alias comparison, so '  N/A  ' matches alias 'N/A'."""
    assert is_empty('  N/A  ', empty_aliases=['N/A'])

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


# ─── validate_type: time (clock time and duration) ──────────────────────────

def test_time_clock_time():
    ok, _ = validate_type(datetime.time(9, 5), 'time', r'.*')
    assert ok

def test_time_timedelta_accepted():
    """Excel duration cells ([h]:mm format) → timedelta; 'time' must accept them."""
    ok, _ = validate_type(datetime.timedelta(hours=8), 'time', r'.*')
    assert ok

def test_time_timedelta_zero():
    ok, _ = validate_type(datetime.timedelta(0), 'time', r'.*')
    assert ok

def test_time_string_rejected():
    ok, reason = validate_type('08:00', 'time', r'.*')
    assert not ok
    assert 'time' in reason

def test_time_int_rejected():
    ok, reason = validate_type(480, 'time', r'.*')
    assert not ok
    assert 'time' in reason

def test_duration_type_alias():
    """'duration' is an explicit alias for timedelta-accepting time."""
    ok, _ = validate_type(datetime.timedelta(hours=1, minutes=30), 'duration', r'.*')
    assert ok

def test_duration_clock_time_accepted():
    ok, _ = validate_type(datetime.time(14, 30), 'duration', r'.*')
    assert ok

def test_duration_string_rejected():
    ok, reason = validate_type('1:30:00', 'duration', r'.*')
    assert not ok


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


# ─── validate_type: boolean ──────────────────────────────────────────────────

def test_bool_python_true():
    ok, _ = validate_type(True, 'boolean', r'.*')
    assert ok

def test_bool_python_false():
    ok, _ = validate_type(False, 'boolean', r'.*')
    assert ok

def test_bool_str_true_upper():
    ok, _ = validate_type('TRUE', 'boolean', r'.*')
    assert ok

def test_bool_str_false_lower():
    ok, _ = validate_type('false', 'boolean', r'.*')
    assert ok

def test_bool_str_yes():
    ok, _ = validate_type('YES', 'boolean', r'.*')
    assert ok

def test_bool_str_no_lower():
    ok, _ = validate_type('no', 'boolean', r'.*')
    assert ok

def test_bool_str_mixed_case():
    ok, _ = validate_type('  Yes  ', 'boolean', r'.*')
    assert ok

def test_bool_str_one():
    ok, _ = validate_type('1', 'boolean', r'.*')
    assert ok

def test_bool_str_zero():
    ok, _ = validate_type('0', 'boolean', r'.*')
    assert ok

def test_bool_int_one():
    ok, _ = validate_type(1, 'boolean', r'.*')
    assert ok

def test_bool_int_zero():
    ok, _ = validate_type(0, 'boolean', r'.*')
    assert ok

def test_bool_int_two_rejected():
    ok, reason = validate_type(2, 'boolean', r'.*')
    assert not ok
    assert 'boolean' in reason

def test_bool_string_arbitrary_rejected():
    ok, reason = validate_type('maybe', 'boolean', r'.*')
    assert not ok
    assert 'boolean' in reason

def test_bool_regex_applied():
    # Regex filters the string representation; "True" must match (?i)true|false
    ok, _ = validate_type(True, 'boolean', r'(?i)true|false')
    assert ok

def test_bool_regex_rejects_yes_when_restricted():
    # If the user restricts to true/false only via regex, YES should not match
    ok, _ = validate_type('YES', 'boolean', r'(?i)true|false')
    assert not ok

def test_bool_alias():
    # 'bool' is an alias for 'boolean'
    ok, _ = validate_type(True, 'bool', r'.*')
    assert ok


# ─── validate_type: url ──────────────────────────────────────────────────────

def test_url_https():
    ok, _ = validate_type('https://example.com', 'url', r'.*')
    assert ok

def test_url_http():
    ok, _ = validate_type('http://example.com/path?q=1', 'url', r'.*')
    assert ok

def test_url_ftp():
    ok, _ = validate_type('ftp://files.example.com/pub/file.zip', 'url', r'.*')
    assert ok

def test_url_ftps():
    ok, _ = validate_type('ftps://files.example.com/pub/file.zip', 'url', r'.*')
    assert ok

def test_url_mailto():
    ok, _ = validate_type('mailto:user@example.com', 'url', r'.*')
    assert ok

def test_url_tel():
    ok, _ = validate_type('tel:+1-800-555-0100', 'url', r'.*')
    assert ok

def test_url_sms():
    ok, _ = validate_type('sms:+15551234567', 'url', r'.*')
    assert ok

def test_url_file():
    ok, _ = validate_type('file:///home/user/document.xlsx', 'url', r'.*')
    assert ok

def test_url_no_scheme_rejected():
    ok, reason = validate_type('not-a-url', 'url', r'.*')
    assert not ok
    assert 'scheme' in reason or 'URL' in reason

def test_url_bare_www_rejected():
    # www.example.com has no scheme — rejected by urlparse-based check
    ok, reason = validate_type('www.example.com', 'url', r'.*')
    assert not ok
    assert 'scheme' in reason or 'URL' in reason

def test_url_javascript_blocked():
    ok, reason = validate_type('javascript:alert(1)', 'url', r'.*')
    assert not ok
    assert 'javascript' in reason

def test_url_unknown_scheme_rejected():
    ok, reason = validate_type('myapp://open?id=123', 'url', r'.*')
    assert not ok
    assert 'not supported' in reason or 'scheme' in reason

def test_url_regex_applied():
    ok, _ = validate_type('https://example.com', 'url', r'https://.*')
    assert ok

def test_url_regex_mismatch():
    ok, _ = validate_type('http://example.com', 'url', r'https://.*')
    assert not ok


# ─── validate_type: image ─────────────────────────────────────────────────────

def test_image_none_passes():
    # Image cells have no text value (openpyxl returns None); always valid.
    ok, reason = validate_type(None, 'image', r'.*')
    assert ok
    assert reason == ''

def test_image_any_value_passes():
    # validate_type for image never rejects based on cell value.
    ok, _ = validate_type('anything', 'image', r'.*')
    assert ok

def test_image_not_unknown_type():
    # 'image' must not fall through to the 'unknown type' error branch.
    ok, reason = validate_type(None, 'image', r'.*')
    assert 'unknown' not in reason


# ─── validate_type: unknown type ─────────────────────────────────────────────

def test_unknown_type():
    ok, reason = validate_type('x', 'blob', r'.*')
    assert not ok
    assert 'unknown' in reason


# ─── validate_type: ignore_case ──────────────────────────────────────────────

def test_ignore_case_off_is_default_sensitive():
    # Default behaviour: case matters.
    ok, _ = validate_type('paid', 'string', r'PAID')
    assert not ok

def test_ignore_case_string_match():
    ok, _ = validate_type('paid', 'string', r'PAID', ignore_case=True)
    assert ok

def test_ignore_case_string_class_match():
    ok, _ = validate_type('ACME', 'string', r'[a-z]+', ignore_case=True)
    assert ok

def test_ignore_case_still_respects_pattern():
    # Case-insensitive does not mean "match anything" — structure still applies.
    ok, _ = validate_type('paid123', 'string', r'PAID', ignore_case=True)
    assert not ok

def test_ignore_case_off_explicit():
    ok, _ = validate_type('paid', 'string', r'PAID', ignore_case=False)
    assert not ok


# ─── ReDoS: runtime timeout backstop ─────────────────────────────────────────

import time as _time
from grepxcel import utils as _utils
from grepxcel.security import check_regex_safety, SecurityError


def test_catastrophic_pattern_bypasses_static_guard_but_is_bounded():
    # (a|aa)+ has NO nested quantifier, so the static AST guard does not flag it…
    pattern = r'(a|aa)+$'
    check_regex_safety(pattern)  # must NOT raise — confirms it's a real bypass

    # …yet it is catastrophic on a non-matching tail. The hard timeout must
    # bound it: return quickly (no hang) with a no-match.
    evil_input = 'a' * 40 + 'X'
    start = _time.monotonic()
    ok = _utils._safe_match(pattern, evil_input, max_len=10_000)
    elapsed = _time.monotonic() - start

    assert ok is False
    assert elapsed < 5.0, f'match was not bounded: took {elapsed:.2f}s'


def test_timeout_is_treated_as_no_match(monkeypatch):
    # Deterministically exercise the timeout branch: force a TimeoutError and
    # confirm _safe_match swallows it and reports no-match (never propagates).
    def _boom(*args, **kwargs):
        raise TimeoutError('simulated catastrophic backtracking')

    monkeypatch.setattr(_utils._re, 'fullmatch', _boom)
    assert _utils._safe_match(r'.*', 'anything') is False


def test_regex_timeout_env_override(monkeypatch):
    monkeypatch.setenv('GREPXCEL_REGEX_TIMEOUT', '1.5')
    assert _utils._regex_timeout() == 1.5
    monkeypatch.setenv('GREPXCEL_REGEX_TIMEOUT', 'garbage')
    assert _utils._regex_timeout() == 0.25  # falls back to default
