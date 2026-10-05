"""The reason a value was rejected must reach the DEFAULT console output.

Found by installing the published 0.4.2rc42 from TestPyPI and running a real
extraction — not by any test. The oracle suite asserts the reason is in
`rec.message` and passed throughout, because `message` is an internal field:
`issue_line()` rendered `found X, expected matches /Y/` and never showed it. So
the recap block headed "ISSUES (cell — reason)" printed no reason, and only `-v`
revealed the cause.

That is the lesson the suite was built on, turned on the suite itself: assert
what renders, not the field behind it. These tests assert the rendered line.

Also covers the two guard cases that motivated the whole exercise. A cell over
--max-cell-len and a regex timeout are BOTH reported by `_safe_match` as a plain
False, which read as "does not match /pattern/" — pointing the user at a pattern
that was never run. They now say so.
"""
import datetime

import pytest

from grepxcel.logger import Category, Logger, LogRecord, Severity, VerbosityLevel
from grepxcel.utils import _clip, _match_with_reason, _MAX_REGEX_INPUT_LEN, validate_type


class TestIssueLineShowsTheReason:
    def _line(self, **kw):
        lg = Logger(level=VerbosityLevel.QUIET)
        return lg.issue_line(LogRecord(
            Severity.WARNING, Category.VALIDATION, 'Value rejected: x',
            location='Sheet1!B1', field='f', **kw))

    def test_the_reason_is_rendered(self):
        line = self._line(reason='42.5 is not a whole number',
                          found='42.5', expected='matches /.*/')
        assert 'is not a whole number' in line

    def test_the_reason_leads_rather_than_the_regex(self):
        """'expected matches /.*/' names a pattern that accepts everything; it
        is noise when the value failed a type check."""
        line = self._line(reason='boolean is not integer',
                          found='True', expected='matches /.*/')
        assert line.index('boolean is not integer') < line.index('True')
        assert 'expected matches' not in line

    def test_found_is_kept_for_context(self):
        line = self._line(reason='boolean is not integer', found='True')
        assert 'True' in line

    def test_without_a_reason_the_old_format_is_unchanged(self):
        """Records that carry no reason (empty field, undefined field, table
        bounds) must render exactly as before."""
        line = self._line(found='3', expected='2 rows')
        assert 'found 3, expected 2 rows' in line

    def test_a_record_with_neither_falls_back_to_the_message(self):
        lg = Logger(level=VerbosityLevel.QUIET)
        line = lg.issue_line(LogRecord(
            Severity.ERROR, Category.STRUCTURAL, 'sheet is exhausted'))
        assert 'sheet is exhausted' in line


class TestGuardsNameThemselves:
    """The two cases the spec cites as the reason this suite exists."""

    def test_over_max_cell_len_says_so(self):
        long_value = 'x' * (_MAX_REGEX_INPUT_LEN + 500)
        ok, reason = validate_type(long_value, 'string', '.*')
        assert not ok
        assert 'over the' in reason and 'limit' in reason
        assert 'never evaluated' in reason
        assert '--max-cell-len' in reason, 'the fix must be named'

    def test_over_max_cell_len_does_not_claim_a_pattern_mismatch(self):
        """'.*' matches anything, so "does not match /.*/" is actively wrong."""
        ok, reason = validate_type('x' * 2000, 'string', '.*')
        assert not ok
        assert 'does not match' not in reason

    def test_a_value_at_exactly_the_limit_is_still_checked(self):
        """N-1, N, N+1 around the constant: the guard must fire at N+1 only."""
        assert validate_type('x' * (_MAX_REGEX_INPUT_LEN - 1), 'string', '.*')[0]
        assert validate_type('x' * _MAX_REGEX_INPUT_LEN, 'string', '.*')[0]
        assert not validate_type('x' * (_MAX_REGEX_INPUT_LEN + 1), 'string', '.*')[0]

    def test_a_genuine_mismatch_still_says_so(self):
        ok, reason = validate_type('abc', 'string', r'\d+')
        assert not ok
        assert 'does not match' in reason


class TestLongValuesAreClipped:
    def test_clip_shortens_and_reports_the_real_length(self):
        out = _clip('y' * 5000)
        assert len(out) < 120
        assert '5000 chars' in out

    def test_clip_leaves_a_short_value_alone(self):
        assert _clip('abc') == "'abc'"

    def test_the_rendered_line_stays_readable(self):
        """A 1500-char cell used to print in full, burying the explanation."""
        lg = Logger(level=VerbosityLevel.QUIET)
        rec = lg.warn_validation(1, 2, 'f', 'string', '.*', 'z' * 1500,
                                 reason='value is 1500 characters, over the '
                                        '1000-character limit')
        line = lg.issue_line(rec)
        assert len(line) < 250, f'issue line is {len(line)} chars'
        assert 'over the' in line

    def test_the_full_length_is_still_recorded(self):
        """Clipping the display must not lose the forensic detail."""
        lg = Logger(level=VerbosityLevel.QUIET)
        rec = lg.warn_validation(1, 2, 'f', 'string', '.*', 'z' * 1500)
        assert rec.value_len == 1500
        assert rec.value_sha8


class TestIntegerHintIsAccurate:
    def test_a_fractional_value_is_not_called_an_integer(self):
        """The hint said "The integer 42 does not satisfy /.*/" for 42.5 —
        it renamed the value and blamed a pattern that matches everything."""
        lg = Logger(level=VerbosityLevel.QUIET)
        rec = lg.warn_validation(1, 2, 'f', 'integer', '.*', 42.5,
                                 reason='42.5 is not a whole number')
        assert 'The integer 42 ' not in rec.hint
        assert 'fractional' in rec.hint or 'not an integer' in rec.hint

    def test_a_whole_value_still_gets_the_regex_hint(self):
        lg = Logger(level=VerbosityLevel.QUIET)
        rec = lg.warn_validation(1, 2, 'f', 'integer', r'^\d{5}$', 42,
                                 reason='42 does not match /^\\d{5}$/')
        assert 'digit count' in rec.hint


class TestStructuredLogsStayClean:
    def test_reason_is_not_allow_listed(self):
        """reason contains the cell value, so it must never reach an NDJSON log
        — the guarantee is an allow-list, and this is the test that keeps it."""
        from grepxcel.logger import _SAFE_LOG_KEYS, _SAFE_SUMMARY_KEYS
        assert 'reason' not in _SAFE_LOG_KEYS
        assert 'reason' not in _SAFE_SUMMARY_KEYS

    def test_a_json_record_carries_no_reason(self, tmp_path):
        log = str(tmp_path / 'log.ndjson')
        lg = Logger(level=VerbosityLevel.QUIET, log_file=log, log_format='json')
        rec = lg.warn_validation(1, 2, 'f', 'integer', '.*', 'SECRET-VALUE',
                                 reason="'SECRET-VALUE' is not integer type")
        lg.commit_warnings([rec])
        lg.close()
        written = open(log, encoding='utf-8').read()
        assert 'SECRET-VALUE' not in written
        assert 'reason' not in written


@pytest.mark.parametrize('flags,text,expected_ok', [
    (0, 'abc', True),
    (0, 'abd', False),
])
def test_match_with_reason_basic(flags, text, expected_ok):
    ok, reason = _match_with_reason('abc', text, flags, 1000)
    assert ok is expected_ok
    assert (reason == '') is expected_ok
