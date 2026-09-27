"""Unit tests for _match_lbl and _resolve_lbl_mode (engine internals)."""
import datetime

import pytest
from grepxcel.engine import _match_lbl, _resolve_lbl_mode, _validate_field
from grepxcel.models import Config, FieldDef


# ── _match_lbl: literal mode ─────────────────────────────────────────────────

class TestMatchLblLiteral:
    def test_exact_match(self):
        assert _match_lbl('Name', 'Name', 'literal', False)

    def test_no_match(self):
        assert not _match_lbl('Name', 'Age', 'literal', False)

    def test_case_sensitive_by_default(self):
        assert not _match_lbl('name', 'Name', 'literal', False)

    def test_ignore_case(self):
        assert _match_lbl('name', 'Name', 'literal', True)

    def test_parens_are_literal(self):
        assert _match_lbl('Term (months):', 'Term (months):', 'literal', False)

    def test_backslash_escaped_parens_do_not_match_plain(self):
        assert not _match_lbl('Term (months):', r'Term \(months\):', 'literal', False)

    def test_empty_pattern_always_matches(self):
        assert _match_lbl('anything', '', 'literal', False)

    def test_none_cell_value(self):
        assert _match_lbl(None, '', 'literal', False)

    def test_none_cell_nonempty_pattern(self):
        assert not _match_lbl(None, 'Name', 'literal', False)

    def test_multiline_cell_exact(self):
        assert _match_lbl('Breaks\n(minutes)', 'Breaks\n(minutes)', 'literal', False)


# ── _match_lbl: glob mode ────────────────────────────────────────────────────

class TestMatchLblGlob:
    def test_star_matches_any(self):
        assert _match_lbl('Invoice 2024-01', 'Invoice *', 'glob', False)

    def test_star_does_not_match_if_prefix_wrong(self):
        assert not _match_lbl('Receipt 2024-01', 'Invoice *', 'glob', False)

    def test_question_mark_matches_one_char(self):
        assert _match_lbl('PO-1', 'PO-?', 'glob', False)
        assert not _match_lbl('PO-12', 'PO-?', 'glob', False)

    def test_star_matches_newline(self):
        assert _match_lbl('Breaks\n(minutes)', 'Breaks*', 'glob', False)

    def test_ignore_case(self):
        assert _match_lbl('invoice 001', 'Invoice *', 'glob', True)

    def test_empty_pattern_always_matches(self):
        assert _match_lbl('x', '', 'glob', False)

    def test_exact_string_matches(self):
        assert _match_lbl('Total', 'Total', 'glob', False)


# ── _match_lbl: regexp mode ──────────────────────────────────────────────────

class TestMatchLblRegexp:
    def test_basic_regex(self):
        assert _match_lbl('PO-1234', r'PO-\d+', 'regexp', False)

    def test_escaped_parens_match_literal_paren(self):
        assert _match_lbl('Term (months):', r'Term \(months\):', 'regexp', False)

    def test_case_insensitive(self):
        assert _match_lbl('TOTAL', 'total', 'regexp', True)

    def test_empty_pattern_always_matches(self):
        assert _match_lbl('x', '', 'regexp', False)

    def test_partial_match(self):
        assert _match_lbl('see Invoice 1234 here', r'Invoice \d+', 'regexp', False)


# ── _resolve_lbl_mode ────────────────────────────────────────────────────────

class TestResolveLblMode:
    def _fd(self, lbl_match=None):
        return FieldDef(name='h', type='string', regex='.*', role='lbl',
                        lbl_match=lbl_match)

    def _cfg(self, lbl_match='literal'):
        return Config(lbl_match=lbl_match)

    def test_no_override_uses_global(self):
        assert _resolve_lbl_mode(self._fd(None), self._cfg('literal')) == 'literal'

    def test_per_field_override_wins(self):
        assert _resolve_lbl_mode(self._fd('glob'), self._cfg('literal')) == 'glob'

    def test_regexp_override_wins(self):
        assert _resolve_lbl_mode(self._fd('regexp'), self._cfg('glob')) == 'regexp'

    def test_global_regexp_no_override(self):
        assert _resolve_lbl_mode(self._fd(None), self._cfg('regexp')) == 'regexp'


# ── _validate_field: lbl: role against date/datetime cell values ────────────
# Regression for the gap where lbl: never got the datetime→isoformat
# conversion that var:literal/glob already has: str(datetime(2024,1,15))
# gives '2024-01-15 00:00:00', not '2024-01-15', which used to silently
# break literal/glob matches (and any $-anchored regexp) against a
# date-formatted label cell.

class TestValidateFieldLblDate:
    def _fd(self, mode, regex):
        return FieldDef(name='h', type='string', regex=regex, role='lbl', lbl_match=mode)

    def test_literal_matches_date_without_time_suffix(self):
        value = datetime.date(2024, 1, 15)
        fd = self._fd('literal', '2024-01-15')
        assert _validate_field(fd, value, Config(), 1000)

    def test_literal_matches_datetime_without_time_suffix(self):
        value = datetime.datetime(2024, 1, 15, 0, 0, 0)
        fd = self._fd('literal', '2024-01-15')
        assert _validate_field(fd, value, Config(), 1000)

    def test_literal_rejects_wrong_date(self):
        value = datetime.date(2024, 1, 15)
        fd = self._fd('literal', '2024-01-16')
        assert not _validate_field(fd, value, Config(), 1000)

    def test_glob_exact_pattern_matches_date(self):
        # Before the fix, an exact (no-wildcard) glob pattern failed the
        # same way literal did, for the same trailing-time-suffix reason.
        value = datetime.date(2024, 1, 15)
        fd = self._fd('glob', '2024-01-15')
        assert _validate_field(fd, value, Config(), 1000)

    def test_regexp_end_anchored_matches_date(self):
        # Before the fix, re.search(r'2024-01-15$', '2024-01-15 00:00:00')
        # failed — the $ anchor landed before the spurious time suffix.
        value = datetime.date(2024, 1, 15)
        fd = self._fd('regexp', r'2024-01-15$')
        assert _validate_field(fd, value, Config(), 1000)

    def test_datetime_with_nonzero_time_still_uses_date_only(self):
        # Matches var:literal/glob's own behaviour: only the date portion
        # is used for lbl: matching, regardless of the time-of-day part.
        value = datetime.datetime(2024, 1, 15, 13, 45, 0)
        fd = self._fd('literal', '2024-01-15')
        assert _validate_field(fd, value, Config(), 1000)

    def test_non_date_value_unaffected(self):
        fd = self._fd('literal', 'Active')
        assert _validate_field(fd, 'Active', Config(), 1000)
