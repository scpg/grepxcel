"""Unit tests for _match_lbl and _resolve_lbl_mode (engine internals)."""
import pytest
from grepxcel.engine import _match_lbl, _resolve_lbl_mode
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
