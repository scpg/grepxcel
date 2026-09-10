"""Unit tests for the col-A modifier system (not-null, var:glob, var:literal, var:re,
trim-whitespace).

These tests verify:
1. Parser: modifier tokens produce the correct FieldDef attributes.
2. Engine: _validate_field dispatches correctly for each var_mode.
3. Engine: required (not-null) check fires a fatal error when the value is empty.
4. Engine: trim-whitespace strips leading/trailing spaces before matching and extraction.
"""
from pathlib import Path

import openpyxl
import pytest

from grepxcel.engine import _validate_field
from grepxcel.models import Config, FieldDef
from grepxcel.pattern_parser import PatternParser, PatternError


# ── Helpers ───────────────────────────────────────────────────────────────────

def _write_pattern(rows, tmp_path: Path) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / 'pattern.xlsx')
    wb.save(path)
    return path


def _field(role='var', regex='.*', type_='string', var_mode=None, required=False,
           lbl_match=None):
    return FieldDef(name='f', type=type_, regex=regex, role=role,
                    var_mode=var_mode, required=required, lbl_match=lbl_match)


_CFG = Config()
_MAX_LEN = 1000


# ── _validate_field: var: default (regexp) ────────────────────────────────────

class TestValidateFieldVarRegexp:
    def test_string_regex_match(self):
        fd = _field(regex=r'[A-Z]{3}')
        assert _validate_field(fd, 'ABC', _CFG, _MAX_LEN)

    def test_string_regex_no_match(self):
        fd = _field(regex=r'[A-Z]{3}')
        assert not _validate_field(fd, '123', _CFG, _MAX_LEN)

    def test_currency_type_match(self):
        fd = _field(type_='currency', regex=r'.*')
        assert _validate_field(fd, 42.5, _CFG, _MAX_LEN)

    def test_currency_type_wrong(self):
        fd = _field(type_='currency', regex=r'.*')
        assert not _validate_field(fd, 'text', _CFG, _MAX_LEN)

    def test_explicit_regexp_mode_same_as_default(self):
        fd = _field(regex=r'\d+', var_mode='regexp')
        assert _validate_field(fd, '42', _CFG, _MAX_LEN)


# ── _validate_field: var:glob ─────────────────────────────────────────────────

class TestValidateFieldVarGlob:
    def test_glob_match(self):
        fd = _field(regex='PROD-*', var_mode='glob')
        assert _validate_field(fd, 'PROD-12345', _CFG, _MAX_LEN)

    def test_glob_no_match(self):
        fd = _field(regex='PROD-*', var_mode='glob')
        assert not _validate_field(fd, 'SKU-12345', _CFG, _MAX_LEN)

    def test_glob_question_mark(self):
        fd = _field(regex='A?C', var_mode='glob')
        assert _validate_field(fd, 'ABC', _CFG, _MAX_LEN)
        assert not _validate_field(fd, 'ABBC', _CFG, _MAX_LEN)

    def test_glob_type_check_fails(self):
        """Type check still runs even in glob mode."""
        fd = _field(type_='currency', regex='1*', var_mode='glob')
        # 'text' is not a currency → type_ok = False → overall False
        assert not _validate_field(fd, 'text', _CFG, _MAX_LEN)

    def test_glob_type_check_passes_then_glob_matches(self):
        fd = _field(type_='currency', regex='1*', var_mode='glob')
        assert _validate_field(fd, 100.0, _CFG, _MAX_LEN)

    def test_glob_type_check_passes_glob_fails(self):
        fd = _field(type_='currency', regex='1*', var_mode='glob')
        assert not _validate_field(fd, 200.0, _CFG, _MAX_LEN)

    def test_glob_ignore_case(self):
        cfg = Config(ignore_case=True)
        fd = _field(regex='prod-*', var_mode='glob')
        assert _validate_field(fd, 'PROD-99', cfg, _MAX_LEN)


# ── _validate_field: var:literal ─────────────────────────────────────────────

class TestValidateFieldVarLiteral:
    def test_exact_match(self):
        fd = _field(regex='Active', var_mode='literal')
        assert _validate_field(fd, 'Active', _CFG, _MAX_LEN)

    def test_no_match(self):
        fd = _field(regex='Active', var_mode='literal')
        assert not _validate_field(fd, 'Inactive', _CFG, _MAX_LEN)

    def test_case_sensitive(self):
        fd = _field(regex='Active', var_mode='literal')
        assert not _validate_field(fd, 'active', _CFG, _MAX_LEN)

    def test_ignore_case(self):
        cfg = Config(ignore_case=True)
        fd = _field(regex='Active', var_mode='literal')
        assert _validate_field(fd, 'ACTIVE', cfg, _MAX_LEN)

    def test_parens_and_special_chars_are_literal(self):
        fd = _field(regex='Term (months):', var_mode='literal')
        assert _validate_field(fd, 'Term (months):', _CFG, _MAX_LEN)
        assert not _validate_field(fd, 'Term x months:', _CFG, _MAX_LEN)


# ── required (not-null) enforcement via engine ────────────────────────────────
# We test the full extraction path using grepxcel.extract() to exercise the
# fatal-error code path inside _process_cell.

from grepxcel.engine import Engine
from grepxcel.logger import Logger, VerbosityLevel


def _write_data(cells: dict, tmp_path: Path, name: str = 'data.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for coord, val in cells.items():
        ws[coord] = val
    path = str(tmp_path / name)
    wb.save(path)
    return path


def _run(pattern_rows, data_cells, tmp_path: Path) -> tuple:
    """Run extraction; return (result_dict, logger).  Uses cell:A1 throughout."""
    pat = _write_pattern(pattern_rows, tmp_path)
    dat = _write_data(data_cells, tmp_path)
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pat, dat, logger=lg)
    return result, lg


def _req_rows(col_a: str, regex: str = '.*') -> list:
    return [
        [col_a, 'v', 'string', regex],
        ['START:'],
        ['cell:A1', 'v'],
        ['END:'],
    ]


class TestRequiredEnforcement:
    def test_not_null_passes_when_value_present(self, tmp_path):
        result, lg = _run(_req_rows('var:not-null'), {'A1': 'hello'}, tmp_path)
        assert not lg.has_errors()
        assert result.get('v') == 'hello'

    def test_not_null_logs_fatal_when_value_empty(self, tmp_path):
        """Required field with empty cell → fatal error logged."""
        _, lg = _run(_req_rows('var:not-null'), {}, tmp_path)
        assert lg.has_errors()

    def test_not_empty_synonym_logs_fatal(self, tmp_path):
        _, lg = _run(_req_rows('var:not-empty'), {}, tmp_path)
        assert lg.has_errors()

    def test_plain_var_allows_empty_without_error(self, tmp_path):
        """Backward compat: plain var: with empty cell → null, no fatal."""
        result, lg = _run(_req_rows('var:'), {}, tmp_path)
        assert not lg.has_errors()
        assert result.get('v') is None

    def test_not_null_with_glob_passes(self, tmp_path):
        result, lg = _run(_req_rows('var:not-null:glob', 'hello*'),
                           {'A1': 'hello world'}, tmp_path)
        assert not lg.has_errors()
        assert result.get('v') == 'hello world'

    def test_not_null_with_glob_logs_fatal_on_empty(self, tmp_path):
        _, lg = _run(_req_rows('var:not-null:glob', 'hello*'), {}, tmp_path)
        assert lg.has_errors()


# ── trim-whitespace modifier ──────────────────────────────────────────────────

class TestTrimWhitespaceParser:
    """Parser correctly sets trim_whitespace on FieldDef and Config."""

    def test_var_trim_sets_field_flag(self, tmp_path):
        path = _write_pattern([
            ['var:trim-whitespace', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['v'].trim_whitespace is True

    def test_plain_var_trim_whitespace_is_false(self, tmp_path):
        path = _write_pattern([
            ['var:', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['v'].trim_whitespace is False

    def test_lbl_trim_sets_field_flag(self, tmp_path):
        path = _write_pattern([
            ['lbl:trim-whitespace', 'lbl', 'string', 'Invoice:'],
            ['var:', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'lbl'],
            ['cell:next', 'v'],
            ['END:'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['lbl'].trim_whitespace is True

    def test_trim_whitespace_combines_with_not_null(self, tmp_path):
        path = _write_pattern([
            ['var:not-null:trim-whitespace', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['v'].trim_whitespace is True
        assert defs['v'].required is True

    def test_trim_whitespace_combines_with_glob(self, tmp_path):
        path = _write_pattern([
            ['var:glob:trim-whitespace', 'v', 'string', 'hello*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['v'].trim_whitespace is True
        assert defs['v'].var_mode == 'glob'

    def test_global_config_trim_whitespace(self, tmp_path):
        path = _write_pattern([
            ['config:', 'trim.whitespace', 'yes'],
            ['var:', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert cfg.trim_whitespace is True

    def test_global_config_trim_whitespace_false(self, tmp_path):
        path = _write_pattern([
            ['config:', 'trim.whitespace', 'no'],
            ['var:', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert cfg.trim_whitespace is False


class TestTrimWhitespaceEngine:
    """Engine strips whitespace before matching and stores the trimmed value."""

    def _trim_rows(self, col_a='var:trim-whitespace', regex='.*') -> list:
        return [
            [col_a, 'v', 'string', regex],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ]

    def test_no_trim_by_default_whitespace_stored_raw(self, tmp_path):
        """Without trim-whitespace, raw value (with spaces) is stored."""
        result, lg = _run(
            [['var:', 'v', 'string', r'\S+'], ['START:'], ['cell:A1', 'v'], ['END:']],
            {'A1': ' hello '},
            tmp_path,
        )
        # '\S+' does NOT match ' hello ' — warning expected, raw value stored
        assert result.get('v') == ' hello '
        assert lg.has_warnings()  # validation warning because of spaces

    def test_trim_removes_leading_trailing_spaces(self, tmp_path):
        result, lg = _run(self._trim_rows(), {'A1': '  hello  '}, tmp_path)
        assert not lg.has_errors()
        assert result.get('v') == 'hello'

    def test_trim_enables_regex_match_that_would_fail_raw(self, tmp_path):
        result, lg = _run(self._trim_rows(regex=r'\w+'), {'A1': ' world '}, tmp_path)
        assert not lg.has_errors()
        assert result.get('v') == 'world'

    def test_trim_glob_mode(self, tmp_path):
        result, lg = _run(
            self._trim_rows(col_a='var:glob:trim-whitespace', regex='hello*'),
            {'A1': '  hello world  '},
            tmp_path,
        )
        assert not lg.has_errors()
        assert result.get('v') == 'hello world'

    def test_trim_literal_mode(self, tmp_path):
        result, lg = _run(
            self._trim_rows(col_a='var:literal:trim-whitespace', regex='Active'),
            {'A1': '  Active  '},
            tmp_path,
        )
        assert not lg.has_errors()
        assert result.get('v') == 'Active'

    def test_global_trim_config_applies_to_all_fields(self, tmp_path):
        """config: | trim.whitespace | yes trims all fields without per-field modifier."""
        rows = [
            ['config:', 'trim.whitespace', 'yes'],
            ['var:', 'v', 'string', r'\w+'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ]
        result, lg = _run(rows, {'A1': '  trimmed  '}, tmp_path)
        assert not lg.has_errors()
        assert result.get('v') == 'trimmed'

    def test_trim_does_not_affect_non_string_types(self, tmp_path):
        """Numeric values are Python objects — trim has no effect."""
        result, lg = _run(
            self._trim_rows(col_a='var:trim-whitespace', regex='.*'),
            {'A1': 42.5},
            tmp_path,
        )
        assert not lg.has_errors()
        assert result.get('v') == 42.5  # numeric value unchanged


class TestTrimWhitespaceWarning:
    """Logger emits a targeted hint when whitespace causes a validation mismatch."""

    def test_whitespace_hint_shown_when_trimming_fixes_mismatch(self, tmp_path):
        """When the raw value fails but trimmed value would pass, the warning says so."""
        rows = [
            ['var:', 'v', 'string', r'\w+'],  # no trim-whitespace
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ]
        result, lg = _run(rows, {'A1': ' hello '}, tmp_path)
        issues = lg.issues()
        assert issues  # a validation warning must have fired
        # The hint in the warning record must mention whitespace / trim-whitespace
        hints = [i.hint for i in issues if hasattr(i, 'hint') and i.hint]
        assert any('trim-whitespace' in h for h in hints), (
            f'Expected whitespace hint in warnings; got hints: {hints}'
        )
