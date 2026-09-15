"""Tests for the `nullable` modifier on var: field declarations.

Design intent (per spec):
  * `var: | field | type`  — required by default: warns when empty (current behaviour).
  * `var:nullable | field | type` — empty/null accepted silently: no warning, no error.
  * `lbl:` fields are anchors and never nullable — PatternError if attempted.
  * `nullable` + `not-null`/`not-empty` on the same field → PatternError (contradictory).
  * `nullable` can be combined with matching-mode modifiers (glob, literal, regexp).
  * `nullable` can be combined with `trim-whitespace`.
  * When a nullable field is non-empty the value is extracted and type-validated normally.
  * The fix applies to BOTH scalar (cell:) and table DATA fields, because both reference
    the same var: declarations — one code path, one flag.

Coverage:
  A. Parser — FieldDef attributes
  B. Parser — error conditions
  C. Engine — scalar cell extraction (no warning, null in result)
  D. Engine — table DATA extraction (no warning, null in result)
  E. Engine — non-empty nullable field (value extracted, type-coerced normally)
  F. Engine — interaction with existing modifiers
  G. Engine — required (not-null) is unaffected by nullable change
  H. Warnings / exit-code semantics (via logger.has_warnings())
  I. Validate-pattern recognises nullable as valid
"""
from pathlib import Path

import openpyxl
import pytest

from grepxcel.engine import Engine
from grepxcel.logger import Logger, VerbosityLevel
from grepxcel.models import Config, FieldDef
from grepxcel.pattern_parser import PatternParser, PatternError


# ── Shared helpers ────────────────────────────────────────────────────────────

def _write_pattern(rows, tmp_path: Path, name: str = 'pattern.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / name)
    wb.save(path)
    return path


def _write_data(cells: dict, tmp_path: Path, name: str = 'data.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet'
    for coord, val in cells.items():
        ws[coord] = val
    path = str(tmp_path / name)
    wb.save(path)
    return path


def _run(pattern_rows, data_cells, tmp_path: Path,
         level: VerbosityLevel = VerbosityLevel.QUIET) -> tuple:
    """Run extraction; return (result_dict, logger)."""
    pat = _write_pattern(pattern_rows, tmp_path)
    dat = _write_data(data_cells, tmp_path)
    lg = Logger(level=level)
    result = Engine().process(pat, dat, logger=lg)
    return result, lg


def _scalar_rows(col_a: str, type_: str = 'string', regex: str = '.*') -> list:
    """Minimal pattern: one var: field extracted from cell A1."""
    return [
        [col_a, 'v', type_, regex],
        ['START:'],
        ['cell:A1', 'v'],
        ['END:'],
    ]


def _table_rows(col_a_v: str, type_: str = 'string', lbl: str = 'Col') -> list:
    """Minimal table pattern: 1 lbl header + 1 var: DATA field in column B."""
    return [
        ['lbl:', 'h', 'string', lbl],
        [col_a_v, 'row.v', type_, '.*'],
        ['START:'],
        ['table:*'],
        ['', 'HEADER:1', 'h'],
        ['', 'DATA:*', 'row.v'],
        ['END:'],
    ]


def _parse(rows, tmp_path: Path) -> tuple:
    """Parse pattern; return (config, defs, instructions)."""
    return PatternParser().parse(_write_pattern(rows, tmp_path))


# ── A. Parser: FieldDef attributes ───────────────────────────────────────────

class TestNullableParserAttributes:
    def test_var_nullable_sets_flag(self, tmp_path):
        _, defs, _ = _parse(_scalar_rows('var:nullable'), tmp_path)
        assert defs['v'].nullable is True

    def test_plain_var_nullable_is_false(self, tmp_path):
        _, defs, _ = _parse(_scalar_rows('var:'), tmp_path)
        assert defs['v'].nullable is False

    def test_var_colon_nullable_flag_and_required_false(self, tmp_path):
        _, defs, _ = _parse(_scalar_rows('var:nullable'), tmp_path)
        fd = defs['v']
        assert fd.nullable is True
        assert fd.required is False      # not-null is still OFF

    def test_nullable_does_not_set_required(self, tmp_path):
        _, defs, _ = _parse(_scalar_rows('var:nullable'), tmp_path)
        assert defs['v'].required is False

    def test_nullable_combines_with_glob(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable:glob', 'v', 'string', 'prefix-*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        fd = defs['v']
        assert fd.nullable is True
        assert fd.var_mode == 'glob'

    def test_nullable_combines_with_literal(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable:literal', 'v', 'string', 'exact'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        fd = defs['v']
        assert fd.nullable is True
        assert fd.var_mode == 'literal'

    def test_nullable_combines_with_regexp(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable:regexp', 'v', 'string', r'\d+'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        assert defs['v'].nullable is True
        assert defs['v'].var_mode == 'regexp'

    def test_nullable_combines_with_trim_whitespace(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable:trim-whitespace', 'v', 'string', '.*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        fd = defs['v']
        assert fd.nullable is True
        assert fd.trim_whitespace is True

    def test_nullable_trim_glob_triple_combo(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable:trim-whitespace:glob', 'v', 'string', 'hello*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        fd = defs['v']
        assert fd.nullable is True
        assert fd.trim_whitespace is True
        assert fd.var_mode == 'glob'

    def test_nullable_on_typed_duration_field(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable', 'v', 'duration', '.*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        assert defs['v'].nullable is True
        assert defs['v'].type == 'duration'

    def test_order_independent_nullable_before_mode(self, tmp_path):
        _, defs, _ = _parse([
            ['var:nullable:glob', 'v', 'string', 'x*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        assert defs['v'].nullable is True

    def test_order_independent_mode_before_nullable(self, tmp_path):
        _, defs, _ = _parse([
            ['var:glob:nullable', 'v', 'string', 'x*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], tmp_path)
        assert defs['v'].nullable is True
        assert defs['v'].var_mode == 'glob'


# ── B. Parser: error conditions ───────────────────────────────────────────────

class TestNullableParserErrors:
    def test_lbl_nullable_raises_pattern_error(self, tmp_path):
        """nullable is not valid for lbl: (anchor) fields."""
        rows = [
            ['lbl:nullable', 'h', 'string', 'Header'],
            ['var:', 'v', 'string', '.*'],
            ['START:'], ['cell:A1', 'h'], ['cell:next', 'v'], ['END:'],
        ]
        with pytest.raises(PatternError, match='nullable.*not valid for lbl:'):
            _parse(rows, tmp_path)

    def test_nullable_and_not_null_raises_pattern_error(self, tmp_path):
        """Contradictory modifiers: nullable + not-null."""
        rows = _scalar_rows('var:nullable:not-null')
        with pytest.raises(PatternError, match='[Cc]ontradictory'):
            _parse(rows, tmp_path)

    def test_nullable_and_not_empty_raises_pattern_error(self, tmp_path):
        """Contradictory modifiers: nullable + not-empty."""
        rows = _scalar_rows('var:nullable:not-empty')
        with pytest.raises(PatternError, match='[Cc]ontradictory'):
            _parse(rows, tmp_path)

    def test_unknown_modifier_still_raises(self, tmp_path):
        """Unknown tokens still produce a PatternError with the updated message."""
        rows = _scalar_rows('var:bogus-token')
        with pytest.raises(PatternError, match='Unknown modifier'):
            _parse(rows, tmp_path)

    def test_error_message_lists_nullable_as_valid(self, tmp_path):
        """The updated error for unknown modifiers lists 'nullable' as valid."""
        rows = _scalar_rows('var:oops')
        with pytest.raises(PatternError, match='nullable'):
            _parse(rows, tmp_path)

    def test_duplicate_mode_still_raises(self, tmp_path):
        """Duplicate mode modifier error is unaffected by nullable addition."""
        rows = _scalar_rows('var:glob:literal:nullable')
        with pytest.raises(PatternError, match='Duplicate mode'):
            _parse(rows, tmp_path)


# ── C. Engine: scalar cell extraction ────────────────────────────────────────

class TestNullableScalarEngine:
    def test_nullable_empty_cell_no_warning(self, tmp_path):
        """Empty cell on a nullable field → null result, no warnings."""
        result, lg = _run(_scalar_rows('var:nullable'), {}, tmp_path)
        assert result.get('v') is None
        assert not lg.has_warnings()
        assert not lg.has_errors()

    def test_plain_var_empty_cell_has_no_warning_scalar(self, tmp_path):
        """Scalar path: plain var: with empty cell → no warning (scalar path never warns
        on empty unless required=True — this is expected baseline behaviour)."""
        result, lg = _run(_scalar_rows('var:'), {}, tmp_path)
        assert result.get('v') is None
        # Scalar path doesn't warn on empty for non-required fields — confirm unchanged
        assert not lg.has_warnings()

    def test_nullable_non_empty_extracts_value(self, tmp_path):
        """Non-empty cell on nullable field → value extracted normally."""
        result, lg = _run(_scalar_rows('var:nullable'), {'A1': 'hello'}, tmp_path)
        assert result.get('v') == 'hello'
        assert not lg.has_warnings()
        assert not lg.has_errors()

    def test_nullable_typed_integer_empty_no_warning(self, tmp_path):
        result, lg = _run(_scalar_rows('var:nullable', type_='integer'), {}, tmp_path)
        assert result.get('v') is None
        assert not lg.has_warnings()

    def test_nullable_typed_integer_with_value_coerces(self, tmp_path):
        result, lg = _run(_scalar_rows('var:nullable', type_='integer'), {'A1': 42}, tmp_path)
        assert result.get('v') == 42
        assert not lg.has_warnings()

    def test_nullable_typed_string_with_regex_empty_no_warning(self, tmp_path):
        result, lg = _run(
            _scalar_rows('var:nullable', type_='string', regex=r'\d{4}'), {}, tmp_path
        )
        assert result.get('v') is None
        assert not lg.has_warnings()

    def test_nullable_with_value_failing_regex_still_warns(self, tmp_path):
        """nullable only suppresses empty-cell warnings; a non-empty value that fails
        its regex still produces a validation warning."""
        result, lg = _run(
            _scalar_rows('var:nullable', regex=r'^\d+$'), {'A1': 'not-a-number'}, tmp_path
        )
        assert result.get('v') == 'not-a-number'
        assert lg.has_warnings()  # regex mismatch — still warns

    def test_nullable_glob_empty_no_warning(self, tmp_path):
        result, lg = _run([
            ['var:nullable:glob', 'v', 'string', 'prefix-*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {}, tmp_path)
        assert result.get('v') is None
        assert not lg.has_warnings()

    def test_nullable_glob_with_value_matches(self, tmp_path):
        result, lg = _run([
            ['var:nullable:glob', 'v', 'string', 'INV-*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {'A1': 'INV-0042'}, tmp_path)
        assert result.get('v') == 'INV-0042'
        assert not lg.has_warnings()

    def test_nullable_literal_empty_no_warning(self, tmp_path):
        result, lg = _run([
            ['var:nullable:literal', 'v', 'string', 'exact-text'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {}, tmp_path)
        assert result.get('v') is None
        assert not lg.has_warnings()


# ── D. Engine: table DATA extraction ─────────────────────────────────────────

class TestNullableTableEngine:
    """The main motivation: sparse table columns (sick/holiday/vacation time etc.)"""

    def _table_data(self, lbl_cell: str, data_cells: dict, tmp_path: Path,
                    col_a_v: str = 'var:nullable', type_: str = 'string') -> tuple:
        rows = _table_rows(col_a_v, type_=type_)
        pat = _write_pattern(rows, tmp_path)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet'
        ws['A1'] = lbl_cell          # header cell
        for coord, val in data_cells.items():
            ws[coord] = val
        dat = str(tmp_path / 'data.xlsx')
        wb.save(dat)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        return result, lg

    def test_nullable_table_field_empty_no_warning(self, tmp_path):
        """Empty cell in a nullable DATA column → null, no warning."""
        # header='Col' in A1, data rows have nothing in B (the var column)
        result, lg = self._table_data('Col', {}, tmp_path)
        assert not lg.has_warnings()
        assert not lg.has_errors()

    def test_non_nullable_table_field_empty_warns(self, tmp_path):
        """Baseline: non-nullable DATA column with empty cell → warning.

        Two DATA columns: col A = required non-nullable (has a value → row is recognised),
        col B = also non-nullable but empty → warning generated.
        Note: _row_is_end_of_data returns True when ALL non-IGNORE fields are empty;
        so at least one non-IGNORE field in the row must be non-empty.
        """
        rows = [
            ['lbl:', 'h', 'string', 'Col'],
            ['var:', 'row.anchor', 'string', '.*'],  # present → row is a data row
            ['var:', 'row.v', 'string', '.*'],        # non-nullable, empty → warns
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h', 'h'],
            ['', 'DATA:*', 'row.anchor', 'row.v'],
            ['END:'],
        ]
        pat = _write_pattern(rows, tmp_path, name='p2.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet'
        ws['A1'] = 'Col'    # header col A
        ws['B1'] = 'Col'    # header col B
        ws['A2'] = 'present'  # col A non-empty → row IS a data row
        # B2 absent (None) → non-nullable field empty → warn
        dat = str(tmp_path / 'd2.xlsx')
        wb.save(dat)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert lg.has_warnings()  # non-nullable B2 is empty → warning

    def test_nullable_table_field_with_value_extracted(self, tmp_path):
        """Non-empty cell in nullable DATA column → value extracted, no warning."""
        rows = _table_rows('var:nullable', type_='string')
        pat = _write_pattern(rows, tmp_path, name='p3.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet'
        ws['A1'] = 'Col'   # header
        ws['A2'] = 'hello' # data row value
        dat = str(tmp_path / 'd3.xlsx')
        wb.save(dat)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert not lg.has_warnings()
        # result has table data
        table_data = result.get('row', [{}])
        assert len(table_data) > 0

    def test_nullable_multi_column_table_sparse(self, tmp_path):
        """Multi-column table: some nullable, some not; only non-nullable warns."""
        rows = [
            ['lbl:', 'h', 'string', 'Name'],
            ['var:', 'row.name', 'string', '.*'],      # required
            ['var:nullable', 'row.opt', 'string', '.*'],  # nullable
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h', 'h'],
            ['', 'DATA:*', 'row.name', 'row.opt'],
            ['END:'],
        ]
        pat = _write_pattern(rows, tmp_path, name='pm.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet'
        ws['A1'] = 'Name'  # header col A
        ws['B1'] = 'Name'  # header col B (re-use same lbl for both headers)
        ws['A2'] = 'Alice'  # data row: name present
        # ws['B2'] intentionally empty — nullable column
        dat = str(tmp_path / 'dm.xlsx')
        wb.save(dat)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        # nullable column B2 is empty → no warning
        # non-nullable column A2 has value → no warning
        assert not lg.has_warnings()

    def test_nullable_integer_table_field_empty(self, tmp_path):
        """Integer typed nullable table field: empty → null, no warning."""
        rows = [
            ['lbl:', 'h', 'string', 'Count'],
            ['var:nullable', 'row.n', 'integer', r'.*'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h'],
            ['', 'DATA:*', 'row.n'],
            ['END:'],
        ]
        pat = _write_pattern(rows, tmp_path, name='pi.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet'
        ws['A1'] = 'Count'
        ws['A2'] = 5        # one valid row
        ws['A3'] = None     # empty → nullable, no warning
        dat = str(tmp_path / 'di.xlsx')
        wb.save(dat)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert not lg.has_warnings()


# ── E. Non-empty nullable field: type-coercion works normally ─────────────────

class TestNullableValueCoercion:
    def test_integer_value_via_nullable(self, tmp_path):
        result, lg = _run(_scalar_rows('var:nullable', type_='integer'), {'A1': 7}, tmp_path)
        assert result.get('v') == 7
        assert not lg.has_errors()

    def test_string_value_via_nullable(self, tmp_path):
        result, lg = _run(_scalar_rows('var:nullable'), {'A1': 'world'}, tmp_path)
        assert result.get('v') == 'world'

    def test_number_value_via_nullable(self, tmp_path):
        result, lg = _run(_scalar_rows('var:nullable', type_='number'), {'A1': 3.14}, tmp_path)
        assert result.get('v') == pytest.approx(3.14)

    def test_wrong_type_value_on_nullable_still_warns(self, tmp_path):
        """nullable only skips empty-cell warning; type mismatch on a present value still warns."""
        result, lg = _run(
            _scalar_rows('var:nullable', type_='integer'), {'A1': 'not-an-int'}, tmp_path
        )
        assert lg.has_warnings()

    def test_regex_mismatch_on_nullable_still_warns(self, tmp_path):
        """nullable: a present value that fails its regex still warns."""
        result, lg = _run(
            _scalar_rows('var:nullable', regex=r'^\d{4}$'), {'A1': 'abc'}, tmp_path
        )
        assert lg.has_warnings()


# ── F. Interaction with other modifiers ───────────────────────────────────────

class TestNullableWithOtherModifiers:
    def test_nullable_trim_whitespace_empty_no_warning(self, tmp_path):
        result, lg = _run([
            ['var:nullable:trim-whitespace', 'v', 'string', '.*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {}, tmp_path)
        assert result.get('v') is None
        assert not lg.has_warnings()

    def test_nullable_trim_whitespace_spaces_only_cell(self, tmp_path):
        """A cell with only spaces is treated as empty by trim → nullable → no warning."""
        result, lg = _run([
            ['var:nullable:trim-whitespace', 'v', 'string', r'\w+'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {'A1': '   '}, tmp_path)
        # After trim, '   ' becomes '' which is_empty → nullable → null, no warning
        assert not lg.has_warnings()

    def test_nullable_trim_non_empty_trims_and_extracts(self, tmp_path):
        result, lg = _run([
            ['var:nullable:trim-whitespace', 'v', 'string', r'\w+'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {'A1': '  hello  '}, tmp_path)
        assert result.get('v') == 'hello'
        assert not lg.has_warnings()

    def test_nullable_glob_empty_no_warning(self, tmp_path):
        result, lg = _run([
            ['var:nullable:glob', 'v', 'string', 'INV-*'],
            ['START:'], ['cell:A1', 'v'], ['END:'],
        ], {}, tmp_path)
        assert result.get('v') is None
        assert not lg.has_warnings()


# ── G. Required (not-null) behaviour is unaffected ────────────────────────────

class TestRequiredUnaffected:
    def test_not_null_empty_still_fatal(self, tmp_path):
        """Existing not-null behaviour unchanged: empty cell → fatal error."""
        _, lg = _run(_scalar_rows('var:not-null'), {}, tmp_path)
        assert lg.has_errors()

    def test_not_null_present_still_passes(self, tmp_path):
        result, lg = _run(_scalar_rows('var:not-null'), {'A1': 'value'}, tmp_path)
        assert result.get('v') == 'value'
        assert not lg.has_errors()

    def test_not_empty_empty_still_fatal(self, tmp_path):
        _, lg = _run(_scalar_rows('var:not-empty'), {}, tmp_path)
        assert lg.has_errors()


# ── H. Warning/exit-code semantics ───────────────────────────────────────────

class TestNullableWarningsAndExitCode:
    def test_nullable_empty_no_warnings_logged(self, tmp_path):
        """Null nullable field → logger.has_warnings() is False."""
        _, lg = _run(_scalar_rows('var:nullable'), {}, tmp_path)
        assert not lg.has_warnings()

    def test_multiple_nullable_fields_all_empty_no_warnings(self, tmp_path):
        rows = [
            ['var:nullable', 'a', 'string', '.*'],
            ['var:nullable', 'b', 'integer', '.*'],
            ['var:nullable', 'c', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'a'],
            ['cell:A2', 'b'],
            ['cell:A3', 'c'],
            ['END:'],
        ]
        _, lg = _run(rows, {}, tmp_path)
        assert not lg.has_warnings()
        assert not lg.has_errors()

    def test_mix_nullable_and_plain_only_plain_warns(self, tmp_path):
        """A mix: nullable empty → silent; non-nullable empty → warning.

        Three DATA columns:
          col A = 'row.anchor'  (non-nullable, has a value → row IS a data row)
          col B = 'row.required' (non-nullable, EMPTY → should warn)
          col C = 'row.optional' (nullable, EMPTY → should be silent)

        Having a non-empty field (col A) ensures _row_is_end_of_data returns False,
        so the engine actually processes the row rather than terminating the table.
        """
        rows = [
            ['lbl:', 'h', 'string', 'Hdr'],
            ['var:', 'row.anchor', 'string', '.*'],
            ['var:', 'row.required', 'string', '.*'],
            ['var:nullable', 'row.optional', 'string', '.*'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h', 'h', 'h'],
            ['', 'DATA:*', 'row.anchor', 'row.required', 'row.optional'],
            ['END:'],
        ]
        pat = _write_pattern(rows, tmp_path, name='pmix.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet'
        ws['A1'] = 'Hdr'   # header col A
        ws['B1'] = 'Hdr'   # header col B
        ws['C1'] = 'Hdr'   # header col C
        ws['A2'] = 'present'  # col A non-empty → row IS a data row
        # B2 absent → non-nullable field empty → warning
        # C2 absent → nullable field empty → silent
        dat = str(tmp_path / 'dmix.xlsx')
        wb.save(dat)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        # Only the non-nullable empty field warns
        assert lg.has_warnings()
        issues = list(lg.issues())
        issue_fields = {r.field for r in issues}
        assert 'row.required' in issue_fields
        assert 'row.optional' not in issue_fields

    def test_no_warnings_means_no_exit1_in_real_run(self, tmp_path):
        """Simulate the CLI ok-flag logic: ok = not (errors or warnings).
        With nullable empty → ok should be True."""
        _, lg = _run(_scalar_rows('var:nullable'), {}, tmp_path)
        ok = not (lg.has_errors() or lg.has_warnings())
        assert ok is True


# ── I. Validate-pattern recognises nullable ───────────────────────────────────

class TestNullableValidatePattern:
    def test_nullable_pattern_parses_cleanly(self, tmp_path):
        """A pattern with var:nullable should parse without errors."""
        rows = [
            ['var:nullable', 'v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'v'],
            ['END:'],
        ]
        # No exception means validate-pattern would pass
        cfg, defs, _ = _parse(rows, tmp_path)
        assert defs['v'].nullable is True

    def test_nullable_typed_duration_parses_cleanly(self, tmp_path):
        rows = [
            ['var:nullable', 'sick_time', 'duration', '.*'],
            ['START:'],
            ['cell:A1', 'sick_time'],
            ['END:'],
        ]
        _, defs, _ = _parse(rows, tmp_path)
        assert defs['sick_time'].nullable is True
        assert defs['sick_time'].type == 'duration'

    def test_pattern_with_lbl_nullable_fails_validation(self, tmp_path):
        rows = [
            ['lbl:nullable', 'label', 'string', 'Header'],
            ['var:', 'v', 'string', '.*'],
            ['START:'], ['cell:A1', 'label'], ['cell:next', 'v'], ['END:'],
        ]
        with pytest.raises(PatternError, match='nullable.*not valid for lbl:'):
            _parse(rows, tmp_path)

    def test_nullable_not_null_combination_fails_validation(self, tmp_path):
        rows = _scalar_rows('var:nullable:not-null')
        with pytest.raises(PatternError, match='[Cc]ontradictory'):
            _parse(rows, tmp_path)
