"""
Tests that verify the engine and parser correctly trap error conditions.

Each test either:
  - asserts logger.has_errors() → engine correctly detected a fatal problem
  - raises PatternError from PatternParser.parse() → parser-level validation
  - asserts specific warning counts → non-fatal validation issues logged

These tests ensure that bad input does not produce silent wrong results.
"""

import openpyxl
import pytest

from grepxcel.engine import Engine
from grepxcel.logger import Logger, VerbosityLevel
from grepxcel.pattern_parser import PatternError, PatternParser
from grepxcel.security import SecurityError


# ── helpers ───────────────────────────────────────────────────────────────────

def _write_pattern(rows: list, tmp_path, name='pattern.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / name)
    wb.save(path)
    return path


def _write_data(cells: dict, tmp_path, name='data.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for coord, val in cells.items():
        ws[coord] = val
    path = str(tmp_path / name)
    wb.save(path)
    return path


def _engine_run(pattern_rows, data_cells, tmp_path) -> Logger:
    pat = _write_pattern(pattern_rows, tmp_path)
    dat = _write_data(data_cells, tmp_path)
    lg = Logger(level=VerbosityLevel.QUIET)
    Engine().process(pat, dat, logger=lg)
    return lg


# ═════════════════════════════════════════════════════════════════════════════
# PARSER-LEVEL ERRORS (PatternError raised before engine runs)
# ═════════════════════════════════════════════════════════════════════════════

class TestParserOrderingErrors:
    """Parser must reject absolute-ref sequences that are impossible to satisfy."""

    def test_lr_b2_then_a1_rejected(self, tmp_path):
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:B2', 'a'], ['cell:A1', 'b'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)

    def test_lr_same_cell_twice_rejected(self, tmp_path):
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:C3', 'a'], ['cell:C3', 'b'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)

    def test_lr_later_row_accepted(self, tmp_path):
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:A1', 'a'], ['cell:A2', 'b'], ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 2

    def test_lr_same_row_later_col_accepted(self, tmp_path):
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:A1', 'a'], ['cell:C1', 'b'], ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 2

    def test_lr_same_row_earlier_col_rejected(self, tmp_path):
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:C1', 'a'], ['cell:A1', 'b'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError):
            PatternParser().parse(path)

    def test_td_b1_then_a2_rejected(self, tmp_path):
        # In TD mode: column primary. B1 (col 2, row 1) comes before A2 (col 1, row 2)?
        # No — TD scans col 1 fully first (A1, A2, ...), then col 2 (B1, B2, ...).
        # So A2 (col 1, row 2) comes BEFORE B1 (col 2, row 1).
        # Pattern: B1 then A2 → B1 position > A2 position in TD → reject
        path = _write_pattern([
            ['config:', 'read.direction', 'TD'],
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:B1', 'a'], ['cell:A2', 'b'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)

    def test_td_a2_then_b1_accepted(self, tmp_path):
        # TD: A2 comes before B1 (column A fully scanned before column B)
        path = _write_pattern([
            ['config:', 'read.direction', 'TD'],
            ['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
            ['START:'], ['cell:A2', 'a'], ['cell:B1', 'b'], ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].target == 'A2'
        assert seq[1].target == 'B1'

    def test_cell_next_between_absolutes_no_ordering_check(self, tmp_path):
        # cell:next doesn't participate in ordering validation
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'],
            ['var:', 'b', 'string', '.*'],
            ['var:', 'c', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'a'],
            ['cell:next', 'b'],   # no coordinate → skipped in ordering check
            ['cell:C1',  'c'],    # C1 > A1 in LR → OK even though next might be between
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 3


class TestParserInvalidSyntax:
    """Parser must reject malformed cell: instructions."""

    def test_cell_foo_rejected(self, tmp_path):
        path = _write_pattern([
            ['START:'], ['cell:foo', 'IGNORE'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='Unknown cell instruction'):
            PatternParser().parse(path)

    def test_cell_0_rejected(self, tmp_path):
        # Row 0 does not exist in Excel
        path = _write_pattern([
            ['START:'], ['cell:0', 'IGNORE'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='Unknown cell instruction'):
            PatternParser().parse(path)

    def test_cell_A0_rejected(self, tmp_path):
        # Row 0 — the regex requires first digit to be [1-9]
        path = _write_pattern([
            ['START:'], ['cell:A0', 'IGNORE'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='Unknown cell instruction'):
            PatternParser().parse(path)

    def test_cell_star_rejected(self, tmp_path):
        path = _write_pattern([
            ['START:'], ['cell:*', 'IGNORE'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='Unknown cell instruction'):
            PatternParser().parse(path)

    def test_cell_2_rejected(self, tmp_path):
        # Only '1' and 'next' are valid sequential multiplicities
        path = _write_pattern([
            ['START:'], ['cell:2', 'IGNORE'], ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='Unknown cell instruction'):
            PatternParser().parse(path)

    def test_cell_1_accepted(self, tmp_path):
        path = _write_pattern([['START:'], ['cell:1', 'IGNORE'], ['END:']], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == '1'

    def test_cell_next_accepted(self, tmp_path):
        path = _write_pattern([['START:'], ['cell:next', 'IGNORE'], ['END:']], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == 'next'

    def test_cell_A1_accepted(self, tmp_path):
        path = _write_pattern([['START:'], ['cell:A1', 'IGNORE'], ['END:']], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == 'abs'
        assert seq[0].target == 'A1'


class TestParserSecurityErrors:
    """Security validation must reject unsafe pattern content."""

    def test_formula_in_pattern_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = '=1+2'
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='Formulas'):
            PatternParser().parse(path)

    def test_numeric_value_in_pattern_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 99
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='plain text'):
            PatternParser().parse(path)

    def test_boolean_value_in_pattern_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = False
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='plain text'):
            PatternParser().parse(path)


# ═════════════════════════════════════════════════════════════════════════════
# ENGINE-LEVEL FATAL ERRORS (logged as errors, processing stops)
# ═════════════════════════════════════════════════════════════════════════════

class TestEngineFatalErrors:
    """Engine must log a fatal error for unrecoverable conditions."""

    def test_lbl_field_empty_at_absolute_ref(self, tmp_path):
        lg = _engine_run(
            [['lbl:', 'title', 'string', 'Expected Title'],
             ['START:'], ['cell:B2', 'title'], ['END:']],
            {'A1': 'something else'},  # B2 is empty
            tmp_path,
        )
        assert lg.has_errors()

    def test_sequential_next_on_empty_sheet(self, tmp_path):
        lg = _engine_run(
            [['var:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:next', 'x.v'], ['END:']],
            {},  # empty sheet
            tmp_path,
        )
        assert lg.has_errors()

    def test_undefined_field_in_start_sequence(self, tmp_path):
        # Field 'ghost' not defined with var:/lbl:
        lg = _engine_run(
            [['var:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:A1', 'ghost'], ['END:']],
            {'A1': 'val'},
            tmp_path,
        )
        assert lg.has_errors()

    def test_cursor_past_absolute_target(self, tmp_path):
        lg = _engine_run(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v'],   # consumes A1, cursor past index 0
             ['cell:A1',   'b.v'],   # A1 already past → fatal
             ['END:']],
            {'A1': 'val', 'B1': 'other'},
            tmp_path,
        )
        assert lg.has_errors()

    def test_pattern_syntax_error_propagates_to_engine(self, tmp_path):
        # PatternError from parser becomes engine fatal
        lg = _engine_run(
            [['START:'], ['cell:notvalid', 'IGNORE'], ['END:']],
            {'A1': 'val'},
            tmp_path,
        )
        assert lg.has_errors()

    def test_out_of_order_refs_propagate_to_engine(self, tmp_path):
        lg = _engine_run(
            [['var:', 'a', 'string', '.*'], ['var:', 'b', 'string', '.*'],
             ['START:'], ['cell:C3', 'a'], ['cell:A1', 'b'], ['END:']],
            {'A1': 'x', 'C3': 'y'},
            tmp_path,
        )
        assert lg.has_errors()

    def test_pattern_with_formula_propagates_to_engine(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = '=SUM(1,2)'
        pat_path = str(tmp_path / 'pattern.xlsx')
        wb.save(pat_path)
        dat_path = _write_data({'A1': 'data'}, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        Engine().process(pat_path, dat_path, logger=lg)
        assert lg.has_errors()


# ═════════════════════════════════════════════════════════════════════════════
# NON-FATAL WARNINGS (engine continues, but issues are logged)
# ═════════════════════════════════════════════════════════════════════════════

class TestEngineWarnings:
    """Engine should warn (not error) for validation failures in DATA rows
    and for regex/type mismatches at cell: instructions."""

    def test_regex_mismatch_at_absolute_ref_is_warning(self, tmp_path):
        lg = _engine_run(
            [['var:', 'x.code', 'string', r'[A-Z]{3}\d{3}'],
             ['START:'], ['cell:A1', 'x.code'], ['END:']],
            {'A1': 'not-a-code'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 1

    def test_regex_mismatch_at_sequential_ref_is_warning(self, tmp_path):
        lg = _engine_run(
            [['var:', 'x.code', 'string', r'[A-Z]+'],
             ['START:'], ['cell:next', 'x.code'], ['END:']],
            {'A1': '12345'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 1

    def test_null_value_no_warning(self, tmp_path):
        # Empty var: at absolute ref → null; NO warning (nothing to validate)
        lg = _engine_run(
            [['var:', 'x.v', 'string', r'REQUIRED'],
             ['START:'], ['cell:Z99', 'x.v'], ['END:']],
            {'A1': 'elsewhere'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 0

    def test_lbl_regex_mismatch_at_abs_ref_is_warning(self, tmp_path):
        # lbl: with wrong value at absolute ref → warning only, not fatal
        # (fatal only if the cell is EMPTY)
        lg = _engine_run(
            [['lbl:', 'hdr', 'string', 'Expected Header'],
             ['START:'], ['cell:A1', 'hdr'], ['END:']],
            {'A1': 'Wrong Header'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 1

    def test_multiple_warnings_accumulated(self, tmp_path):
        lg = _engine_run(
            [['var:', 'a.x', 'integer', r'\d+'],
             ['var:', 'b.y', 'integer', r'\d+'],
             ['var:', 'c.z', 'integer', r'\d+'],
             ['START:'],
             ['cell:A1', 'a.x'],
             ['cell:B1', 'b.y'],
             ['cell:C1', 'c.z'],
             ['END:']],
            {'A1': 'bad', 'B1': 'also bad', 'C1': 'still bad'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 3

    def test_one_valid_one_invalid_mixed(self, tmp_path):
        lg = _engine_run(
            [['var:', 'a.x', 'integer', r'\d+'],
             ['var:', 'b.y', 'integer', r'\d+'],
             ['START:'],
             ['cell:A1', 'a.x'],
             ['cell:B1', 'b.y'],
             ['END:']],
            {'A1': 42, 'B1': 'not-int'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 1


# ═════════════════════════════════════════════════════════════════════════════
# PARTIAL RESULTS AFTER FATAL ERRORS
# ═════════════════════════════════════════════════════════════════════════════

class TestPartialResultsAfterFatal:
    """After a fatal error, the engine returns whatever it extracted before
    the error. Verify the partial result is sane."""

    def test_cells_before_fatal_are_in_partial_result(self, tmp_path):
        pat = _write_pattern([
            ['var:', 'good.val', 'string', '.*'],
            ['lbl:', 'missing_lbl', 'string', 'Must Be Here'],
            ['START:'],
            ['cell:A1', 'good.val'],
            ['cell:B1', 'missing_lbl'],  # B1 is empty → fatal
            ['END:'],
        ], tmp_path)
        dat = _write_data({'A1': 'extracted before error'}, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert lg.has_errors()
        # 'good.val' was extracted before the fatal
        assert result.get('good', {}).get('val') == 'extracted before error'

    def test_partial_result_has_no_crashed_fields(self, tmp_path):
        pat = _write_pattern([
            ['var:', 'a.v', 'string', '.*'],
            ['var:', 'b.v', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'a.v'],
            ['cell:next', 'b.v'],  # sheet has only A1 → exhausted
            ['END:'],
        ], tmp_path)
        dat = _write_data({'A1': 'only cell'}, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert lg.has_errors()
        # a.v extracted before exhaustion
        assert result.get('a', {}).get('v') == 'only cell'
        # b.v was never extracted
        assert result.get('b') is None


# ═════════════════════════════════════════════════════════════════════════════
# CORRECT BEHAVIOR PRESERVED (regression guard)
# ═════════════════════════════════════════════════════════════════════════════

class TestRegressionGuards:
    """Ensure that adding cell:next/cell:A1 did not break existing patterns
    that use cell:1 syntax."""

    def test_cell_1_still_works_sequentially(self, tmp_path):
        lg = _engine_run(
            [['var:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:1', 'x.v'], ['END:']],
            {'A1': 'backward compat'},
            tmp_path,
        )
        assert not lg.has_errors()

    def test_cell_1_and_absolute_mixed(self, tmp_path):
        lg = _engine_run(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:1',  'a.v'],
             ['cell:C1', 'b.v'],
             ['END:']],
            {'A1': 'seq', 'B1': 'skip', 'C1': 'abs'},
            tmp_path,
        )
        assert not lg.has_errors()

    def test_def_alias_still_works(self, tmp_path):
        # def: is the old alias for var: — must still parse and run
        lg = _engine_run(
            [['def:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:A1', 'x.v'], ['END:']],
            {'A1': 'old alias'},
            tmp_path,
        )
        assert not lg.has_errors()

    def test_doc_rows_ignored(self, tmp_path):
        lg = _engine_run(
            [['doc:', 'This is a comment — ignored'],
             ['var:', 'x.v', 'string', '.*'],
             ['doc:', 'Another comment'],
             ['START:'],
             ['cell:A1', 'x.v'],
             ['END:']],
            {'A1': 'value'},
            tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 0
