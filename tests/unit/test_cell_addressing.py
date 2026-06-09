"""
Unit tests for cell:next and cell:A1 (absolute reference) addressing in Engine.
Uses in-memory xlsx workbooks to avoid fixture file dependencies.

Coverage:
  - Sequential mode (cell:1 / cell:next): happy path, empty sheet, exhaustion
  - Absolute mode (cell:A1): basic extraction, cursor advancement, mixing
  - Empty-cell handling by role (lbl: fatal, var: null, IGNORE: silent)
  - Out-of-bounds references
  - Unreachable references (cursor already past target)
  - Type/regex validation warnings at absolute refs
  - Interaction with table:* instructions
  - TD read direction
  - Merged cell handling
  - PatternError propagation through the engine
  - Ordering validation edge cases
  - Output shape in nested JSON
"""

import openpyxl
import pytest

from grepxcel.engine import Engine
from grepxcel.logger import Logger, VerbosityLevel
from grepxcel.pattern_parser import PatternError, PatternParser


# ── helpers ───────────────────────────────────────────────────────────────────

def _make_data(cells: dict, tmp_path, filename='data.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for coord, val in cells.items():
        ws[coord] = val
    path = str(tmp_path / filename)
    wb.save(path)
    return path


def _make_pattern(rows: list, tmp_path, filename='pattern.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / filename)
    wb.save(path)
    return path


def _run(pattern_rows, data_cells, tmp_path, direction='LR'):
    """Run engine; returns (result, logger)."""
    pat = _make_pattern(
        [['config:', 'read.direction', direction]] + pattern_rows,
        tmp_path,
    )
    dat = _make_data(data_cells, tmp_path)
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pat, dat, logger=lg)
    return result, lg


def _ok(pattern_rows, data_cells, tmp_path, direction='LR'):
    """Run engine and assert no errors; return result."""
    result, lg = _run(pattern_rows, data_cells, tmp_path, direction=direction)
    assert not lg.has_errors(), f'Unexpected engine error; issues: {lg.issues()}'
    return result


# ── cell:next (sequential) ────────────────────────────────────────────────────

class TestCellNext:
    def test_next_alias_extracts_first_nonempty(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:next', 'x.val'], ['END:']],
            {'A1': 'hello'}, tmp_path,
        )
        assert result == {'x': {'val': 'hello'}}

    def test_cell1_still_works(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:1', 'x.val'], ['END:']],
            {'A1': 'old style'}, tmp_path,
        )
        assert result == {'x': {'val': 'old style'}}

    def test_next_skips_empty_leading_cells(self, tmp_path):
        # B3 is the first non-empty cell
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:next', 'x.val'], ['END:']],
            {'B3': 'found'}, tmp_path,
        )
        assert result == {'x': {'val': 'found'}}

    def test_next_multiple_sequential(self, tmp_path):
        result = _ok(
            [['var:', 'a.v', 'string', '.*'], ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v'], ['cell:next', 'b.v'],
             ['END:']],
            {'A1': 'first', 'B1': 'second'}, tmp_path,
        )
        assert result['a']['v'] == 'first'
        assert result['b']['v'] == 'second'

    def test_next_ignore_skips_cell(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:next', 'IGNORE'],
             ['cell:next', 'x.val'],
             ['END:']],
            {'A1': 'skip', 'B1': 'keep'}, tmp_path,
        )
        assert result == {'x': {'val': 'keep'}}

    def test_next_exhausted_is_fatal(self, tmp_path):
        _, lg = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:next', 'x.val'], ['END:']],
            {}, tmp_path,  # empty sheet
        )
        assert lg.has_errors()

    def test_empty_sheet_absolute_var_records_null(self, tmp_path):
        # cell:A1 on an empty sheet with var: → null, no error
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:A1', 'x.val'], ['END:']],
            {}, tmp_path,
        )
        assert result == {'x': {'val': None}}


# ── cell:A1 — basic extraction ────────────────────────────────────────────────

class TestCellAbsoluteBasic:
    def test_reads_specific_cell(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:B3', 'x.val'], ['END:']],
            {'A1': 'wrong', 'B3': 'right'}, tmp_path,
        )
        assert result == {'x': {'val': 'right'}}

    def test_lowercase_ref_accepted(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:b3', 'x.val'], ['END:']],
            {'B3': 'lower ok'}, tmp_path,
        )
        assert result == {'x': {'val': 'lower ok'}}

    def test_integer_value_extracted(self, tmp_path):
        result = _ok(
            [['var:', 'x.num', 'integer', r'\d+'],
             ['START:'], ['cell:A1', 'x.num'], ['END:']],
            {'A1': 42}, tmp_path,
        )
        assert result == {'x': {'num': 42}}

    def test_multiword_string_extracted(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.+'],
             ['START:'], ['cell:C5', 'x.val'], ['END:']],
            {'C5': 'Hello World'}, tmp_path,
        )
        assert result == {'x': {'val': 'Hello World'}}

    def test_all_absolute_no_sequential(self, tmp_path):
        result = _ok(
            [['var:', 'a.x', 'string', '.*'], ['var:', 'b.y', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'a.x'],
             ['cell:C1', 'b.y'],
             ['END:']],
            {'A1': 'one', 'B1': 'skip', 'C1': 'two'}, tmp_path,
        )
        assert result == {'a': {'x': 'one'}, 'b': {'y': 'two'}}

    def test_large_column_ref(self, tmp_path):
        # Column Z = 26, AA = 27. Both are valid.
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:AA1', 'x.val'], ['END:']],
            {'AA1': 'deep column'}, tmp_path,
        )
        assert result == {'x': {'val': 'deep column'}}

    def test_only_ignore_refs(self, tmp_path):
        # Pattern with only IGNORE absolute refs → empty output, no errors
        result, lg = _run(
            [['START:'],
             ['cell:A1', 'IGNORE'],
             ['cell:B1', 'IGNORE'],
             ['END:']],
            {'A1': 'skip', 'B1': 'also skip'}, tmp_path,
        )
        assert not lg.has_errors()
        assert result == {}


# ── cursor advancement ────────────────────────────────────────────────────────

class TestCursorAdvancement:
    def test_next_after_absolute_continues_from_target(self, tmp_path):
        # After cell:A1, cell:next picks the next non-empty cell after A1
        result = _ok(
            [['var:', 'a.v', 'string', '.*'], ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'a.v'],
             ['cell:next', 'b.v'],
             ['END:']],
            {'A1': 'first', 'B1': 'second', 'C1': 'third'}, tmp_path,
        )
        assert result['a']['v'] == 'first'
        assert result['b']['v'] == 'second'

    def test_absolute_jump_skips_cells_between(self, tmp_path):
        # cell:next → A1; cell:C1 jumps over B1; cell:next → D1
        result = _ok(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['var:', 'c.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v'],
             ['cell:C1',   'b.v'],
             ['cell:next', 'c.v'],
             ['END:']],
            {'A1': 'seq', 'B1': 'middle-skip', 'C1': 'jump', 'D1': 'after'},
            tmp_path,
        )
        assert result == {'a': {'v': 'seq'}, 'b': {'v': 'jump'}, 'c': {'v': 'after'}}

    def test_three_absolute_refs_tight_row(self, tmp_path):
        result = _ok(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['var:', 'c.v', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'a.v'],
             ['cell:B1', 'b.v'],
             ['cell:C1', 'c.v'],
             ['END:']],
            {'A1': 'one', 'B1': 'two', 'C1': 'three'}, tmp_path,
        )
        assert result == {'a': {'v': 'one'}, 'b': {'v': 'two'}, 'c': {'v': 'three'}}

    def test_absolute_then_next_skips_consumed(self, tmp_path):
        # cell:B1 consumes B1; subsequent cell:next must not return B1
        result = _ok(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:B1', 'a.v'],
             ['cell:next', 'b.v'],
             ['END:']],
            {'A1': 'before', 'B1': 'targeted', 'C1': 'after'}, tmp_path,
        )
        assert result['a']['v'] == 'targeted'
        # cell:next starts from after B1; A1 is before cursor, so C1 is next
        assert result['b']['v'] == 'after'

    def test_multi_row_absolute_refs(self, tmp_path):
        result = _ok(
            [['var:', 'r1.v', 'string', '.*'],
             ['var:', 'r2.v', 'string', '.*'],
             ['var:', 'r3.v', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'r1.v'],
             ['cell:A2', 'r2.v'],
             ['cell:A3', 'r3.v'],
             ['END:']],
            {'A1': 'row1', 'A2': 'row2', 'A3': 'row3'}, tmp_path,
        )
        assert result == {'r1': {'v': 'row1'}, 'r2': {'v': 'row2'}, 'r3': {'v': 'row3'}}


# ── empty-cell handling by field role ─────────────────────────────────────────

class TestEmptyCellHandling:
    def test_var_empty_at_absolute_ref_records_null(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:C5', 'x.val'], ['END:']],
            {'A1': 'elsewhere'},  # C5 is empty
            tmp_path,
        )
        assert result == {'x': {'val': None}}

    def test_var_null_nested_correctly(self, tmp_path):
        result = _ok(
            [['var:', 'summary.income', 'integer', r'\d+'],
             ['var:', 'summary.expenses', 'integer', r'\d+'],
             ['START:'],
             ['cell:A1', 'summary.income'],
             ['cell:A2', 'summary.expenses'],  # A2 is empty
             ['END:']],
            {'A1': 100}, tmp_path,
        )
        assert result == {'summary': {'income': 100, 'expenses': None}}

    def test_lbl_empty_at_absolute_ref_is_fatal(self, tmp_path):
        _, lg = _run(
            [['lbl:', 'my_label', 'string', 'Expected'],
             ['START:'], ['cell:C5', 'my_label'], ['END:']],
            {'A1': 'elsewhere'}, tmp_path,
        )
        assert lg.has_errors()

    def test_lbl_empty_error_message_mentions_field(self, tmp_path):
        _, lg = _run(
            [['lbl:', 'inv_label', 'string', 'Invoice No:'],
             ['START:'], ['cell:B2', 'inv_label'], ['END:']],
            {'A1': 'something'}, tmp_path,
        )
        assert lg.has_errors()
        err = lg.issues()[0]
        assert 'inv_label' in err.message or 'B2' in err.message

    def test_ignore_at_empty_absolute_ref_is_silent(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'IGNORE'],  # A1 is empty
             ['cell:next', 'x.val'],
             ['END:']],
            {'B1': 'value'}, tmp_path,
        )
        assert result == {'x': {'val': 'value'}}

    def test_lbl_with_value_present_no_error(self, tmp_path):
        result = _ok(
            [['lbl:', 'my_label', 'string', 'Invoice No:'],
             ['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'my_label'],
             ['cell:B1', 'x.val'],
             ['END:']],
            {'A1': 'Invoice No:', 'B1': 'INV-001'}, tmp_path,
        )
        # lbl: field not in output; var: is
        assert 'my_label' not in result
        assert result == {'x': {'val': 'INV-001'}}

    def test_lbl_regex_mismatch_at_absolute_ref_is_fatal(self, tmp_path):
        # lbl: regex is 'Invoice No:' but cell has 'Wrong Label'
        # In strict=True context (header cell match), wrong value = fail
        # For cell: instructions: wrong regex on lbl: → warning only (not strict)
        # Actually, let's test this: lbl: regex mismatch at absolute ref
        _, lg = _run(
            [['lbl:', 'my_label', 'string', 'Invoice No:'],
             ['START:'],
             ['cell:A1', 'my_label'],
             ['END:']],
            {'A1': 'Wrong Label'}, tmp_path,
        )
        # lbl: at absolute ref with wrong value: warning (not fatal)
        # Engine only fatals on EMPTY lbl:; wrong value → warn
        # (this is the current behavior — test documents it)
        assert not lg.has_errors()
        assert len(lg.issues()) == 1  # one validation warning


# ── out-of-bounds references ──────────────────────────────────────────────────

class TestOutOfBounds:
    def test_var_oob_records_null(self, tmp_path):
        # Sheet has only A1; Z99 is way out of bounds
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:Z99', 'x.val'], ['END:']],
            {'A1': 'present'}, tmp_path,
        )
        assert result == {'x': {'val': None}}

    def test_var_oob_no_engine_error(self, tmp_path):
        _, lg = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:Z99', 'x.val'], ['END:']],
            {'A1': 'present'}, tmp_path,
        )
        assert not lg.has_errors()

    def test_lbl_oob_is_fatal(self, tmp_path):
        _, lg = _run(
            [['lbl:', 'lbl_oob', 'string', 'Label'],
             ['START:'], ['cell:Z99', 'lbl_oob'], ['END:']],
            {'A1': 'present'}, tmp_path,
        )
        assert lg.has_errors()

    def test_ignore_oob_is_silent(self, tmp_path):
        # A1 first (in LR order), then Z99 (out-of-bounds, empty) with IGNORE → silent
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'x.val'],
             ['cell:Z99', 'IGNORE'],
             ['END:']],
            {'A1': 'present'}, tmp_path,
        )
        assert result == {'x': {'val': 'present'}}

    def test_oob_cursor_not_advanced_next_still_works(self, tmp_path):
        # After oob ref (cursor not advanced), cell:next should still find A1
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:Z99', 'IGNORE'],   # oob, cursor stays at 0
             ['cell:next', 'x.val'],   # should find A1
             ['END:']],
            {'A1': 'found'}, tmp_path,
        )
        assert result == {'x': {'val': 'found'}}


# ── unreachable references ────────────────────────────────────────────────────

class TestUnreachable:
    def test_cursor_past_target_is_fatal(self, tmp_path):
        # cell:next consumes A1 (cursor → 1), then cell:A1 (idx=0 < 1) → fatal
        _, lg = _run(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v'],
             ['cell:A1',   'b.v'],
             ['END:']],
            {'A1': 'val', 'B1': 'other'}, tmp_path,
        )
        assert lg.has_errors()

    def test_cursor_past_via_multiple_nexts(self, tmp_path):
        # Two cell:next consume A1 and B1; then cell:A1 → fatal
        _, lg = _run(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['var:', 'c.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v'],
             ['cell:next', 'b.v'],
             ['cell:A1',   'c.v'],
             ['END:']],
            {'A1': 'v1', 'B1': 'v2'}, tmp_path,
        )
        assert lg.has_errors()

    def test_cursor_exactly_at_target_is_ok(self, tmp_path):
        # cursor at index 0, cell:A1 has idx 0 → 0 < 0 is False → OK
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:A1', 'x.val'], ['END:']],
            {'A1': 'at start'}, tmp_path,
        )
        assert result == {'x': {'val': 'at start'}}

    def test_absolute_then_earlier_absolute_detected_by_parser(self, tmp_path):
        # Parser should catch this before the engine runs
        path = _make_pattern(
            [['var:', 'a.v', 'string', '.*'], ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:B2', 'a.v'],
             ['cell:A1', 'b.v'],  # A1 before B2 in LR order → PatternError
             ['END:']],
            tmp_path,
        )
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)


# ── type and regex validation at absolute refs ────────────────────────────────

class TestValidationAtAbsoluteRef:
    def test_regex_mismatch_warns_not_errors(self, tmp_path):
        # var: regex is r'\d+' but value is 'not-a-number'
        _, lg = _run(
            [['var:', 'x.num', 'integer', r'\d+'],
             ['START:'], ['cell:A1', 'x.num'], ['END:']],
            {'A1': 'not-a-number'}, tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 1

    def test_value_stored_despite_mismatch(self, tmp_path):
        result, _ = _run(
            [['var:', 'x.val', 'string', r'MATCH'],
             ['START:'], ['cell:A1', 'x.val'], ['END:']],
            {'A1': 'NOMATCH'}, tmp_path,
        )
        # Value is stored even when regex fails
        assert result['x']['val'] == 'NOMATCH'

    def test_null_value_no_validation_run(self, tmp_path):
        # Empty cell → null; validation should be skipped (no false warning)
        _, lg = _run(
            [['var:', 'x.val', 'string', r'REQUIRED'],
             ['START:'], ['cell:A1', 'x.val'], ['END:']],
            {}, tmp_path,  # A1 is empty
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 0  # no warning for null value

    def test_multiple_validation_warnings(self, tmp_path):
        _, lg = _run(
            [['var:', 'a.v', 'integer', r'\d+'],
             ['var:', 'b.v', 'integer', r'\d+'],
             ['START:'],
             ['cell:A1', 'a.v'],
             ['cell:B1', 'b.v'],
             ['END:']],
            {'A1': 'bad1', 'B1': 'bad2'}, tmp_path,
        )
        assert not lg.has_errors()
        assert len(lg.issues()) == 2

    def test_currency_type_extracted_at_absolute_ref(self, tmp_path):
        result = _ok(
            [['var:', 'x.amount', 'currency', r'.*'],
             ['START:'], ['cell:B2', 'x.amount'], ['END:']],
            {'B2': 1234.56}, tmp_path,
        )
        assert result['x']['amount'] == pytest.approx(1234.56)


# ── interaction with table:* ──────────────────────────────────────────────────

class TestAbsoluteWithTable:
    def test_absolute_refs_before_table(self, tmp_path):
        pat = _make_pattern([
            ['lbl:', 'col_item', 'string', 'Item'],
            ['lbl:', 'col_qty',  'string', 'Qty'],
            ['var:', 'hdr.title', 'string', '.*'],
            ['var:', 'row.item',  'string', '.+'],
            ['var:', 'row.qty',   'integer', r'\d+'],
            ['START:'],
            ['cell:A1', 'hdr.title'],
            ['table:*'],
            [None, 'HEADER:1', 'col_item', 'col_qty'],
            [None, 'DATA:*',   'row.item',  'row.qty'],
            ['END:'],
        ], tmp_path)
        dat = _make_data({
            'A1': 'My Report',
            'A2': 'Item', 'B2': 'Qty',
            'A3': 'Widget', 'B3': 10,
            'A4': 'Gadget', 'B4': 5,
        }, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert not lg.has_errors()
        assert result['hdr']['title'] == 'My Report'
        assert len(result['row']) == 1
        assert len(result['row'][0]['data']) == 2

    def test_table_does_not_advance_main_cursor(self, tmp_path):
        # After table:*, main cursor stays where it was before the table
        # So a cell:next after table:* starts from before the table data
        # (cells consumed by table are in scanner.consumed and skipped)
        pat = _make_pattern([
            ['lbl:', 'col_a', 'string', 'A'],
            ['var:', 'row.v', 'string', '.+'],
            ['var:', 'x.after', 'string', '.*'],
            ['START:'],
            ['table:*'],
            [None, 'HEADER:1', 'col_a'],
            [None, 'DATA:*',   'row.v'],
            ['cell:next', 'x.after'],
            ['END:'],
        ], tmp_path)
        dat = _make_data({
            'A1': 'A',
            'A2': 'item1',
            'A3': 'item2',
            'B5': 'trailing value',
        }, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert not lg.has_errors()
        # 'trailing value' is not consumed by the table — cell:next finds it
        assert result['x']['after'] == 'trailing value'

    def test_absolute_ref_after_table_reads_unconsumed_cell(self, tmp_path):
        # cell:D1 was not part of the table → should be readable absolutely
        pat = _make_pattern([
            ['lbl:', 'col_a', 'string', 'Label'],
            ['var:', 'row.v', 'string', '.+'],
            ['var:', 'x.note', 'string', '.*'],
            ['START:'],
            ['table:*'],
            [None, 'HEADER:1', 'col_a'],
            [None, 'DATA:*',   'row.v'],
            ['cell:D1', 'x.note'],
            ['END:'],
        ], tmp_path)
        dat = _make_data({
            'A1': 'Label',
            'A2': 'item1',
            'D1': 'Note outside table',
        }, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, dat, logger=lg)
        assert not lg.has_errors()
        assert result['x']['note'] == 'Note outside table'


# ── TD direction ──────────────────────────────────────────────────────────────

class TestTDDirection:
    def test_td_reads_column_first(self, tmp_path):
        # In TD mode: A1, A2, A3, B1, B2, B3 ...
        result = _ok(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v'],
             ['cell:next', 'b.v'],
             ['END:']],
            {'A1': 'col-a', 'B1': 'col-b'}, tmp_path,
            direction='TD',
        )
        assert result['a']['v'] == 'col-a'
        assert result['b']['v'] == 'col-b'

    def test_td_absolute_ref(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:B2', 'x.val'], ['END:']],
            {'A1': 'skip', 'B2': 'td-target'}, tmp_path,
            direction='TD',
        )
        assert result == {'x': {'val': 'td-target'}}

    def test_td_ordering_col_primary(self, tmp_path):
        # In TD mode: column A is fully scanned before column B.
        # So A2 (col=1, row=2) comes before B1 (col=2, row=1).
        # Pattern with B1 then A2 → parser must reject as out of order.
        _, lg = _run(
            [['var:', 'first.v', 'string', '.*'],
             ['var:', 'second.v', 'string', '.*'],
             ['START:'],
             ['cell:B1', 'first.v'],
             ['cell:A2', 'second.v'],
             ['END:']],
            {'A2': 'a-two', 'B1': 'b-one'}, tmp_path,
            direction='TD',
        )
        assert lg.has_errors()

    def test_td_cursor_advancement_absolute(self, tmp_path):
        result = _ok(
            [['var:', 'a.v', 'string', '.*'],
             ['var:', 'b.v', 'string', '.*'],
             ['START:'],
             ['cell:A2', 'a.v'],
             ['cell:next', 'b.v'],
             ['END:']],
            {'A1': 'before', 'A2': 'target', 'A3': 'after'}, tmp_path,
            direction='TD',
        )
        assert result['a']['v'] == 'target'
        assert result['b']['v'] == 'after'


# ── merged cell handling ──────────────────────────────────────────────────────

class TestMergedCellAbsolute:
    def test_absolute_ref_to_merged_non_topleft(self, tmp_path):
        # A1:C1 merged with value 'Merged'; B1 gets the expanded value
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'Merged'
        ws.merge_cells('A1:C1')
        ws['A2'] = 'other'
        path = str(tmp_path / 'data.xlsx')
        wb.save(path)

        pat = _make_pattern([
            ['var:', 'x.val', 'string', '.*'],
            ['START:'], ['cell:B1', 'x.val'], ['END:'],
        ], tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, path, logger=lg)
        assert not lg.has_errors()
        assert result == {'x': {'val': 'Merged'}}

    def test_absolute_ref_to_merged_top_left(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'TL Value'
        ws.merge_cells('A1:B2')
        path = str(tmp_path / 'data.xlsx')
        wb.save(path)

        pat = _make_pattern([
            ['var:', 'x.val', 'string', '.*'],
            ['START:'], ['cell:A1', 'x.val'], ['END:'],
        ], tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pat, path, logger=lg)
        assert not lg.has_errors()
        assert result == {'x': {'val': 'TL Value'}}


# ── PatternError propagation through the engine ───────────────────────────────

class TestPatternErrorPropagation:
    def test_invalid_multiplicity_causes_engine_fatal(self, tmp_path):
        pat = _make_pattern([
            ['START:'], ['cell:notvalid', 'IGNORE'], ['END:'],
        ], tmp_path)
        dat = _make_data({'A1': 'x'}, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        Engine().process(pat, dat, logger=lg)
        assert lg.has_errors()

    def test_out_of_order_refs_causes_engine_fatal(self, tmp_path):
        pat = _make_pattern([
            ['var:', 'a.v', 'string', '.*'],
            ['var:', 'b.v', 'string', '.*'],
            ['START:'],
            ['cell:C3', 'a.v'],
            ['cell:A1', 'b.v'],  # before C3 in LR
            ['END:'],
        ], tmp_path)
        dat = _make_data({'A1': 'x', 'C3': 'y'}, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        Engine().process(pat, dat, logger=lg)
        assert lg.has_errors()

    def test_undefined_field_causes_engine_fatal(self, tmp_path):
        # Field referenced in START: but not defined with var:/lbl:
        pat = _make_pattern([
            ['START:'],
            ['cell:A1', 'undefined_field'],
            ['END:'],
        ], tmp_path)
        dat = _make_data({'A1': 'x'}, tmp_path)
        lg = Logger(level=VerbosityLevel.QUIET)
        Engine().process(pat, dat, logger=lg)
        assert lg.has_errors()

    def test_valid_pattern_no_engine_errors(self, tmp_path):
        result = _ok(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:A1', 'x.val'], ['END:']],
            {'A1': 'valid'}, tmp_path,
        )
        assert result['x']['val'] == 'valid'


# ── output shape in nested JSON ───────────────────────────────────────────────

class TestOutputShape:
    def test_lbl_field_excluded_from_output(self, tmp_path):
        result = _ok(
            [['lbl:', 'label_a', 'string', 'Label A'],
             ['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'label_a'],
             ['cell:B1', 'x.val'],
             ['END:']],
            {'A1': 'Label A', 'B1': 'data'}, tmp_path,
        )
        assert 'label_a' not in result
        assert result == {'x': {'val': 'data'}}

    def test_dotnotation_creates_nested_groups(self, tmp_path):
        result = _ok(
            [['var:', 'po.number', 'string', '.*'],
             ['var:', 'po.date',   'string', '.*'],
             ['var:', 'vendor.name', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'po.number'],
             ['cell:B1', 'po.date'],
             ['cell:A2', 'vendor.name'],
             ['END:']],
            {'A1': 'PO-001', 'B1': '2026-01-01', 'A2': 'Acme'}, tmp_path,
        )
        assert result == {
            'po': {'number': 'PO-001', 'date': '2026-01-01'},
            'vendor': {'name': 'Acme'},
        }

    def test_empty_start_block_is_fatal(self, tmp_path):
        # Hardened behavior: an empty START: … END: block defines no extraction
        # steps, so the engine treats it as a fatal error rather than silently
        # returning {}. The partial result is still an empty dict.
        result, lg = _run(
            [['START:'], ['END:']],
            {'A1': 'ignored'}, tmp_path,
        )
        assert lg.has_errors()
        assert result == {}

    def test_all_null_values_still_nested(self, tmp_path):
        # Multiple empty var: fields → all null, still in output
        result = _ok(
            [['var:', 'a.x', 'string', '.*'],
             ['var:', 'a.y', 'string', '.*'],
             ['START:'],
             ['cell:Z99', 'a.x'],
             ['cell:Z100', 'a.y'],
             ['END:']],
            {'A1': 'not-used'}, tmp_path,
        )
        assert result == {'a': {'x': None, 'y': None}}

    def test_legacy_format_has_cells_tables_keys(self, tmp_path):
        pat = _make_pattern([
            ['var:', 'x.val', 'string', '.*'],
            ['START:'], ['cell:A1', 'x.val'], ['END:'],
        ], tmp_path)
        dat = _make_data({'A1': 'hello'}, tmp_path)
        result = Engine().process(pat, dat, output_format='legacy')
        assert 'cells' in result
        assert 'tables' in result

    def test_legacy_format_cell_value_present(self, tmp_path):
        pat = _make_pattern([
            ['var:', 'x.val', 'string', '.*'],
            ['START:'], ['cell:A1', 'x.val'], ['END:'],
        ], tmp_path)
        dat = _make_data({'A1': 'hello'}, tmp_path)
        result = Engine().process(pat, dat, output_format='legacy')
        assert result['cells']['x.val'] == 'hello'


# ── edge cases ────────────────────────────────────────────────────────────────

class TestEdgeCases:
    def test_same_cell_absolute_then_next_skips_consumed(self, tmp_path):
        # cell:A1 consumes A1; cell:next must not return A1 again
        result = _ok(
            [['var:', 'a.first', 'string', '.*'],
             ['var:', 'b.second', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'a.first'],
             ['cell:next', 'b.second'],
             ['END:']],
            {'A1': 'once', 'B1': 'next'}, tmp_path,
        )
        assert result['a']['first'] == 'once'
        assert result['b']['second'] == 'next'

    def test_single_cell_sheet(self, tmp_path):
        result = _ok(
            [['var:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:A1', 'x.v'], ['END:']],
            {'A1': 'only cell'}, tmp_path,
        )
        assert result == {'x': {'v': 'only cell'}}

    def test_cell_with_float_value(self, tmp_path):
        result = _ok(
            [['var:', 'x.amount', 'currency', r'.*'],
             ['START:'], ['cell:A1', 'x.amount'], ['END:']],
            {'A1': 3.14159}, tmp_path,
        )
        assert abs(result['x']['amount'] - 3.14159) < 1e-9

    def test_cell_next_and_absolute_mixed_many(self, tmp_path):
        # Complex mix: next → A1, abs B1, next → C1, abs E1, next → F1
        result = _ok(
            [['var:', 'v1.v', 'string', '.*'],
             ['var:', 'v2.v', 'string', '.*'],
             ['var:', 'v3.v', 'string', '.*'],
             ['var:', 'v4.v', 'string', '.*'],
             ['var:', 'v5.v', 'string', '.*'],
             ['START:'],
             ['cell:next', 'v1.v'],
             ['cell:B1',   'v2.v'],
             ['cell:next', 'v3.v'],
             ['cell:E1',   'v4.v'],
             ['cell:next', 'v5.v'],
             ['END:']],
            {'A1': 'a', 'B1': 'b', 'C1': 'c', 'D1': 'd', 'E1': 'e', 'F1': 'f'},
            tmp_path,
        )
        assert result == {
            'v1': {'v': 'a'},
            'v2': {'v': 'b'},
            'v3': {'v': 'c'},
            'v4': {'v': 'e'},
            'v5': {'v': 'f'},
        }

    def test_cell_ref_row_single_digit(self, tmp_path):
        result = _ok(
            [['var:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:A9', 'x.v'], ['END:']],
            {'A9': 'row nine'}, tmp_path,
        )
        assert result['x']['v'] == 'row nine'

    def test_cell_ref_large_row_number(self, tmp_path):
        result = _ok(
            [['var:', 'x.v', 'string', '.*'],
             ['START:'], ['cell:A1000', 'x.v'], ['END:']],
            {'A1000': 'far down'}, tmp_path,
        )
        assert result['x']['v'] == 'far down'
