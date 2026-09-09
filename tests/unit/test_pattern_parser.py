"""
Unit tests for PatternParser: var:, lbl:, doc:, def: (alias), config,
cell/table sequence, and security validation.
"""

from pathlib import Path

import openpyxl
import pytest

from grepxcel.pattern_parser import PatternParser, PatternError
from grepxcel.security import SecurityError


def _write_pattern(rows: list, tmp_path: Path) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / 'pattern.xlsx')
    wb.save(path)
    return path


# ── Field role parsing ────────────────────────────────────────────────────────

class TestFieldRoles:
    def test_var_produces_var_role(self, tmp_path):
        path = _write_pattern([['var:', 'po.number', 'string', r'PO-\d+']], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['po.number'].role == 'var'

    def test_lbl_produces_lbl_role(self, tmp_path):
        path = _write_pattern([['lbl:', 'po_label', 'string', 'PO Number:']], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['po_label'].role == 'lbl'

    def test_def_is_alias_for_var(self, tmp_path):
        path = _write_pattern([['def:', 'inv.number', 'string', r'INV-\d+']], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['inv.number'].role == 'var'

    def test_var_and_lbl_coexist(self, tmp_path):
        path = _write_pattern([
            ['lbl:', 'po_label',  'string', 'PO Number:'],
            ['var:', 'po.number', 'string', r'PO-\d+'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['po_label'].role == 'lbl'
        assert defs['po.number'].role == 'var'

    def test_type_and_regex_preserved(self, tmp_path):
        path = _write_pattern([['var:', 'amount', 'currency', r'\d+\.\d{2}']], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['amount'].type == 'currency'
        assert defs['amount'].regex == r'\d+\.\d{2}'

    def test_multiple_fields_all_parsed(self, tmp_path):
        path = _write_pattern([
            ['lbl:', 'a', 'string', 'A:'],
            ['lbl:', 'b', 'string', 'B:'],
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'integer', r'\d+'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert len(defs) == 4
        assert sum(1 for d in defs.values() if d.role == 'lbl') == 2
        assert sum(1 for d in defs.values() if d.role == 'var') == 2


# ── Comment rows ──────────────────────────────────────────────────────────────

class TestCommentRows:
    def test_doc_rows_produce_no_defs(self, tmp_path):
        path = _write_pattern([
            ['doc:', 'This is a comment'],
            ['var:', 'x', 'string', '.*'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert len(defs) == 1
        assert 'x' in defs

    def test_info_rows_produce_no_defs(self, tmp_path):
        path = _write_pattern([
            ['info:', 'Some info'],
            ['var:', 'y', 'string', '.*'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert len(defs) == 1

    def test_doc_does_not_interfere_with_subsequent_fields(self, tmp_path):
        path = _write_pattern([
            ['doc:', 'Section A'],
            ['lbl:', 'l1', 'string', 'Label:'],
            ['doc:', 'Section B'],
            ['var:', 'v1', 'string', '.*'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['l1'].role == 'lbl'
        assert defs['v1'].role == 'var'


# ── Config parsing ────────────────────────────────────────────────────────────

class TestConfigParsing:
    def test_read_direction(self, tmp_path):
        path = _write_pattern([['config:', 'read.direction', 'TD']], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert cfg.read_direction == 'TD'

    def test_currency_sign(self, tmp_path):
        path = _write_pattern([['config:', 'currency.sign', '$']], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert cfg.currency_sign == '$'

    def test_empty_aliases(self, tmp_path):
        path = _write_pattern([
            ['config:', 'empty.aliases', 'N/A'],
            ['config:', 'empty.aliases', '-'],
        ], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert 'N/A' in cfg.empty_aliases
        assert '-' in cfg.empty_aliases

    def test_defaults_when_no_config(self, tmp_path):
        path = _write_pattern([['var:', 'x', 'string', '.*']], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert cfg.read_direction == 'LR'
        assert cfg.currency_sign == '€'


# ── START/END sequence ────────────────────────────────────────────────────────

class TestSequenceParsing:
    def test_cell_instruction(self, tmp_path):
        path = _write_pattern([
            ['var:', 'inv.number', 'string', '.*'],
            ['START:'],
            ['cell:1', 'inv.number'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 1
        assert seq[0].field == 'inv.number'

    def test_cell_ignore(self, tmp_path):
        path = _write_pattern([
            ['START:'],
            ['cell:1', 'IGNORE'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].field == 'IGNORE'

    def test_table_instruction_parsed(self, tmp_path):
        from grepxcel.models import TableInstruction
        path = _write_pattern([
            ['lbl:', 'col_a', 'string', 'A'],
            ['var:', 'line.x', 'string', '.*'],
            ['START:'],
            ['table:*'],
            [None, 'HEADER:1', 'col_a'],
            [None, 'DATA:*', 'line.x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 1
        assert isinstance(seq[0], TableInstruction)
        assert seq[0].multiplicity == '*'
        assert len(seq[0].rows) == 2

    def test_mixed_cells_and_table(self, tmp_path):
        from grepxcel.models import CellInstruction, TableInstruction
        path = _write_pattern([
            ['var:', 'po.number', 'string', '.*'],
            ['var:', 'line.x', 'string', '.*'],
            ['START:'],
            ['cell:1', 'po.number'],
            ['table:*'],
            [None, 'DATA:*', 'line.x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert isinstance(seq[0], CellInstruction)
        assert isinstance(seq[1], TableInstruction)


# ── SKIP_IF ──────────────────────────────────────────────────────────────────

class TestSkipIfParsing:
    """SKIP_IF is valid with DATA:{n,m} and DATA:* — only forbidden with DATA:1."""

    def _skip_if_block(self, data_mult: str) -> list:
        return [
            ['lbl:', 'col_a', 'string', 'A'],
            ['var:', 'line.x', 'string', '.*'],
            ['START:'],
            ['table:*'],
            [None, 'HEADER:1', 'col_a'],
            [None, 'SKIP_IF', 'EMPTY'],
            [None, f'DATA:{data_mult}', 'line.x'],
            ['END:'],
        ]

    def test_skip_if_with_bounded_data_is_valid(self, tmp_path):
        path = _write_pattern(self._skip_if_block('{0,10}'), tmp_path)
        _, _, seq = PatternParser().parse(path)  # must not raise
        from grepxcel.models import TableInstruction
        table = next(s for s in seq if isinstance(s, TableInstruction))
        assert any(r.row_type == 'SKIP_IF' for r in table.rows)

    def test_skip_if_with_star_data_is_valid(self, tmp_path):
        """SKIP_IF with DATA:* must parse successfully — filter semantics are clear."""
        path = _write_pattern(self._skip_if_block('*'), tmp_path)
        _, _, seq = PatternParser().parse(path)  # must not raise
        from grepxcel.models import TableInstruction
        table = next(s for s in seq if isinstance(s, TableInstruction))
        assert any(r.row_type == 'SKIP_IF' for r in table.rows)

    def test_skip_if_with_data_1_raises(self, tmp_path):
        """SKIP_IF with DATA:1 is still invalid — ambiguous semantics."""
        path = _write_pattern(self._skip_if_block('1'), tmp_path)
        with pytest.raises(PatternError, match='SKIP_IF'):
            PatternParser().parse(path)


# ── Security validation ───────────────────────────────────────────────────────

class TestPatternParserSecurity:
    def test_formula_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = '=1+1'
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='Formulas are not allowed'):
            PatternParser().parse(path)

    def test_numeric_value_coerced_to_string(self, tmp_path):
        """Numbers are silently coerced — users often type 1 for pattern.version
        and Excel stores it as an integer."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 42        # integer stored by Excel
        ws['A2'] = 1.0       # whole float → "1"
        path = str(tmp_path / 'num.xlsx')
        wb.save(path)
        # parse() raises PatternError (empty/invalid sequence) but NOT SecurityError
        try:
            PatternParser().parse(path)
        except SecurityError:
            pytest.fail('SecurityError raised for numeric cell — should be coerced')
        except Exception:
            pass  # PatternError or similar is fine

    def test_boolean_value_coerced_to_string(self, tmp_path):
        """Booleans are silently coerced — Excel can store TRUE/FALSE as booleans."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = True
        path = str(tmp_path / 'bool.xlsx')
        wb.save(path)
        try:
            PatternParser().parse(path)
        except SecurityError:
            pytest.fail('SecurityError raised for boolean cell — should be coerced')
        except Exception:
            pass

    def test_datetime_value_rejected(self, tmp_path):
        """Dates/times are still rejected — no sensible string representation exists
        for a value that appears in a pattern configuration cell."""
        import datetime
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = datetime.datetime(2026, 1, 1)
        path = str(tmp_path / 'date.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='date/time'):
            PatternParser().parse(path)

    def test_overly_long_value_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'x' * 1001
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='too long'):
            PatternParser().parse(path)


# ── Cell addressing (cell:next / cell:A1) ────────────────────────────────────

class TestCellAddressing:
    def test_cell_1_parses_as_sequential(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['cell:1', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == '1'
        assert seq[0].target is None

    def test_cell_next_parses_as_sequential(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == 'next'
        assert seq[0].target is None

    def test_cell_A1_parses_as_absolute(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['cell:B5', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == 'abs'
        assert seq[0].target == 'B5'
        assert seq[0].field == 'x'

    def test_absolute_ref_normalized_to_uppercase(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['cell:b5', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].target == 'B5'

    def test_cell_ignore_with_absolute_ref(self, tmp_path):
        path = _write_pattern([
            ['START:'],
            ['cell:A1', 'IGNORE'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert seq[0].multiplicity == 'abs'
        assert seq[0].field == 'IGNORE'
        assert seq[0].target == 'A1'

    def test_out_of_order_absolute_refs_rejected(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['cell:B2', 'x'],
            ['cell:A1', 'y'],   # A1 comes before B2 in LR order
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)

    def test_same_position_absolute_refs_rejected(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'x'],
            ['cell:A1', 'y'],   # same cell twice
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)

    def test_cell_next_after_absolute_is_ok(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'x'],
            ['cell:next', 'y'],  # cell:next after absolute is always fine
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 2
        assert seq[0].multiplicity == 'abs'
        assert seq[1].multiplicity == 'next'

    def test_cell_next_before_absolute_is_ok(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['cell:next', 'x'],  # sequential before absolute is fine
            ['cell:B5', 'y'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert len(seq) == 2
        assert seq[1].multiplicity == 'abs'

    def test_invalid_cell_multiplicity_rejected(self, tmp_path):
        path = _write_pattern([
            ['START:'],
            ['cell:foo', 'IGNORE'],
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match="Unknown cell instruction"):
            PatternParser().parse(path)

    def test_td_direction_ordering_validated(self, tmp_path):
        # In TD mode, column is primary — B1 comes before A2
        path = _write_pattern([
            ['config:', 'read.direction', 'TD'],
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['cell:B1', 'x'],   # col B > col A → comes first in TD
            ['cell:A2', 'y'],   # col A < col B → A2 comes before B1 in TD
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='unreachable'):
            PatternParser().parse(path)

    def test_multiple_absolute_refs_in_order(self, tmp_path):
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'],
            ['var:', 'b', 'string', '.*'],
            ['var:', 'c', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'a'],
            ['cell:C1', 'b'],   # same row, later col
            ['cell:B2', 'c'],   # next row
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert [s.target for s in seq] == ['A1', 'C1', 'B2']


# ── ignore.case config ────────────────────────────────────────────────────────

class TestIgnoreCaseConfig:
    def test_default_is_case_sensitive(self, tmp_path):
        path = _write_pattern([['var:', 'po.number', 'string', r'PO-\d+']], tmp_path)
        config, _, _ = PatternParser().parse(path)
        assert config.ignore_case is False

    @pytest.mark.parametrize('value', ['yes', 'YES', 'true', 'True', '1', 'on', 'y'])
    def test_truthy_values_enable(self, tmp_path, value):
        path = _write_pattern([
            ['config:', 'ignore.case', value],
            ['var:', 'po.number', 'string', r'PO-\d+'],
        ], tmp_path)
        config, _, _ = PatternParser().parse(path)
        assert config.ignore_case is True

    @pytest.mark.parametrize('value', ['no', 'No', 'false', '0', 'off', 'n', ''])
    def test_falsy_values_disable(self, tmp_path, value):
        path = _write_pattern([
            ['config:', 'ignore.case', value],
            ['var:', 'po.number', 'string', r'PO-\d+'],
        ], tmp_path)
        config, _, _ = PatternParser().parse(path)
        assert config.ignore_case is False

    def test_table_inherits_global_ignore_case(self, tmp_path):
        path = _write_pattern([
            ['config:', 'ignore.case', 'yes'],
            ['START:'],
            ['table:*'],
            [None, 'HEADER:1', 'col_a'],
            [None, 'DATA:*', 'col_a'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        table = [s for s in seq if type(s).__name__ == 'TableInstruction'][0]
        assert table.config.ignore_case is True


# ── seek: instruction ─────────────────────────────────────────────────────────

class TestSeekInstruction:
    """Parser-level tests for the seek: cursor-repositioning instruction."""

    def test_seek_parses_to_seek_instruction(self, tmp_path):
        from grepxcel.models import SeekInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:G5'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        seek_instrs = [s for s in seq if isinstance(s, SeekInstruction)]
        assert len(seek_instrs) == 1
        assert seek_instrs[0].target == 'G5'

    def test_seek_target_is_uppercased(self, tmp_path):
        from grepxcel.models import SeekInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:g5'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        seek_instrs = [s for s in seq if isinstance(s, SeekInstruction)]
        assert seek_instrs[0].target == 'G5'

    def test_seek_invalid_address_raises_pattern_error(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:NOTACELL'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='seek:'):
            PatternParser().parse(path)

    def test_seek_empty_address_raises_pattern_error(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='seek:'):
            PatternParser().parse(path)

    def test_seek_resets_abs_ref_ordering(self, tmp_path):
        """seek:B2 after cell:K30 resets the forward-order check — must not raise."""
        path = _write_pattern([
            ['var:', 'late',  'string', '.*'],
            ['var:', 'early', 'string', '.*'],
            ['START:'],
            ['cell:K30', 'late'],   # sets last_abs_pos = K30 (row 30)
            ['seek:B2'],            # resets last_abs_pos to B2 (row 2) — no error
            ['cell:next', 'early'],
            ['END:'],
        ], tmp_path)
        from grepxcel.models import SeekInstruction
        _, _, seq = PatternParser().parse(path)
        assert any(isinstance(s, SeekInstruction) for s in seq)

    def test_seek_then_same_cell_abs_is_valid(self, tmp_path):
        """seek:I4 followed immediately by cell:I4 must parse without error.
        seek resets the ordering constraint entirely (last_abs_pos → None),
        so the first abs ref after seek is unchecked."""
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:I4'],
            ['cell:I4', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)  # must not raise
        from grepxcel.models import CellInstruction, SeekInstruction
        assert any(isinstance(s, SeekInstruction) for s in seq)
        assert any(isinstance(s, CellInstruction) and s.target == 'I4' for s in seq)

    def test_abs_ref_after_seek_no_parser_constraint(self, tmp_path):
        """After seek:, the abs-ref ordering constraint is fully reset to None.
        Backward abs refs are not rejected at parse time — the engine catches
        them at runtime. This allows seek:I4; cell:I4 and similar patterns."""
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['seek:B2'],
            ['cell:A1', 'y'],   # A1 < B2 in LR — allowed at parse time, caught at runtime
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)   # must NOT raise
        from grepxcel.models import CellInstruction
        abs_cells = [s for s in seq if isinstance(s, CellInstruction) and s.multiplicity == 'abs']
        assert any(c.target == 'A1' for c in abs_cells)

    def test_abs_ref_after_seek_forward_is_valid(self, tmp_path):
        """cell:B5 after seek:B2 is valid — B5 is forward of B2 in LR order."""
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:B2'],
            ['cell:B5', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)   # must not raise
        from grepxcel.models import CellInstruction
        abs_cells = [s for s in seq if isinstance(s, CellInstruction) and s.multiplicity == 'abs']
        assert len(abs_cells) == 1
        assert abs_cells[0].target == 'B5'


# ── pattern.version ───────────────────────────────────────────────────────────

class TestPatternVersion:
    _SEQ = [['var:', 'x', 'string', '.*'], ['START:'], ['cell:A1', 'x'], ['END:']]

    def test_absent_defaults_to_1(self, tmp_path):
        cfg, _, _ = PatternParser().parse(_write_pattern(self._SEQ, tmp_path))
        assert cfg.pattern_version == 1
        assert cfg.pattern_version_explicit is False

    def test_explicit_version_parsed(self, tmp_path):
        rows = [['config:', 'pattern.version', '1']] + self._SEQ
        cfg, _, _ = PatternParser().parse(_write_pattern(rows, tmp_path))
        assert cfg.pattern_version == 1
        assert cfg.pattern_version_explicit is True

    def test_version_too_new_raises(self, tmp_path):
        rows = [['config:', 'pattern.version', '999']] + self._SEQ
        with pytest.raises(PatternError, match='newer|upgrade|pattern.version'):
            PatternParser().parse(_write_pattern(rows, tmp_path))

    def test_invalid_version_raises(self, tmp_path):
        rows = [['config:', 'pattern.version', 'abc']] + self._SEQ
        with pytest.raises(PatternError, match='pattern.version'):
            PatternParser().parse(_write_pattern(rows, tmp_path))


# ── case-insensitive keywords ─────────────────────────────────────────────────

class TestCaseInsensitiveKeywords:
    """Structural keywords are case-insensitive (CELL:, Cell:, cell: all work).
    The keyword case must not matter; field names and addresses are unaffected."""

    def test_uppercase_cell_keyword(self, tmp_path):
        from grepxcel.models import CellInstruction
        path = _write_pattern([
            ['var:', 'totals.x', 'currency', '.*'],
            ['START:'],
            ['CELL:J59', 'totals.x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        abs_cells = [s for s in seq if isinstance(s, CellInstruction)]
        assert len(abs_cells) == 1
        assert abs_cells[0].target == 'J59'
        assert abs_cells[0].field == 'totals.x'

    def test_uppercase_cell_next(self, tmp_path):
        from grepxcel.models import CellInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['CELL:NEXT', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        cells = [s for s in seq if isinstance(s, CellInstruction)]
        assert cells[0].multiplicity == 'next'   # normalised to lowercase

    def test_mixed_case_seek(self, tmp_path):
        from grepxcel.models import SeekInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['Seek:G5'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert any(isinstance(s, SeekInstruction) and s.target == 'G5' for s in seq)

    def test_uppercase_dir(self, tmp_path):
        from grepxcel.models import DirectionInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['DIR:TD'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        assert any(isinstance(s, DirectionInstruction) and s.direction == 'TD' for s in seq)

    def test_uppercase_table_and_row_types(self, tmp_path):
        from grepxcel.models import TableInstruction
        path = _write_pattern([
            ['lbl:', 'h', 'string', 'Name'],
            ['var:', 'item.name', 'string', '.*'],
            ['START:'],
            ['TABLE:*'],
            ['', 'HEADER:1', 'h'],
            ['', 'data:*', 'item.name'],   # lowercase row type
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        tables = [s for s in seq if isinstance(s, TableInstruction)]
        assert len(tables) == 1
        row_types = {r.row_type for r in tables[0].rows}
        assert row_types == {'HEADER', 'DATA'}   # normalised to uppercase

    def test_uppercase_section_keywords(self, tmp_path):
        path = _write_pattern([
            ['CONFIG:', 'read.direction', 'LR'],
            ['VAR:', 'a', 'string', '.*'],
            ['LBL:', 'b', 'string', 'Label'],
            ['Start:'],
            ['cell:next', 'a'],
            ['End:'],
        ], tmp_path)
        cfg, defs, seq = PatternParser().parse(path)
        assert 'a' in defs and 'b' in defs
        assert len(seq) == 1

    def test_field_names_keep_their_case(self, tmp_path):
        """Only keywords are case-folded — field names stay exactly as written."""
        path = _write_pattern([
            ['var:', 'totalCost.projected', 'currency', '.*'],
            ['START:'],
            ['CELL:J59', 'totalCost.projected'],
            ['END:'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert 'totalCost.projected' in defs   # camelCase preserved


# ── unknown instruction rows are not silently dropped ─────────────────────────

class TestUnknownInstructionRow:
    """A non-empty, unrecognised row inside START: is a fatal error — never
    silently ignored (which would produce wrong/empty output)."""

    def test_typo_instruction_raises(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['celll:J5', 'x'],   # typo — extra l
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='[Uu]nrecognis|[Uu]nrecogniz|[Uu]nknown'):
            PatternParser().parse(path)

    def test_misspelled_seek_raises(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['sek:G5'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError):
            PatternParser().parse(path)

    def test_doc_row_inside_start_is_allowed(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['doc:', 'a separator comment'],
            ['cell:next', 'x'],
            ['DOC:', 'uppercase comment too'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)   # must not raise
        from grepxcel.models import CellInstruction
        assert any(isinstance(s, CellInstruction) for s in seq)

    def test_blank_rows_inside_start_allowed(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            [None, None],
            ['cell:next', 'x'],
            ['', ''],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)   # must not raise
        from grepxcel.models import CellInstruction
        assert len([s for s in seq if isinstance(s, CellInstruction)]) == 1


# ── dir: instruction ──────────────────────────────────────────────────────────

class TestDirectionInstruction:
    """Parser-level tests for the dir: scan-direction-switch instruction."""

    def test_dir_lr_parses(self, tmp_path):
        from grepxcel.models import DirectionInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['dir:LR'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        dirs = [s for s in seq if isinstance(s, DirectionInstruction)]
        assert len(dirs) == 1
        assert dirs[0].direction == 'LR'

    def test_dir_td_parses(self, tmp_path):
        from grepxcel.models import DirectionInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['dir:TD'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        dirs = [s for s in seq if isinstance(s, DirectionInstruction)]
        assert dirs[0].direction == 'TD'

    def test_dir_lowercase_is_uppercased(self, tmp_path):
        from grepxcel.models import DirectionInstruction
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['dir:td'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        dirs = [s for s in seq if isinstance(s, DirectionInstruction)]
        assert dirs[0].direction == 'TD'

    def test_dir_invalid_value_raises(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['dir:DIAGONAL'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='dir:'):
            PatternParser().parse(path)

    def test_dir_empty_value_raises(self, tmp_path):
        path = _write_pattern([
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['dir:'],
            ['cell:next', 'x'],
            ['END:'],
        ], tmp_path)
        with pytest.raises(PatternError, match='dir:'):
            PatternParser().parse(path)

    def test_dir_resets_abs_ref_ordering(self, tmp_path):
        """A dir: switch resets the abs-ref forward-order check (like seek:).
        cell:K30 then dir:TD then cell:A1 must parse without error."""
        from grepxcel.models import DirectionInstruction
        path = _write_pattern([
            ['var:', 'late',  'string', '.*'],
            ['var:', 'early', 'string', '.*'],
            ['START:'],
            ['cell:K30', 'late'],
            ['dir:TD'],
            ['cell:A1', 'early'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)   # must not raise
        assert any(isinstance(s, DirectionInstruction) for s in seq)

    def test_dir_changes_abs_ordering_frame(self, tmp_path):
        """After dir:TD, the abs-ref ordering is checked in TD order.
        cell:A5 then cell:A10 (forward in TD) must be valid."""
        path = _write_pattern([
            ['var:', 'a', 'string', '.*'],
            ['var:', 'b', 'string', '.*'],
            ['START:'],
            ['dir:TD'],
            ['cell:A5', 'a'],
            ['cell:A10', 'b'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)   # must not raise
        from grepxcel.models import CellInstruction
        abs_cells = [s for s in seq if isinstance(s, CellInstruction) and s.multiplicity == 'abs']
        assert len(abs_cells) == 2


# ── lbl.match mode ───────────────────────────────────────────────────────────

class TestLblMatchMode:
    def _parse(self, rows, tmp_path):
        return PatternParser().parse(_write_pattern(rows, tmp_path))

    def _base(self, lbl_pattern, lbl_suffix='lbl:'):
        return [
            [lbl_suffix, 'h', 'string', lbl_pattern],
            ['var:', 'v', 'string', '.*'],
            ['START:'],
            ['cell:1', 'h'],
            ['cell:1', 'v'],
            ['END:'],
        ]

    def test_default_lbl_match_is_literal(self, tmp_path):
        cfg, _, _ = self._parse(self._base('Hello'), tmp_path)
        assert cfg.lbl_match == 'literal'

    def test_lbl_field_has_no_override_by_default(self, tmp_path):
        _, defs, _ = self._parse(self._base('Hello'), tmp_path)
        assert defs['h'].lbl_match is None

    def test_lbl_literal_suffix_sets_override(self, tmp_path):
        _, defs, _ = self._parse(self._base('Hello', 'lbl:literal'), tmp_path)
        assert defs['h'].lbl_match == 'literal'

    def test_lbl_glob_suffix_sets_override(self, tmp_path):
        _, defs, _ = self._parse(self._base('Hello *', 'lbl:glob'), tmp_path)
        assert defs['h'].lbl_match == 'glob'

    def test_lbl_regexp_suffix_sets_override(self, tmp_path):
        _, defs, _ = self._parse(self._base(r'Hello \w+', 'lbl:regexp'), tmp_path)
        assert defs['h'].lbl_match == 'regexp'

    def test_lbl_unknown_suffix_raises(self, tmp_path):
        with pytest.raises(PatternError, match='Unknown modifier'):
            self._parse(self._base('Hello', 'lbl:fuzzy'), tmp_path)

    def test_config_lbl_match_sets_global(self, tmp_path):
        rows = [
            ['config:', 'lbl.match', 'glob'],
            ['lbl:', 'h', 'string', 'Hello *'],
            ['var:', 'v', 'string', '.*'],
            ['START:'], ['cell:1', 'h'], ['cell:1', 'v'], ['END:'],
        ]
        cfg, _, _ = self._parse(rows, tmp_path)
        assert cfg.lbl_match == 'glob'

    def test_config_lbl_match_invalid_raises(self, tmp_path):
        rows = [
            ['config:', 'lbl.match', 'exact'],
            ['var:', 'v', 'string', '.*'],
            ['START:'], ['cell:1', 'v'], ['END:'],
        ]
        with pytest.raises(PatternError, match='Invalid lbl.match'):
            self._parse(rows, tmp_path)

    def test_literal_mode_skips_regex_safety_check(self, tmp_path):
        """Metacharacter-heavy label patterns are accepted without error in literal mode."""
        self._parse(self._base('Term (months):', 'lbl:literal'), tmp_path)  # no raise

    def test_glob_mode_skips_regex_safety_check(self, tmp_path):
        """Glob patterns with * are accepted without regex safety check."""
        self._parse(self._base('Invoice *', 'lbl:glob'), tmp_path)  # no raise

    def test_regexp_mode_enforces_regex_safety(self, tmp_path):
        """In regexp mode the ReDoS guard still applies."""
        with pytest.raises((PatternError, Exception), match='[Rr]e[Dd]o[Ss]|catastrophic|backtrack'):
            self._parse(self._base(r'(a+)+$', 'lbl:regexp'), tmp_path)

    # ── New order-independent modifier syntax ─────────────────────────────────

    def test_lbl_re_alias_normalises_to_regexp(self, tmp_path):
        """lbl:re is a short alias for lbl:regexp."""
        _, defs, _ = self._parse(self._base(r'Hello \w+', 'lbl:re'), tmp_path)
        assert defs['h'].lbl_match == 'regexp'

    def test_lbl_not_null_sets_required(self, tmp_path):
        _, defs, _ = self._parse(self._base('Hello', 'lbl:not-null'), tmp_path)
        assert defs['h'].required is True
        assert defs['h'].lbl_match is None  # mode stays default (literal)

    def test_lbl_not_empty_synonym(self, tmp_path):
        _, defs, _ = self._parse(self._base('Hello', 'lbl:not-empty'), tmp_path)
        assert defs['h'].required is True

    def test_lbl_glob_not_null_order_independent(self, tmp_path):
        """'lbl:glob:not-null' and 'lbl:not-null:glob' both set glob + required."""
        _, defs1, _ = self._parse(self._base('Hello*', 'lbl:glob:not-null'), tmp_path)
        _, defs2, _ = self._parse(self._base('Hello*', 'lbl:not-null:glob'), tmp_path)
        assert defs1['h'].lbl_match == 'glob' and defs1['h'].required is True
        assert defs2['h'].lbl_match == 'glob' and defs2['h'].required is True

    def test_lbl_re_not_empty(self, tmp_path):
        _, defs, _ = self._parse(self._base(r'Hello \w+', 'lbl:re:not-empty'), tmp_path)
        assert defs['h'].lbl_match == 'regexp'
        assert defs['h'].required is True

    def test_lbl_duplicate_mode_raises(self, tmp_path):
        with pytest.raises(PatternError, match='Duplicate mode'):
            self._parse(self._base('x', 'lbl:glob:literal'), tmp_path)

    def test_var_not_null(self, tmp_path):
        rows = [
            ['var:not-null', 'amount', 'currency', r'\d+\.?\d*'],
            ['START:'], ['cell:1', 'amount'], ['END:'],
        ]
        _, defs, _ = self._parse(rows, tmp_path)
        assert defs['amount'].required is True
        assert defs['amount'].var_mode is None  # still default (regexp)

    def test_var_glob_mode(self, tmp_path):
        rows = [
            ['var:glob', 'code', 'string', 'PROD-*'],
            ['START:'], ['cell:1', 'code'], ['END:'],
        ]
        _, defs, _ = self._parse(rows, tmp_path)
        assert defs['code'].var_mode == 'glob'
        assert defs['code'].required is False

    def test_var_glob_not_null(self, tmp_path):
        rows = [
            ['var:not-null:glob', 'code', 'string', 'PROD-*'],
            ['START:'], ['cell:1', 'code'], ['END:'],
        ]
        _, defs, _ = self._parse(rows, tmp_path)
        assert defs['code'].var_mode == 'glob'
        assert defs['code'].required is True

    def test_var_re_explicit(self, tmp_path):
        rows = [
            ['var:re', 'amount', 'currency', r'\d+'],
            ['START:'], ['cell:1', 'amount'], ['END:'],
        ]
        _, defs, _ = self._parse(rows, tmp_path)
        assert defs['amount'].var_mode == 'regexp'

    def test_var_regexp_not_empty_backward_compat(self, tmp_path):
        rows = [
            ['var:regexp:not-empty', 'v', 'string', r'[A-Z]+'],
            ['START:'], ['cell:1', 'v'], ['END:'],
        ]
        _, defs, _ = self._parse(rows, tmp_path)
        assert defs['v'].var_mode == 'regexp'
        assert defs['v'].required is True

    def test_var_glob_skips_regex_safety_check(self, tmp_path):
        """glob patterns are not compiled as regexes — no safety check."""
        rows = [
            ['var:glob', 'code', 'string', '(((bad-regex'],
            ['START:'], ['cell:1', 'code'], ['END:'],
        ]
        self._parse(rows, tmp_path)  # no raise

    def test_var_literal_skips_regex_safety_check(self, tmp_path):
        rows = [
            ['var:literal', 'status', 'string', 'Active (primary)'],
            ['START:'], ['cell:1', 'status'], ['END:'],
        ]
        self._parse(rows, tmp_path)  # no raise

    def test_def_alias_still_works(self, tmp_path):
        """def: is a backward-compat alias for var: and keeps working."""
        rows = [
            ['def:', 'v', 'string', '.*'],
            ['START:'], ['cell:1', 'v'], ['END:'],
        ]
        _, defs, _ = self._parse(rows, tmp_path)
        assert defs['v'].role == 'var'
