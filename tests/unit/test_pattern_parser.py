"""
Unit tests for PatternParser: var:, lbl:, doc:, def: (alias), config,
cell/table sequence, and security validation.
"""

from pathlib import Path

import openpyxl
import pytest

from engine.pattern_parser import PatternParser, PatternError
from engine.security import SecurityError


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
        from engine.models import TableInstruction
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
        from engine.models import CellInstruction, TableInstruction
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

    def test_numeric_value_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 42
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='plain text'):
            PatternParser().parse(path)

    def test_boolean_value_rejected(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = True
        path = str(tmp_path / 'bad.xlsx')
        wb.save(path)
        with pytest.raises(SecurityError, match='plain text'):
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
