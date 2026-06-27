"""Unit tests for the grepxcel wizard command."""

import csv
import datetime
import io
import os

import pytest

from grepxcel.wizard import (
    WizardState,
    _propose_type,
    _slugify,
    _col_label,
    _write_pattern,
)


# ── _propose_type ─────────────────────────────────────────────────────────────

class TestProposeType:
    def test_none_is_skip(self):
        assert _propose_type(None) == 'skip'

    def test_empty_string_is_skip(self):
        assert _propose_type('') == 'skip'

    def test_blank_string_is_skip(self):
        assert _propose_type('   ') == 'skip'

    def test_short_string_is_label(self):
        assert _propose_type('Invoice Number') == 'label'

    def test_exactly_40_chars_is_label(self):
        assert _propose_type('A' * 40) == 'label'

    def test_41_chars_is_var_string(self):
        assert _propose_type('A' * 41) == 'var:string'

    def test_long_string_is_var_string(self):
        assert _propose_type('This is a very long description that exceeds forty chars') == 'var:string'

    def test_integer_is_var_integer(self):
        assert _propose_type(42) == 'var:integer'

    def test_whole_float_is_var_integer(self):
        assert _propose_type(3.0) == 'var:integer'

    def test_fractional_float_is_var_currency(self):
        assert _propose_type(3.14) == 'var:currency'

    def test_date_is_var_date(self):
        assert _propose_type(datetime.date(2026, 1, 1)) == 'var:date'

    def test_datetime_is_var_datetime(self):
        assert _propose_type(datetime.datetime(2026, 1, 1, 12, 0)) == 'var:datetime'


# ── _slugify ──────────────────────────────────────────────────────────────────

class TestSlugify:
    def test_lower_and_replace_spaces(self):
        assert _slugify('Invoice Number') == 'invoice_number'

    def test_strips_special_chars(self):
        assert _slugify('Total (€)') == 'total'

    def test_already_slug(self):
        assert _slugify('amount_net') == 'amount_net'


# ── _col_label ────────────────────────────────────────────────────────────────

class TestColLabel:
    def test_col_1_is_A(self):
        assert _col_label(1) == 'A'

    def test_col_26_is_Z(self):
        assert _col_label(26) == 'Z'

    def test_col_27_is_AA(self):
        assert _col_label(27) == 'AA'

    def test_col_28_is_AB(self):
        assert _col_label(28) == 'AB'


# ── _write_pattern ────────────────────────────────────────────────────────────

class TestWritePattern:
    def _read_csv(self, path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.reader(f))

    def test_config_rows_written(self, tmp_path):
        state = WizardState(direction='LR', currency_sign='€')
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['config:', 'read.direction', 'LR'] in rows
        assert ['config:', 'currency.sign', '€'] in rows

    def test_ignore_case_written_when_true(self, tmp_path):
        state = WizardState(direction='TD', ignore_case=True, currency_sign='$')
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['config:', 'ignore.case', 'yes'] in rows

    def test_ignore_case_absent_when_false(self, tmp_path):
        state = WizardState(ignore_case=False)
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert not any(r[:2] == ['config:', 'ignore.case'] for r in rows)

    def test_lbl_def_written(self, tmp_path):
        state = WizardState()
        state.lbl_defs.append(('inv_number_lbl', 'string', 'Invoice Number'))
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['lbl:', 'inv_number_lbl', 'string', 'Invoice Number'] in rows

    def test_var_def_written(self, tmp_path):
        state = WizardState()
        state.var_defs.append(('inv.number', 'string', '[A-Z]{2}[0-9]+'))
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['var:', 'inv.number', 'string', '[A-Z]{2}[0-9]+'] in rows

    def test_start_end_always_present(self, tmp_path):
        state = WizardState()
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        first_col = [r[0] for r in rows]
        assert 'START:' in first_col
        assert 'END:' in first_col

    def test_scalar_cell_instruction(self, tmp_path):
        state = WizardState()
        state.body_rows.append(['cell:1', 'inv.number'])
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['cell:1', 'inv.number'] in rows

    def test_table_block_written(self, tmp_path):
        state = WizardState()
        state.body_rows.append(['table:*'])
        state.body_rows.append(['', 'HEADER:1', 'item.name', 'item.qty'])
        state.body_rows.append(['', 'DATA:*', 'item.name', 'item.qty'])
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['table:*'] in rows
        assert ['', 'HEADER:1', 'item.name', 'item.qty'] in rows
        assert ['', 'DATA:*', 'item.name', 'item.qty'] in rows

    def test_footer_row_written(self, tmp_path):
        state = WizardState()
        state.body_rows.append(['table:1'])
        state.body_rows.append(['', 'HEADER:1', 'label', 'total'])
        state.body_rows.append(['', 'DATA:*', 'label', 'total'])
        state.body_rows.append(['', 'FOOTER:1', 'footer.label', 'footer.total'])
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        rows = self._read_csv(out)
        assert ['', 'FOOTER:1', 'footer.label', 'footer.total'] in rows

    def test_output_is_valid_utf8_csv(self, tmp_path):
        state = WizardState(currency_sign='€')
        state.lbl_defs.append(('lbl', 'string', 'Értéke'))
        out = str(tmp_path / 'p.csv')
        _write_pattern(state, out)
        content = open(out, encoding='utf-8').read()
        assert 'Értéke' in content


# ── WizardState defaults ──────────────────────────────────────────────────────

class TestWizardState:
    def test_defaults(self):
        s = WizardState()
        assert s.direction == 'LR'
        assert s.ignore_case is False
        assert s.currency_sign == '€'
        assert s.lbl_defs == []
        assert s.var_defs == []
        assert s.body_rows == []
