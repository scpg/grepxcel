"""
Unit tests for the CSV pattern-source front-end.

The pattern parser reads .xlsx and .csv into one common 2D-grid IR, then runs a
single semantic parser over the grid. These tests cover the CSV reader's grid
production (empty-field → None, BOM handling, comma-in-regex quoting) and its
content guards (formula rejection, length cap), plus the format-aware
validate_pattern_file dispatcher.

Parity with the xlsx reader across all real fixtures is verified separately in
tests/integration/test_csv_pattern_parity.py.
"""

import csv
import os
import tempfile

import pytest

from grepxcel.pattern_parser import PatternParser, PatternError
from grepxcel.security import SecurityError, validate_pattern_file


# ── helpers ─────────────────────────────────────────────────────────────────

def _write_csv(rows: list, tmp_path, name: str = 'pattern.csv',
               encoding: str = 'utf-8') -> str:
    """Write rows (list of lists) to a CSV file and return its path."""
    path = str(tmp_path / name)
    with open(path, 'w', newline='', encoding=encoding) as fh:
        csv.writer(fh).writerows(rows)
    return path


def _write_text(text: str, tmp_path, name: str = 'pattern.csv',
                encoding: str = 'utf-8') -> str:
    path = str(tmp_path / name)
    with open(path, 'w', encoding=encoding) as fh:
        fh.write(text)
    return path


# ── CSV reader → grid → semantic parse ──────────────────────────────────────

class TestCsvPatternParse:
    def test_basic_lbl_var_parse(self, tmp_path):
        path = _write_csv([
            ['lbl:', 'inv_label', 'string', 'Invoice No:'],
            ['var:', 'inv.number', 'string', r'INV-\d+'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['inv_label'].role == 'lbl'
        assert defs['inv.number'].role == 'var'
        assert defs['inv.number'].regex == r'INV-\d+'

    def test_config_rows_parsed(self, tmp_path):
        path = _write_csv([
            ['config:', 'read.direction', 'TD'],
            ['config:', 'currency.sign', '$'],
            ['var:', 'x', 'integer', r'\d+'],
        ], tmp_path)
        cfg, _, _ = PatternParser().parse(path)
        assert cfg.read_direction == 'TD'
        assert cfg.currency_sign == '$'

    def test_empty_field_becomes_none_so_table_template_parses(self, tmp_path):
        """A blank column A marks a table-template row — empty CSV field must
        be None, exactly like a blank xlsx cell, or the table block breaks."""
        path = _write_csv([
            ['lbl:', 'col_item', 'string', 'Item'],
            ['var:', 'line.item', 'string', '.*'],
            ['var:', 'line.qty', 'integer', r'\d+'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'col_item'],          # blank col A → template row
            ['', 'DATA:*', 'line.item', 'line.qty'],
            ['END:'],
        ], tmp_path)
        _, _, seq = PatternParser().parse(path)
        # One TableInstruction with HEADER + DATA template rows
        from grepxcel.models import TableInstruction
        tables = [s for s in seq if isinstance(s, TableInstruction)]
        assert len(tables) == 1
        row_types = [r.row_type for r in tables[0].rows]
        assert 'HEADER' in row_types and 'DATA' in row_types

    def test_cell_sequence_parses(self, tmp_path):
        path = _write_csv([
            ['lbl:', 'po_label', 'string', 'PO:'],
            ['var:', 'po.number', 'string', r'PO-\d+'],
            ['START:'],
            ['cell:next', 'po_label'],
            ['cell:next', 'po.number'],
            ['cell:B5', 'IGNORE'],
            ['END:'],
        ], tmp_path)
        from grepxcel.models import CellInstruction
        _, _, seq = PatternParser().parse(path)
        cells = [s for s in seq if isinstance(s, CellInstruction)]
        assert [c.field for c in cells] == ['po_label', 'po.number', 'IGNORE']
        assert cells[2].multiplicity == 'abs' and cells[2].target == 'B5'

    def test_comma_in_quoted_regex_preserved(self, tmp_path):
        """A regex containing a comma must survive CSV round-trip when quoted.
        csv.writer quotes it automatically; the reader must unquote it."""
        path = _write_csv([
            ['var:', 'line.qty', 'integer', r'\d{1,3}'],
        ], tmp_path)
        _, defs, _ = PatternParser().parse(path)
        assert defs['line.qty'].regex == r'\d{1,3}'

    def test_bom_is_stripped(self, tmp_path):
        """Excel's 'Save as CSV' writes a UTF-8 BOM; utf-8-sig must strip it so
        the first keyword is 'config:' and not '\\ufeffconfig:'."""
        path = _write_csv([
            ['config:', 'read.direction', 'LR'],
            ['var:', 'x', 'string', '.*'],
        ], tmp_path, encoding='utf-8-sig')
        cfg, defs, _ = PatternParser().parse(path)
        assert cfg.read_direction == 'LR'
        assert 'x' in defs

    def test_leading_equals_rejected(self, tmp_path):
        path = _write_csv([
            ['var:', 'x', 'string', '=SUM(A1:A2)'],
        ], tmp_path)
        with pytest.raises(SecurityError, match='Formulas are not allowed'):
            PatternParser().parse(path)

    def test_overlong_cell_rejected(self, tmp_path):
        path = _write_csv([
            ['var:', 'x', 'string', 'a' * 1001],
        ], tmp_path)
        with pytest.raises(SecurityError, match='too long'):
            PatternParser().parse(path)

    def test_invalid_utf8_rejected(self, tmp_path):
        path = str(tmp_path / 'bad.csv')
        with open(path, 'wb') as fh:
            fh.write(b'var:,x,string,\xff\xfe\x00bad')
        with pytest.raises(SecurityError, match='not valid UTF-8'):
            PatternParser().parse(path)

    def test_redos_regex_rejected_in_csv(self, tmp_path):
        """The per-regex ReDoS guard still runs for CSV-sourced patterns."""
        path = _write_csv([
            ['var:', 'x', 'string', '(a+)+'],
        ], tmp_path)
        with pytest.raises(SecurityError, match='[Uu]nsafe regex'):
            PatternParser().parse(path)

    def test_blank_lines_ignored(self, tmp_path):
        path = _write_text(
            'config:,read.direction,LR\n'
            '\n'
            'var:,x,string,.*\n'
            '\n',
            tmp_path,
        )
        cfg, defs, _ = PatternParser().parse(path)
        assert cfg.read_direction == 'LR'
        assert 'x' in defs

    def test_semicolon_delimiter_autodetected(self, tmp_path):
        """Excel in many locales exports CSV with ';' (comma is the decimal
        separator there). The reader must sniff the delimiter, not assume ','."""
        path = _write_text(
            'config:;read.direction;LR\n'
            'lbl:;po_label;string;PO:\n'
            'var:;po.number;string;PO-\\d+\n'
            'START:\n'
            'cell:next;po_label\n'
            'cell:next;po.number\n'
            'END:\n',
            tmp_path,
        )
        cfg, defs, seq = PatternParser().parse(path)
        assert cfg.read_direction == 'LR'
        assert defs['po.number'].regex == r'PO-\d+'
        assert len(seq) == 2          # the START: block is seen, not collapsed

    def test_tab_delimiter_autodetected(self, tmp_path):
        path = _write_text(
            'config:\tread.direction\tLR\n'
            'var:\tx\tstring\t.*\n'
            'START:\n'
            'cell:next\tx\n'
            'END:\n',
            tmp_path,
        )
        cfg, defs, seq = PatternParser().parse(path)
        assert cfg.read_direction == 'LR'
        assert 'x' in defs and len(seq) == 1


# ── validate_pattern_file dispatcher ────────────────────────────────────────

class TestValidatePatternFile:
    def test_csv_passes_text_checks(self, tmp_path):
        path = _write_csv([['var:', 'x', 'string', '.*']], tmp_path)
        validate_pattern_file(path)  # must not raise

    def test_csv_missing_file_rejected(self, tmp_path):
        with pytest.raises(SecurityError, match='File not found'):
            validate_pattern_file(str(tmp_path / 'nope.csv'))

    def test_csv_oversized_rejected(self, tmp_path):
        path = _write_csv([['var:', 'x', 'string', '.*']], tmp_path)
        with pytest.raises(SecurityError, match='exceeds'):
            validate_pattern_file(path, max_file_mb=0)

    def test_xlsx_delegates_to_full_validation(self):
        """A real fixture pattern.xlsx must pass the full xlsx validation path."""
        xlsx = os.path.join(
            os.path.dirname(__file__), '..', 'fixtures',
            '01_simple_invoice', 'pattern.xlsx',
        )
        validate_pattern_file(xlsx)  # must not raise

    def test_unknown_extension_rejected(self, tmp_path):
        path = str(tmp_path / 'pattern.txt')
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write('var:,x,string,.*')
        with pytest.raises(SecurityError, match='Unsupported pattern file type'):
            validate_pattern_file(path)
