"""Unit tests for --format csv output."""

import csv
import io
import json
import os
import sys

import pytest

from grepxcel.cli import main
from grepxcel.csv_writer import nested_to_csv, count_table_instructions
from tests.conftest import find_pattern_xlsx

_INVOICE_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '01_simple_invoice')
_INVOICE_PAT = find_pattern_xlsx(_INVOICE_FIX)
_INVOICE_DATA = os.path.join(_INVOICE_FIX, 'data.xlsx')

_CATALOG_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '02_product_catalog')
_CATALOG_PAT = find_pattern_xlsx(_CATALOG_FIX)
_CATALOG_DATA = os.path.join(_CATALOG_FIX, 'data.xlsx')

_EXPENSE_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '05_expense_report')
_EXPENSE_PAT = find_pattern_xlsx(_EXPENSE_FIX)
_EXPENSE_DATA = os.path.join(_EXPENSE_FIX, 'data.xlsx')


def _run_cli(*args):
    """Run CLI and capture stdout + return code."""
    old_stdout = sys.stdout
    sys.stdout = buf = io.StringIO()
    try:
        rc = main(list(args))
    except SystemExit as e:
        rc = e.code
    finally:
        sys.stdout = old_stdout
    return rc, buf.getvalue()


# ── nested_to_csv unit tests ────────────────────────────────────────────────

class TestNestedToCsv:
    def test_scalar_only(self):
        data = {'inv': {'number': 'AB123', 'date': '2026-01-01'}}
        out = nested_to_csv(data)
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        assert len(rows) == 1
        assert rows[0]['inv.number'] == 'AB123'
        assert rows[0]['inv.date'] == '2026-01-01'

    def test_table_only(self):
        data = {
            'items': [
                {'data': [
                    {'name': 'A', 'qty': 1},
                    {'name': 'B', 'qty': 2},
                ]},
            ],
        }
        out = nested_to_csv(data)
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        assert len(rows) == 2
        assert rows[0]['name'] == 'A'
        assert rows[1]['qty'] == '2'

    def test_mixed_scalar_and_table(self):
        data = {
            'vendor': {'name': 'Acme'},
            'items': [
                {'data': [
                    {'item': 'Pen', 'price': 2},
                    {'item': 'Ink', 'price': 5},
                ]},
            ],
        }
        out = nested_to_csv(data)
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        assert len(rows) == 2
        assert all(r['vendor.name'] == 'Acme' for r in rows)
        assert rows[0]['item'] == 'Pen'
        assert rows[1]['item'] == 'Ink'

    def test_source_excluded(self):
        data = {
            'items': [
                {
                    '_source': {'sheet': 'Sheet1', 'ref': 'A1:D5'},
                    'data': [{'name': 'A'}],
                },
            ],
        }
        out = nested_to_csv(data)
        assert '_source' not in out

    def test_multi_instance_concatenated(self):
        data = {
            'items': [
                {'data': [{'x': 1}]},
                {'data': [{'x': 2}, {'x': 3}]},
            ],
        }
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        assert len(rows) == 3

    def test_empty_result(self):
        out = nested_to_csv({})
        assert out.strip() == ''


# ── count_table_instructions ─────────────────────────────────────────────────

class TestCountTableInstructions:
    def test_zero_for_scalar_pattern(self):
        assert count_table_instructions(_INVOICE_PAT) == 0

    def test_one_for_single_table_pattern(self):
        assert count_table_instructions(_CATALOG_PAT) == 1


# ── CLI integration: single-table pattern produces valid CSV ─────────────────

class TestSingleTable:
    def test_csv_output_is_valid(self):
        rc, out = _run_cli('extract', '-p', _CATALOG_PAT, _CATALOG_DATA,
                           '--format', 'csv')
        assert rc == 0
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        assert len(rows) >= 3

    def test_csv_has_header_row(self):
        rc, out = _run_cli('extract', '-p', _CATALOG_PAT, _CATALOG_DATA,
                           '--format', 'csv')
        lines = out.strip().split('\n')
        assert len(lines) >= 2
        assert ',' in lines[0]

    def test_csv_data_matches_json_row_count(self):
        _, json_out = _run_cli('extract', '-p', _CATALOG_PAT, _CATALOG_DATA,
                               '--format', 'nested')
        _, csv_out = _run_cli('extract', '-p', _CATALOG_PAT, _CATALOG_DATA,
                              '--format', 'csv')
        json_data = json.loads(json_out)
        csv_rows = list(csv.DictReader(io.StringIO(csv_out)))

        table_key = [k for k in json_data if isinstance(json_data[k], list)][0]
        json_rows = []
        for instance in json_data[table_key]:
            json_rows.extend(instance.get('data', []))

        assert len(csv_rows) == len(json_rows)


# ── scalar-only pattern produces single-row CSV ─────────────────────────────

class TestScalarOnly:
    def test_scalar_only_produces_csv(self):
        rc, out = _run_cli('extract', '-p', _INVOICE_PAT, _INVOICE_DATA,
                           '--format', 'csv')
        assert rc == 0
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        assert len(rows) == 1

    def test_scalar_columns_are_dot_flattened(self):
        rc, out = _run_cli('extract', '-p', _INVOICE_PAT, _INVOICE_DATA,
                           '--format', 'csv')
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        cols = set(rows[0].keys())
        assert any('.' in c for c in cols)


# ── mixed scalar + table: scalars denormalized as extra columns ──────────────

class TestMixedScalarTable:
    def test_scalars_denormalized(self):
        rc, out = _run_cli('extract', '-p', _EXPENSE_PAT, _EXPENSE_DATA,
                           '--format', 'csv')
        assert rc == 0
        reader = csv.DictReader(io.StringIO(out))
        rows = list(reader)
        assert len(rows) >= 3
        cols = set(rows[0].keys())
        assert any('emp.' in c for c in cols)


# ── multi-table guard ────────────────────────────────────────────────────────

class TestMultiTableGuard:
    def _make_multi_table_pattern(self, tmp_path):
        """Create a CSV pattern with two table: blocks."""
        pat = tmp_path / 'pattern.csv'
        pat.write_text(
            'lbl:,col_name,string,Name,,\n'
            'var:,items.name,string,.+,,\n'
            'lbl:,col_price,string,Price,,\n'
            'var:,prices.value,currency,.*,,\n'
            'START:,,,,,\n'
            'table:*,,,,,\n'
            ',HEADER:1,col_name,,,\n'
            ',DATA:*,items.name,,,\n'
            'table:*,,,,,\n'
            ',HEADER:1,col_price,,,\n'
            ',DATA:*,prices.value,,,\n'
            'END:,,,,,\n',
            encoding='utf-8',
        )
        return str(pat)

    def test_multi_table_refused(self, tmp_path):
        pat = self._make_multi_table_pattern(tmp_path)
        rc, _ = _run_cli('extract', '-p', pat, _CATALOG_DATA, '--format', 'csv')
        assert rc == 2

    def test_error_message_mentions_json(self, tmp_path, capsys):
        pat = self._make_multi_table_pattern(tmp_path)
        _run_cli('extract', '-p', pat, _CATALOG_DATA, '--format', 'csv')
        captured = capsys.readouterr()
        assert 'json' in captured.err.lower() or 'nested' in captured.err.lower()


# ── --all-sheets guard ───────────────────────────────────────────────────────

class TestAllSheetsGuard:
    def test_csv_refuses_all_sheets(self):
        rc, _ = _run_cli('extract', '-p', _INVOICE_PAT, _INVOICE_DATA,
                         '--format', 'csv', '--all-sheets')
        assert rc == 2

    def test_error_mentions_sheet_option(self, capsys):
        _run_cli('extract', '-p', _INVOICE_PAT, _INVOICE_DATA,
                 '--format', 'csv', '--all-sheets')
        captured = capsys.readouterr()
        assert '--sheet' in captured.err or 'nested' in captured.err.lower()


# ── _meta exclusion ──────────────────────────────────────────────────────────

class TestMetaExclusion:
    def test_meta_not_in_csv_columns(self):
        rc, out = _run_cli('extract', '-p', _INVOICE_PAT, _INVOICE_DATA,
                           '--format', 'csv', '--meta')
        assert rc == 0
        header = out.strip().split('\n')[0]
        assert '_meta' not in header
        assert 'run_id' not in header

    def test_meta_excluded_unit(self):
        from grepxcel.csv_writer import nested_to_csv
        data = {
            'inv': {'number': 'X1'},
            '_meta': {'run_id': 'abc', 'stats': {'errors': 0}},
        }
        out = nested_to_csv(data)
        assert '_meta' not in out
        assert 'run_id' not in out
        assert 'X1' in out


# ── output to file ───────────────────────────────────────────────────────────

class TestOutputFile:
    def test_csv_writes_to_file(self, tmp_path):
        rc, _ = _run_cli('extract', '-p', _CATALOG_PAT, _CATALOG_DATA,
                         '--format', 'csv', '-o', str(tmp_path))
        assert rc == 0
        csv_files = list(tmp_path.glob('*.csv'))
        assert len(csv_files) == 1
        content = csv_files[0].read_text(encoding='utf-8')
        rows = list(csv.DictReader(io.StringIO(content)))
        assert len(rows) >= 3
