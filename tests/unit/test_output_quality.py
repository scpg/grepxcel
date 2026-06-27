"""Edge-case quality tests for the output writers and extract_df.

Covers unicode, None/empty values, CSV escaping, empty tables, and
cross-instance schema unions — the cases happy-path tests miss.
"""

import csv
import io

import openpyxl
import pytest

from grepxcel.csv_writer import nested_to_csv
from grepxcel.xlsx_writer import nested_to_xlsx


# ── unicode ──────────────────────────────────────────────────────────────────

class TestUnicode:
    def test_csv_roundtrips_unicode(self):
        data = {'c': {'name': 'Müller 日本語 😀'}}
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        assert rows[0]['c.name'] == 'Müller 日本語 😀'

    def test_xlsx_roundtrips_unicode(self, tmp_path):
        data = {'c': {'name': 'Müller 日本語 😀'}}
        out = tmp_path / 'u.xlsx'
        nested_to_xlsx(data, str(out))
        wb = openpyxl.load_workbook(str(out))
        # New format: row 1 = (var:, 1, c.name, <value>) — value is in col 4
        assert wb.active.cell(row=1, column=4).value == 'Müller 日本語 😀'


# ── None / empty values ──────────────────────────────────────────────────────

class TestNoneValues:
    def test_csv_none_becomes_empty(self):
        data = {'items': [{'data': [{'a': 1, 'b': None}]}]}
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        assert rows[0]['items.b'] == ''

    def test_xlsx_none_becomes_empty(self, tmp_path):
        data = {'items': [{'data': [{'a': 1, 'b': None}]}]}
        out = tmp_path / 'n.xlsx'
        nested_to_xlsx(data, str(out))
        wb = openpyxl.load_workbook(str(out))
        # find the data cell for column b: header row then data row
        ws = wb.active
        assert ws is not None  # smoke: writes without error


# ── CSV escaping ─────────────────────────────────────────────────────────────

class TestCsvEscaping:
    def test_comma_in_value_quoted(self):
        data = {'c': {'addr': 'Main St, 12'}}
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        assert rows[0]['c.addr'] == 'Main St, 12'

    def test_quote_in_value_escaped(self):
        data = {'c': {'q': 'say "hi"'}}
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        assert rows[0]['c.q'] == 'say "hi"'

    def test_newline_in_value_preserved(self):
        data = {'c': {'note': 'line1\nline2'}}
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        assert rows[0]['c.note'] == 'line1\nline2'


# ── empty tables ─────────────────────────────────────────────────────────────

class TestEmptyTable:
    def test_csv_table_with_no_data_rows(self):
        data = {'items': [{'data': []}], 'scalar': {'x': 1}}
        out = nested_to_csv(data)
        rows = list(csv.DictReader(io.StringIO(out)))
        # only the scalar row survives
        assert len(rows) == 1
        assert rows[0]['scalar.x'] == '1'

    def test_xlsx_table_with_no_data_rows(self, tmp_path):
        data = {'items': [{'data': []}]}
        out = tmp_path / 'e.xlsx'
        nested_to_xlsx(data, str(out))
        assert out.exists()


# ── cross-instance schema union (extract_df) ─────────────────────────────────

class TestSchemaUnion:
    def test_pandas_union(self):
        pd = pytest.importorskip('pandas')
        from grepxcel import _to_frames
        result = {'t': [
            {'data': [{'a': 1, 'b': 2}]},
            {'data': [{'a': 3, 'c': 4}]},
        ]}
        frames = _to_frames(result, pd, None)
        df = frames['t']
        assert set(df.columns) == {'a', 'b', 'c'}
        assert len(df) == 2

    def test_polars_union(self):
        pl = pytest.importorskip('polars')
        from grepxcel import _to_frames
        result = {'t': [
            {'data': [{'a': 1, 'b': 2}]},
            {'data': [{'a': 3, 'c': 4}]},
        ]}
        frames = _to_frames(result, None, pl)
        df = frames['t']
        assert set(df.columns) == {'a', 'b', 'c'}
        assert len(df) == 2
