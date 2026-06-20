"""Tests for `grepxcel schema` — JSON Schema generation from pattern files."""
import io
import json
import os

import openpyxl
import pytest

from grepxcel.schema import generate_schema
from grepxcel.pattern_parser import PatternParser


def _write(rows, tmp_path, name='pattern.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    path = str(tmp_path / name)
    wb.save(path)
    return path


# ─── Basic structure ────────────────────────────────────────────────────────

class TestBasicSchema:
    def test_returns_valid_json_schema(self, tmp_path):
        path = _write([
            ['var:', 'name', 'string', '.*'],
            ['START:'], ['cell:A1', 'name'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert schema['type'] == 'object'
        assert '$schema' in schema

    def test_scalar_string_field(self, tmp_path):
        path = _write([
            ['var:', 'title', 'string', '.*'],
            ['START:'], ['cell:A1', 'title'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert schema['properties']['title']['type'] == ['string', 'null']

    def test_scalar_integer_field(self, tmp_path):
        path = _write([
            ['var:', 'count', 'integer', r'\d+'],
            ['START:'], ['cell:A1', 'count'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert schema['properties']['count']['type'] == ['integer', 'null']

    def test_scalar_currency_field(self, tmp_path):
        path = _write([
            ['var:', 'price', 'currency', '.*'],
            ['START:'], ['cell:A1', 'price'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert schema['properties']['price']['type'] == ['number', 'null']

    def test_scalar_boolean_field(self, tmp_path):
        path = _write([
            ['var:', 'active', 'boolean', ''],
            ['START:'], ['cell:A1', 'active'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert schema['properties']['active']['type'] == ['boolean', 'null']

    def test_scalar_date_field(self, tmp_path):
        path = _write([
            ['var:', 'due', 'date', ''],
            ['START:'], ['cell:A1', 'due'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        prop = schema['properties']['due']
        assert 'string' in prop['type']
        assert prop.get('format') == 'date-time'

    def test_scalar_time_field(self, tmp_path):
        path = _write([
            ['var:', 'start', 'time', ''],
            ['START:'], ['cell:A1', 'start'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        prop = schema['properties']['start']
        assert 'string' in prop['type']

    def test_scalar_duration_field(self, tmp_path):
        path = _write([
            ['var:', 'elapsed', 'duration', ''],
            ['START:'], ['cell:A1', 'elapsed'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert 'string' in schema['properties']['elapsed']['type']


# ─── Nested objects via dot notation ────────────────────────────────────────

class TestNestedObjects:
    def test_dot_notation_creates_nested_object(self, tmp_path):
        path = _write([
            ['var:', 'po.number', 'string', '.*'],
            ['var:', 'po.date', 'date', ''],
            ['START:'], ['cell:A1', 'po.number'], ['cell:A2', 'po.date'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        po = schema['properties']['po']
        assert po['type'] == 'object'
        assert 'string' in po['properties']['number']['type']
        assert 'string' in po['properties']['date']['type']

    def test_multiple_groups(self, tmp_path):
        path = _write([
            ['var:', 'a.x', 'string', '.*'],
            ['var:', 'b.y', 'integer', r'\d+'],
            ['START:'], ['cell:A1', 'a.x'], ['cell:A2', 'b.y'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert 'a' in schema['properties']
        assert 'b' in schema['properties']


# ─── Table arrays ───────────────────────────────────────────────────────────

class TestTableArrays:
    def test_table_produces_array(self, tmp_path):
        path = _write([
            ['lbl:', 'h', 'string', 'Name'],
            ['var:', 'item.name', 'string', '.*'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h'],
            ['', 'DATA:*', 'item.name'],
            ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        item = schema['properties']['item']
        assert item['type'] == 'array'
        assert item['items']['type'] == 'object'

    def test_table_data_fields_in_items(self, tmp_path):
        path = _write([
            ['lbl:', 'h', 'string', 'Name'],
            ['var:', 'row.name', 'string', '.*'],
            ['var:', 'row.qty', 'integer', r'\d+'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h', 'h'],
            ['', 'DATA:*', 'row.name', 'row.qty'],
            ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        data_props = schema['properties']['row']['items']['properties']['data']
        assert data_props['type'] == 'array'
        item_props = data_props['items']['properties']
        assert 'string' in item_props['name']['type']
        assert 'integer' in item_props['qty']['type']

    def test_table_with_footer(self, tmp_path):
        path = _write([
            ['lbl:', 'h', 'string', 'Val'],
            ['var:', 'line.amount', 'currency', '.*'],
            ['var:', 'line.total', 'currency', '.*'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'h'],
            ['', 'DATA:*', 'line.amount'],
            ['', 'FOOTER:1', 'line.total'],
            ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        instance = schema['properties']['line']['items']
        assert 'footer' in instance['properties']
        footer = instance['properties']['footer']
        assert 'number' in footer['properties']['total']['type']

    def test_table_with_header_fields(self, tmp_path):
        path = _write([
            ['lbl:', 'cat_lbl', 'string', 'Category'],
            ['var:', 'item.category', 'string', '.*'],
            ['var:', 'item.name', 'string', '.*'],
            ['START:'],
            ['table:*'],
            ['', 'HEADER:1', 'item.category'],
            ['', 'DATA:*', 'item.name'],
            ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        instance = schema['properties']['item']['items']
        assert 'header' in instance['properties']

    def test_source_metadata_in_table(self, tmp_path):
        path = _write([
            ['var:', 'r.v', 'string', '.*'],
            ['START:'],
            ['table:*'],
            ['', 'DATA:*', 'r.v'],
            ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        instance = schema['properties']['r']['items']
        assert '_source' in instance['properties']


# ─── lbl: fields excluded ──────────────────────────────────────────────────

class TestLblExclusion:
    def test_lbl_fields_not_in_schema(self, tmp_path):
        path = _write([
            ['lbl:', 'anchor', 'string', 'Total'],
            ['var:', 'amount', 'currency', '.*'],
            ['START:'], ['cell:A1', 'anchor'], ['cell:A2', 'amount'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert 'anchor' not in schema['properties']
        assert 'amount' in schema['properties']


# ─── Nullable fields ───────────────────────────────────────────────────────

class TestNullable:
    def test_fields_allow_null(self, tmp_path):
        """Extraction can return null for empty cells — schema must allow it."""
        path = _write([
            ['var:', 'val', 'string', '.*'],
            ['START:'], ['cell:A1', 'val'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        prop = schema['properties']['val']
        assert 'null' in prop['type'] or prop.get('type') == ['string', 'null']


# ─── _meta schema property ────────────────────────────────────────────────

class TestMetaSchema:
    def test_meta_property_present_and_optional(self, tmp_path):
        path = _write([
            ['var:', 'x', 'string', '.*'],
            ['START:'], ['cell:A1', 'x'], ['END:'],
        ], tmp_path)
        schema = generate_schema(path)
        assert '_meta' in schema['properties']
        meta = schema['properties']['_meta']
        assert meta['type'] == 'object'
        assert 'run_id' in meta['properties']
        assert 'stats' in meta['properties']
        assert 'issues' in meta['properties']
        assert '_meta' not in schema.get('required', [])
