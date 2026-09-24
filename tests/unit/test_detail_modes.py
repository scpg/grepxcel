"""
Unit tests for --detail minimal / normal / extended output modes.

Tests cover:
  - _build_minimal_output: strips _source from table instances and top-level _ keys
  - _build_extended_output: adds _ext block from cell_meta
  - Engine.process() with detail_level parameter (integration smoke test)
  - schema.generate_schema() with detail_level parameter
"""
import os
import pytest

from grepxcel.engine import _build_minimal_output, _build_extended_output, _build_nested_output
from grepxcel.models import FieldDef


# ── helpers ───────────────────────────────────────────────────────────────────

def _fd(name: str, type_: str = 'string', role: str = 'var') -> FieldDef:
    return FieldDef(name=name, type=type_, regex='.*', role=role)


def _defs(**kwargs) -> dict:
    return {name: _fd(name, **kw) for name, kw in kwargs.items()}


_SIMPLE_NESTED = {
    'invoice_number': 'INV-001',
    'amount': 1200.0,
    'items': [
        {
            '_source': {'sheet': 'Sheet1', 'ref': 'A1:C5'},
            'header': {'category': 'Widgets'},
            'data': [{'description': 'A', 'qty': 5}],
        }
    ],
}


# ── _build_minimal_output ────────────────────────────────────────────────────

class TestBuildMinimalOutput:
    def test_strips_source_from_table_instances(self):
        out = _build_minimal_output(_SIMPLE_NESTED)
        assert '_source' not in out['items'][0]

    def test_preserves_data_rows(self):
        out = _build_minimal_output(_SIMPLE_NESTED)
        assert out['items'][0]['data'] == [{'description': 'A', 'qty': 5}]

    def test_preserves_scalar_fields(self):
        out = _build_minimal_output(_SIMPLE_NESTED)
        assert out['invoice_number'] == 'INV-001'
        assert out['amount'] == 1200.0

    def test_strips_top_level_meta(self):
        nested = dict(_SIMPLE_NESTED)
        nested['_meta'] = {'run_id': 'abc123'}
        nested['_images'] = {}
        out = _build_minimal_output(nested)
        assert '_meta' not in out
        assert '_images' not in out

    def test_strips_top_level_source(self):
        nested = {'_source': {'file': 'x.xlsx'}, 'name': 'Alice'}
        out = _build_minimal_output(nested)
        assert '_source' not in out
        assert out['name'] == 'Alice'

    def test_non_dict_list_items_preserved(self):
        nested = {'tags': ['a', 'b', 'c']}
        out = _build_minimal_output(nested)
        assert out['tags'] == ['a', 'b', 'c']

    def test_empty_tables_preserved(self):
        nested = {'items': []}
        out = _build_minimal_output(nested)
        assert out['items'] == []


# ── _build_extended_output ───────────────────────────────────────────────────

class TestBuildExtendedOutput:
    def _make_raw(self, cells=None, cell_meta=None, tables=None):
        return {
            'cells': cells or {},
            'tables': tables or [],
            'cell_meta': cell_meta or {},
        }

    def test_no_meta_returns_normal(self):
        raw = self._make_raw(cells={'inv': 'INV-001'})
        defs = _defs(inv={})
        out = _build_extended_output(raw, defs)
        assert '_ext' not in out

    def test_ext_block_present(self):
        raw = self._make_raw(
            cells={'inv': 'INV-001'},
            cell_meta={'inv': {'type': 'string', 'cell_ref': 'B2',
                                'number_format': 'General', 'raw_excel': 'INV-001'}},
        )
        defs = _defs(inv={})
        out = _build_extended_output(raw, defs)
        assert '_ext' in out

    def test_ext_contains_field_meta(self):
        meta = {'type': 'currency', 'cell_ref': 'B4',
                 'number_format': '#,##0.00', 'raw_excel': 1200.0}
        raw = self._make_raw(cells={'amount': 1200.0}, cell_meta={'amount': meta})
        defs = _defs(amount={'type_': 'currency'})
        out = _build_extended_output(raw, defs)
        assert out['_ext']['amount'] == meta

    def test_data_values_unchanged(self):
        raw = self._make_raw(
            cells={'amount': 1200.0},
            cell_meta={'amount': {'type': 'number', 'cell_ref': 'B4',
                                   'number_format': '#,##0', 'raw_excel': 1200.0}},
        )
        defs = _defs(amount={})
        out = _build_extended_output(raw, defs)
        assert out['amount'] == 1200.0

    def test_lbl_fields_excluded_from_ext(self):
        raw = self._make_raw(
            cells={'Amount:': 'Amount:', 'inv': 'INV-001'},
            cell_meta={
                'Amount:': {'type': 'string', 'cell_ref': 'A4', 'number_format': 'General', 'raw_excel': 'Amount:'},
                'inv': {'type': 'string', 'cell_ref': 'B1', 'number_format': 'General', 'raw_excel': 'INV-001'},
            },
        )
        defs = {'Amount:': _fd('Amount:', role='lbl'), 'inv': _fd('inv')}
        out = _build_extended_output(raw, defs)
        assert 'Amount:' not in out.get('_ext', {})
        assert 'inv' in out.get('_ext', {})

    def test_dotted_fields_nested_in_ext(self):
        raw = self._make_raw(
            cells={'inv.number': 'INV-001'},
            cell_meta={'inv.number': {'type': 'string', 'cell_ref': 'B2',
                                       'number_format': 'General', 'raw_excel': 'INV-001'}},
        )
        defs = _defs(**{'inv.number': {}})
        out = _build_extended_output(raw, defs)
        assert out['_ext']['inv']['number']['cell_ref'] == 'B2'


# ── schema with detail_level ─────────────────────────────────────────────────

class TestSchemaDetailLevel:
    _FIXTURE_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures',
                                 '01_simple_invoice')

    def _pattern(self):
        p = os.path.join(self._FIXTURE_DIR, '01_simple_invoice_pattern-from-draft.xlsx')
        if not os.path.exists(p):
            pytest.skip('fixture 01 pattern-from-draft not found')
        return p

    def test_normal_has_meta(self):
        from grepxcel.schema import generate_schema
        schema = generate_schema(self._pattern(), detail_level='normal')
        assert '_meta' in schema['properties']

    def test_minimal_no_meta(self):
        from grepxcel.schema import generate_schema
        schema = generate_schema(self._pattern(), detail_level='minimal')
        assert '_meta' not in schema['properties']

    def test_minimal_no_source_in_tables(self):
        from grepxcel.schema import generate_schema
        # Use a fixture with tables
        pattern = os.path.join(
            os.path.dirname(__file__), '..', 'fixtures',
            '02_product_catalog', '02_product_catalog_pattern-from-draft.xlsx',
        )
        if not os.path.exists(pattern):
            pytest.skip('fixture 02 not found')
        schema = generate_schema(pattern, detail_level='minimal')
        for key, value in schema['properties'].items():
            if isinstance(value, dict) and value.get('type') == 'array':
                instance_props = value.get('items', {}).get('properties', {})
                assert '_source' not in instance_props

    def test_extended_has_defs(self):
        from grepxcel.schema import generate_schema
        schema = generate_schema(self._pattern(), detail_level='extended')
        assert '$defs' in schema
        assert 'FieldDetail' in schema['$defs']

    def test_extended_has_ext_block(self):
        from grepxcel.schema import generate_schema
        schema = generate_schema(self._pattern(), detail_level='extended')
        assert '_ext' in schema['properties']

    def test_extended_ext_refs_field_detail(self):
        from grepxcel.schema import generate_schema
        schema = generate_schema(self._pattern(), detail_level='extended')
        ext_props = schema['properties']['_ext']['properties']

        def _collect_refs(d):
            refs = []
            if isinstance(d, dict):
                if d.get('$ref') == '#/$defs/FieldDetail':
                    refs.append(d)
                for v in d.values():
                    refs.extend(_collect_refs(v))
            return refs

        refs = _collect_refs(ext_props)
        assert refs, 'no $ref to FieldDetail found in _ext properties (recursive)'


# ── CLI smoke: --detail minimal strips _source ───────────────────────────────

class TestCLIDetailLevel:
    def test_process_minimal_strips_source_from_tables(self, tmp_path):
        """Smoke test: engine.process with detail_level='minimal' strips _source."""
        # Use fixture 02 which has tables with _source blocks
        import grepxcel
        fixture_dir = os.path.join(
            os.path.dirname(__file__), '..', 'fixtures', '02_product_catalog',
        )
        pattern = os.path.join(fixture_dir, '02_product_catalog_pattern-from-draft.xlsx')
        data = os.path.join(fixture_dir, '02_product_catalog_data.xlsx')
        if not os.path.exists(pattern) or not os.path.exists(data):
            pytest.skip('fixture 02 not found')

        from grepxcel.engine import Engine
        from grepxcel.logger import Logger, VerbosityLevel
        engine = Engine()
        result = engine.process(pattern, data,
                                logger=Logger(level=VerbosityLevel.QUIET),
                                detail_level='minimal')
        # Find table arrays and check no _source
        for key, val in result.items():
            if isinstance(val, list):
                for inst in val:
                    assert '_source' not in inst, f'_source found in {key} instance'

    def test_process_extended_has_ext(self):
        """Smoke test: engine.process with detail_level='extended' adds _ext."""
        fixture_dir = os.path.join(
            os.path.dirname(__file__), '..', 'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixture_dir, '01_simple_invoice_pattern-from-draft.xlsx')
        data = os.path.join(fixture_dir, '01_simple_invoice_data.xlsx')
        if not os.path.exists(pattern) or not os.path.exists(data):
            pytest.skip('fixture 01 not found')

        from grepxcel.engine import Engine
        from grepxcel.logger import Logger, VerbosityLevel
        engine = Engine()
        result = engine.process(pattern, data,
                                logger=Logger(level=VerbosityLevel.QUIET),
                                detail_level='extended')
        assert '_ext' in result, '_ext block missing in extended output'

    def test_process_extended_ext_has_cell_ref(self):
        """Extended mode: each field in _ext has a cell_ref."""
        fixture_dir = os.path.join(
            os.path.dirname(__file__), '..', 'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixture_dir, '01_simple_invoice_pattern-from-draft.xlsx')
        data = os.path.join(fixture_dir, '01_simple_invoice_data.xlsx')
        if not os.path.exists(pattern) or not os.path.exists(data):
            pytest.skip('fixture 01 not found')

        from grepxcel.engine import Engine
        from grepxcel.logger import Logger, VerbosityLevel
        engine = Engine()
        result = engine.process(pattern, data,
                                logger=Logger(level=VerbosityLevel.QUIET),
                                detail_level='extended')
        ext = result.get('_ext', {})
        assert ext, '_ext is empty'
        # Check at least one leaf has cell_ref
        def _find_cell_refs(d):
            if isinstance(d, dict):
                if 'cell_ref' in d:
                    return [d['cell_ref']]
                return [r for v in d.values() for r in _find_cell_refs(v)]
            return []
        refs = _find_cell_refs(ext)
        assert refs, 'no cell_ref found in _ext'

    def test_process_normal_unchanged(self):
        """Normal mode must produce identical output to before (no _ext, _source present)."""
        fixture_dir = os.path.join(
            os.path.dirname(__file__), '..', 'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixture_dir, '01_simple_invoice_pattern-from-draft.xlsx')
        data = os.path.join(fixture_dir, '01_simple_invoice_data.xlsx')
        if not os.path.exists(pattern) or not os.path.exists(data):
            pytest.skip('fixture 01 not found')

        from grepxcel.engine import Engine
        from grepxcel.logger import Logger, VerbosityLevel
        engine = Engine()
        result = engine.process(pattern, data,
                                logger=Logger(level=VerbosityLevel.QUIET),
                                detail_level='normal')
        assert '_ext' not in result
