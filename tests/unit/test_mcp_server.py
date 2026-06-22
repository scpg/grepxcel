"""Tests for the MCP server module."""

import io
import json
import os
import sys

import pytest

mcp_sdk = pytest.importorskip("mcp", reason="mcp SDK not installed")

from grepxcel.mcp_server import create_server, _json_default

FIXTURES = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


def _fixture(name, filename):
    return os.path.join(FIXTURES, name, filename)


class TestCreateServer:
    """Server creation and tool registration."""

    def test_creates_server(self):
        server = create_server()
        assert server is not None
        assert server.name == "grepxcel"

    def test_has_expected_tools(self):
        server = create_server()
        tool_names = set()
        for tool in server._tool_manager._tools.values():
            tool_names.add(tool.fn.__name__)
        expected = {
            'extract', 'validate_pattern', 'lint', 'schema',
            'docs', 'doctor', 'generate_examples',
        }
        assert expected.issubset(tool_names)


class TestExtractTool:
    """The extract MCP tool."""

    def test_extract_simple_invoice(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture('01_simple_invoice', 'pattern-from-draft.xlsx'),
            data=_fixture('01_simple_invoice', 'data.xlsx'),
        )
        parsed = json.loads(result)
        assert 'inv' in parsed
        assert parsed['inv']['number'] == 'AB123456'

    def test_extract_with_sheet(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture('01_simple_invoice', 'pattern-from-draft.xlsx'),
            data=_fixture('01_simple_invoice', 'data.xlsx'),
            sheet='Sheet1',
        )
        parsed = json.loads(result)
        assert isinstance(parsed, dict)

    def test_extract_all_sheets(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture('12_multi_sheet', 'pattern-from-draft.xlsx'),
            data=_fixture('12_multi_sheet', 'data.xlsx'),
            all_sheets=True,
        )
        parsed = json.loads(result)
        assert isinstance(parsed, dict)
        assert len(parsed) >= 2

    def test_extract_returns_valid_json(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture('02_product_catalog', 'pattern-from-draft.xlsx'),
            data=_fixture('02_product_catalog', 'data.xlsx'),
        )
        parsed = json.loads(result)
        assert isinstance(parsed, dict)

    def test_extract_missing_file_returns_empty(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(pattern='/nonexistent/pattern.xlsx', data='/nonexistent/data.xlsx')
        parsed = json.loads(result)
        assert isinstance(parsed, dict)


class TestValidatePatternTool:
    """The validate_pattern MCP tool."""

    def test_valid_pattern(self):
        server = create_server()
        fn = server._tool_manager._tools['validate_pattern'].fn
        result = fn(
            pattern=_fixture('01_simple_invoice', 'pattern-from-draft.xlsx'),
        )
        assert 'VALID' in result

    def test_valid_pattern_verbose(self):
        server = create_server()
        fn = server._tool_manager._tools['validate_pattern'].fn
        result = fn(
            pattern=_fixture('01_simple_invoice', 'pattern-from-draft.xlsx'),
            verbose=True,
        )
        assert 'fields:' in result or 'VALID' in result

    def test_invalid_file_returns_error(self):
        server = create_server()
        fn = server._tool_manager._tools['validate_pattern'].fn
        result = fn(pattern='/nonexistent/pattern.xlsx')
        assert 'failed' in result.lower() or 'error' in result.lower() or 'not' in result.lower()


class TestLintTool:
    """The lint MCP tool."""

    def test_lint_valid_file(self):
        server = create_server()
        fn = server._tool_manager._tools['lint'].fn
        result = fn(file=_fixture('01_simple_invoice', 'data.xlsx'))
        assert len(result) > 0

    def test_lint_returns_string(self):
        server = create_server()
        fn = server._tool_manager._tools['lint'].fn
        result = fn(file=_fixture('02_product_catalog', 'data.xlsx'))
        assert isinstance(result, str)


class TestSchemaTool:
    """The schema MCP tool."""

    def test_schema_returns_valid_json_schema(self):
        server = create_server()
        fn = server._tool_manager._tools['schema'].fn
        result = fn(
            pattern=_fixture('01_simple_invoice', 'pattern-from-draft.xlsx'),
        )
        parsed = json.loads(result)
        assert '$schema' in parsed or 'type' in parsed
        assert parsed.get('type') == 'object'

    def test_schema_has_properties(self):
        server = create_server()
        fn = server._tool_manager._tools['schema'].fn
        result = fn(
            pattern=_fixture('01_simple_invoice', 'pattern-from-draft.xlsx'),
        )
        parsed = json.loads(result)
        assert 'properties' in parsed


class TestDocsTool:
    """The docs MCP tool."""

    def test_docs_creates_file(self, tmp_path):
        server = create_server()
        fn = server._tool_manager._tools['docs'].fn
        result = fn(output_dir=str(tmp_path))
        assert 'pattern-reference.xlsx' in result
        assert (tmp_path / 'pattern-reference.xlsx').is_file()

    def test_docs_default_temp_dir(self):
        server = create_server()
        fn = server._tool_manager._tools['docs'].fn
        result = fn()
        assert 'pattern-reference.xlsx' in result


class TestDoctorTool:
    """The doctor MCP tool."""

    def test_doctor_returns_report(self):
        server = create_server()
        fn = server._tool_manager._tools['doctor'].fn
        result = fn()
        assert isinstance(result, str)
        assert len(result) > 0

    def test_doctor_extract_area(self):
        server = create_server()
        fn = server._tool_manager._tools['doctor'].fn
        result = fn(area='extract')
        assert isinstance(result, str)


class TestGenerateExamplesTool:
    """The generate_examples MCP tool."""

    def test_creates_examples(self, tmp_path):
        server = create_server()
        fn = server._tool_manager._tools['generate_examples'].fn
        out = str(tmp_path / 'examples')
        result = fn(output_dir=out)
        assert 'examples' in result.lower()
        subdirs = [d for d in (tmp_path / 'examples').iterdir() if d.is_dir()]
        assert len(subdirs) == 4

    def test_extraction_works_on_generated_examples(self, tmp_path):
        server = create_server()
        gen_fn = server._tool_manager._tools['generate_examples'].fn
        extract_fn = server._tool_manager._tools['extract'].fn

        out = str(tmp_path / 'examples')
        gen_fn(output_dir=out)

        for d in sorted((tmp_path / 'examples').iterdir()):
            if not d.is_dir():
                continue
            result = extract_fn(
                pattern=str(d / 'pattern.xlsx'),
                data=str(d / 'data.xlsx'),
            )
            parsed = json.loads(result)
            assert isinstance(parsed, dict)
            assert len(parsed) > 0, f"{d.name} produced empty output"


class TestJsonDefault:
    """The _json_default serializer handles datetime types."""

    def test_date(self):
        import datetime
        assert _json_default(datetime.date(2026, 1, 15)) == '2026-01-15'

    def test_datetime(self):
        import datetime
        dt = datetime.datetime(2026, 1, 15, 10, 30, 0)
        assert '2026-01-15' in _json_default(dt)

    def test_timedelta(self):
        import datetime
        td = datetime.timedelta(hours=8, minutes=30)
        assert '8:30:00' in _json_default(td)

    def test_unsupported_type_raises(self):
        with pytest.raises(TypeError):
            _json_default(set())
