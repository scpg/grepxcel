"""Tests for the MCP server module."""

import io
import json
import os
import sys

import pytest

mcp_sdk = pytest.importorskip("mcp", reason="mcp SDK not installed")

from grepxcel.mcp_server import (
    PathSandboxError, _safe_path, create_server, _json_default,
)

FIXTURES = os.path.join(os.path.dirname(__file__), '..', 'fixtures')

from tests.conftest import find_data_file, find_pattern_xlsx


def _fixture(name, filename):
    return os.path.join(FIXTURES, name, filename)


def _fixture_data(name):
    return find_data_file(os.path.join(FIXTURES, name))


def _fixture_pattern(name, variant='pattern-from-draft.xlsx'):
    return _fixture(name, f'{name}_{variant}')


class TestPathSandbox:
    """Path sandboxing rejects escapes and allows safe paths."""

    def test_relative_path_within_sandbox(self, tmp_path):
        (tmp_path / 'data.xlsx').touch()
        result = _safe_path('data.xlsx', str(tmp_path))
        assert result == str(tmp_path / 'data.xlsx')

    def test_nested_relative_path(self, tmp_path):
        sub = tmp_path / 'sub'
        sub.mkdir()
        (sub / 'file.xlsx').touch()
        result = _safe_path('sub/file.xlsx', str(tmp_path))
        assert result == str(sub / 'file.xlsx')

    def test_dotdot_traversal_rejected(self, tmp_path):
        with pytest.raises(PathSandboxError, match='outside the sandbox'):
            _safe_path('../../../etc/passwd', str(tmp_path))

    def test_absolute_path_outside_rejected(self, tmp_path):
        with pytest.raises(PathSandboxError, match='outside the sandbox'):
            _safe_path('/etc/passwd', str(tmp_path))

    def test_absolute_path_inside_allowed(self, tmp_path):
        target = tmp_path / 'ok.xlsx'
        target.touch()
        result = _safe_path(str(target), str(tmp_path))
        assert result == str(target)

    def test_symlink_escape_rejected(self, tmp_path):
        link = tmp_path / 'sneaky'
        link.symlink_to('/tmp')
        with pytest.raises(PathSandboxError, match='outside the sandbox'):
            _safe_path('sneaky/secret.txt', str(tmp_path))

    def test_double_dotdot_after_subdir(self, tmp_path):
        (tmp_path / 'a').mkdir()
        with pytest.raises(PathSandboxError, match='outside the sandbox'):
            _safe_path('a/../../escape', str(tmp_path))

    def test_sandbox_root_itself_allowed(self, tmp_path):
        result = _safe_path('.', str(tmp_path))
        assert result == os.path.realpath(str(tmp_path))


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

    def test_custom_sandbox_root(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        assert server is not None


class TestExtractTool:
    """The extract MCP tool."""

    def test_extract_simple_invoice(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture_pattern('01_simple_invoice'),
            data=_fixture_data('01_simple_invoice'),
        )
        parsed = json.loads(result)
        assert 'inv' in parsed
        assert parsed['inv']['number'] == 'AB123456'

    def test_extract_with_sheet(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture_pattern('01_simple_invoice'),
            data=_fixture_data('01_simple_invoice'),
            sheet='Sheet1',
        )
        parsed = json.loads(result)
        assert isinstance(parsed, dict)

    def test_extract_all_sheets(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture_pattern('12_multi_sheet'),
            data=_fixture_data('12_multi_sheet'),
            all_sheets=True,
        )
        parsed = json.loads(result)
        assert isinstance(parsed, dict)
        assert len(parsed) >= 2

    def test_extract_returns_valid_json(self):
        server = create_server()
        fn = server._tool_manager._tools['extract'].fn
        result = fn(
            pattern=_fixture_pattern('02_product_catalog'),
            data=_fixture_data('02_product_catalog'),
        )
        parsed = json.loads(result)
        assert isinstance(parsed, dict)

    def test_extract_path_traversal_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['extract'].fn
        with pytest.raises(PathSandboxError):
            fn(pattern='../../../etc/passwd', data='data.xlsx')

    def test_extract_absolute_escape_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['extract'].fn
        with pytest.raises(PathSandboxError):
            fn(pattern='/etc/passwd', data='/etc/shadow')


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

    def test_invalid_file_returns_error(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['validate_pattern'].fn
        result = fn(pattern='nonexistent.xlsx')
        assert 'failed' in result.lower() or 'error' in result.lower() or 'not' in result.lower()

    def test_validate_path_traversal_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['validate_pattern'].fn
        with pytest.raises(PathSandboxError):
            fn(pattern='../../secret.csv')


class TestLintTool:
    """The lint MCP tool."""

    def test_lint_valid_file(self):
        server = create_server()
        fn = server._tool_manager._tools['lint'].fn
        result = fn(file=_fixture_data('01_simple_invoice'))
        assert len(result) > 0

    def test_lint_returns_string(self):
        server = create_server()
        fn = server._tool_manager._tools['lint'].fn
        result = fn(file=_fixture_data('02_product_catalog'))
        assert isinstance(result, str)

    def test_lint_path_traversal_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['lint'].fn
        with pytest.raises(PathSandboxError):
            fn(file='/etc/passwd')


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

    def test_schema_path_traversal_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['schema'].fn
        with pytest.raises(PathSandboxError):
            fn(pattern='../../../etc/passwd')


class TestDocsTool:
    """The docs MCP tool."""

    def test_docs_creates_file(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['docs'].fn
        out = tmp_path / 'docs-out'
        out.mkdir()
        result = fn(output_dir=str(out))
        assert 'pattern-reference.xlsx' in result
        assert (out / 'pattern-reference.xlsx').is_file()

    def test_docs_default_temp_dir(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['docs'].fn
        result = fn()
        assert 'pattern-reference.xlsx' in result

    def test_docs_path_traversal_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['docs'].fn
        with pytest.raises(PathSandboxError):
            fn(output_dir='/tmp/evil')


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
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['generate_examples'].fn
        out = str(tmp_path / 'examples')
        result = fn(output_dir=out)
        assert 'examples' in result.lower()
        subdirs = [d for d in (tmp_path / 'examples').iterdir() if d.is_dir()]
        assert len(subdirs) == 4

    def test_generate_examples_path_traversal_rejected(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
        fn = server._tool_manager._tools['generate_examples'].fn
        with pytest.raises(PathSandboxError):
            fn(output_dir='/tmp/evil-examples')

    def test_extraction_works_on_generated_examples(self, tmp_path):
        server = create_server(sandbox_root=str(tmp_path))
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
