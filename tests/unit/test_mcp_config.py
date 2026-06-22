"""Tests for the MCP config emitter."""

import io
import json
import sys

import pytest

from grepxcel.mcp_config import generate_config, run_mcp_config, _find_grepxcel_command


class TestFindCommand:
    """Detect the grepxcel command path."""

    def test_finds_command(self):
        cmd = _find_grepxcel_command()
        assert 'grepxcel' in cmd

    def test_returns_absolute_path_when_in_venv(self):
        cmd = _find_grepxcel_command()
        if '.venv' in sys.executable:
            assert cmd.startswith('/')


class TestGenerateConfig:
    """Config generation for different targets."""

    @pytest.mark.parametrize('target', ['claude-code', 'claude-desktop', 'cursor'])
    def test_contains_valid_json(self, target):
        output = generate_config(target)
        lines = [l for l in output.splitlines() if not l.startswith('#')]
        body = '\n'.join(lines).strip()
        parsed = json.loads(body)
        assert 'mcpServers' in parsed
        assert 'grepxcel' in parsed['mcpServers']

    @pytest.mark.parametrize('target', ['claude-code', 'claude-desktop', 'cursor'])
    def test_server_has_command_and_args(self, target):
        output = generate_config(target)
        lines = [l for l in output.splitlines() if not l.startswith('#')]
        body = '\n'.join(lines).strip()
        parsed = json.loads(body)
        server = parsed['mcpServers']['grepxcel']
        assert 'command' in server
        assert server['args'] == ['mcp']

    def test_claude_code_header(self):
        output = generate_config('claude-code')
        assert 'settings.json' in output

    def test_claude_desktop_header(self):
        output = generate_config('claude-desktop')
        assert 'claude_desktop_config.json' in output

    def test_cursor_header(self):
        output = generate_config('cursor')
        assert '.cursor/mcp.json' in output

    def test_unknown_target_raises(self):
        with pytest.raises(ValueError, match='Unknown target'):
            generate_config('unknown-agent')

    def test_command_path_is_grepxcel(self):
        output = generate_config('claude-code')
        lines = [l for l in output.splitlines() if not l.startswith('#')]
        parsed = json.loads('\n'.join(lines).strip())
        cmd = parsed['mcpServers']['grepxcel']['command']
        assert 'grepxcel' in cmd


class TestRunMcpConfig:
    """The run_mcp_config entry point."""

    def test_writes_to_stdout_by_default(self, capsys):
        rc = run_mcp_config('claude-code')
        assert rc == 0
        captured = capsys.readouterr()
        assert 'mcpServers' in captured.out

    def test_writes_to_custom_stream(self):
        buf = io.StringIO()
        rc = run_mcp_config('claude-code', out=buf)
        assert rc == 0
        assert 'mcpServers' in buf.getvalue()

    @pytest.mark.parametrize('target', ['claude-code', 'claude-desktop', 'cursor'])
    def test_all_targets_succeed(self, target):
        buf = io.StringIO()
        rc = run_mcp_config(target, out=buf)
        assert rc == 0
        assert len(buf.getvalue()) > 0
