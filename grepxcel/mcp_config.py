"""Emit the MCP server configuration for various AI agents."""

import json
import os
import sys


def _find_grepxcel_command() -> str:
    """Detect the installed grepxcel command path."""
    grepxcel_bin = os.path.join(os.path.dirname(sys.executable), 'grepxcel')
    if os.path.isfile(grepxcel_bin):
        return grepxcel_bin
    import shutil
    system_bin = shutil.which('grepxcel')
    if system_bin:
        return system_bin
    return 'grepxcel'


def generate_config(target: str = 'claude-code') -> str:
    """Return the MCP config JSON for the given target.

    Targets:
        claude-code     — .claude/settings.json format
        claude-desktop  — claude_desktop_config.json format
        cursor          — .cursor/mcp.json format
    """
    cmd = _find_grepxcel_command()

    server_def = {
        "command": cmd,
        "args": ["mcp"],
    }

    if target == 'claude-code':
        config = {
            "mcpServers": {
                "grepxcel": server_def,
            }
        }
        header = (
            "# Preferred: run once in your terminal —\n"
            "#   claude mcp add grepxcel -- grepxcel mcp\n"
            "#\n"
            "# Or add manually to ~/.claude.json (global) or\n"
            "# .claude/settings.json (project-scoped):\n"
        )
    elif target == 'claude-desktop':
        config = {
            "mcpServers": {
                "grepxcel": server_def,
            }
        }
        header = (
            "# Add to ~/Library/Application Support/Claude/"
            "claude_desktop_config.json\n"
            "# (macOS) or %APPDATA%\\Claude\\claude_desktop_config.json "
            "(Windows):\n"
        )
    elif target == 'cursor':
        config = {
            "mcpServers": {
                "grepxcel": server_def,
            }
        }
        header = "# Add to .cursor/mcp.json in your project root:\n"
    else:
        raise ValueError(f"Unknown target: {target!r}")

    body = json.dumps(config, indent=2)
    return f"{header}\n{body}\n"


def run_mcp_config(target: str = 'claude-code', out=None) -> int:
    """Print the MCP config and return 0."""
    if out is None:
        out = sys.stdout
    out.write(generate_config(target))
    return 0
