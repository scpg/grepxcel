"""Tests for `grepxcel generate-skill` (grepxcel.skill)."""
import pytest

from grepxcel.skill import generate_skill, run_skill


# Subcommands that must appear in the introspected command reference.
_EXPECTED_COMMANDS = ('extract', 'validate-pattern', 'draft', 'schema', 'lint',
                      'docs', 'doctor', 'generate-skill')


class TestGenerateSkill:
    def test_claude_has_frontmatter(self):
        text = generate_skill('claude')
        assert text.startswith('---\n')
        assert 'name: grepxcel' in text
        assert 'description:' in text

    def test_agents_md_has_no_frontmatter(self):
        text = generate_skill('agents-md')
        assert not text.startswith('---')
        assert text.lstrip().startswith('# grepxcel')

    def test_command_reference_is_introspected(self):
        text = generate_skill('claude')
        for cmd in _EXPECTED_COMMANDS:
            assert f'grepxcel {cmd}' in text, f'{cmd} missing from skill doc'

    def test_includes_when_to_use_and_gotchas(self):
        text = generate_skill('agents-md')
        assert 'When to use' in text
        assert 'stdout' in text          # output convention gotcha

    def test_cursor_has_frontmatter(self):
        text = generate_skill('cursor')
        assert text.startswith('---\n')
        assert 'globs:' in text
        assert 'alwaysApply: false' in text
        assert '**/*.xlsx' in text

    def test_unknown_target_raises(self):
        with pytest.raises(ValueError, match='target'):
            generate_skill('unknown_engine_xyz')


class TestRunSkill:
    def test_writes_file(self, tmp_path):
        out = tmp_path / 'SKILL.md'
        rc = run_skill('claude', str(out))
        assert rc == 0
        assert out.read_text(encoding='utf-8').startswith('---\n')

    def test_stdout_default(self, tmp_path, capsys):
        rc = run_skill('agents-md')
        assert rc == 0
        assert '# grepxcel' in capsys.readouterr().out

    def test_unknown_target_returns_1(self, capsys):
        assert run_skill('nope') == 1
