"""CLI unit tests (ported from copilot/analyze-test-coverage with compatibility fixes).

Fixes applied vs. the original Copilot-generated version:
  - _resolve_sheet now returns the string value unchanged (no int coercion);
    assertions updated to compare against the string form.
  - Logger.__init__ accepts keyword-only extras (log_format, source); FakeLogger
    constructors updated to **kwargs so they stay forward-compatible.
  - _process_file returns (ok, issue_fields) tuple; tests now unpack it.
  - drafter fake module must expose all names imported at the top of _run_draft
    (GitHubModelsBackend, NvidiaBackend, OpenAICompatBackend) even for tests
    that only exercise other code paths.
"""
import datetime
import types
from types import SimpleNamespace

import pytest

from grepxcel import cli


def test_json_default_handles_dates_and_rejects_unknown_types():
    d = datetime.date(2026, 1, 2)
    dt = datetime.datetime(2026, 1, 2, 3, 4, 5)

    assert cli._json_default(d) == '2026-01-02'
    assert cli._json_default(dt) == '2026-01-02T03:04:05'

    with pytest.raises(TypeError):
        cli._json_default(object())


def test_resolve_sheet_returns_string_unchanged():
    # _resolve_sheet no longer coerces to int; downstream resolution happens in
    # the engine so that a sheet actually *named* '2025' isn't mistaken for an index.
    assert cli._resolve_sheet(SimpleNamespace(sheet='2')) == '2'
    assert cli._resolve_sheet(SimpleNamespace(sheet='Sheet 1')) == 'Sheet 1'
    assert cli._resolve_sheet(SimpleNamespace()) is None


def test_output_stem_uses_parent_for_duplicate_basenames(tmp_path):
    left = tmp_path / 'a' / 'report.xlsx'
    right = tmp_path / 'b' / 'report.xlsx'
    left.parent.mkdir()
    right.parent.mkdir()
    left.write_text('x', encoding='utf-8')
    right.write_text('x', encoding='utf-8')

    stem_left = cli._output_stem(str(left), [str(left), str(right)])
    stem_right = cli._output_stem(str(right), [str(left), str(right)])

    assert stem_left == 'a_report'
    assert stem_right == 'b_report'


def test_process_file_writes_stdout_and_uses_process(monkeypatch, capsys):
    calls = {}
    logger_instances = []

    class FakeLogger:
        # Logger is called with keyword arguments including log_format and source;
        # use **kwargs so the fake stays compatible with future signature changes.
        def __init__(self, **kwargs):
            self._err = False
            self._warn = False
            self.closed = False
            self.last_stats = None
            logger_instances.append(self)

        def close(self):
            self.closed = True

        def has_errors(self):
            return self._err

        def has_warnings(self):
            return self._warn

        def issues(self):
            return []

        def build_meta(self):
            return {}

    class FakeEngine:
        def process(self, pattern, data_file, **kwargs):
            calls['process'] = (pattern, data_file, kwargs)
            return {'ok': True}

    monkeypatch.setattr(cli, 'Logger', FakeLogger)
    monkeypatch.setattr(cli, 'Engine', FakeEngine)

    args = SimpleNamespace(
        debug=False,
        verbose=0,
        files=['data.xlsx'],
        log=None,
        format='nested',
        all_sheets=False,
        max_size=5,
        max_uncompressed=50,
        max_cell_len=1000,
        max_rows=2048,
        max_columns=1024,
        sheet='3',
        output=None,
    )

    ok, _ = cli._process_file('pattern.xlsx', 'data.xlsx', args, stem='data')

    assert ok is True
    assert '"ok": true' in capsys.readouterr().out.lower()
    # _resolve_sheet now returns string '3', not int 3
    assert calls['process'][2]['sheet'] == '3'
    assert logger_instances[0].closed is True


def test_process_file_with_all_sheets_flag(monkeypatch, tmp_path, capsys):
    calls = {}

    class FakeLogger:
        def __init__(self, **kwargs):
            self.closed = False
            self.last_stats = None

        def close(self):
            self.closed = True

        def has_errors(self):
            return False

        def has_warnings(self):
            return True

        def issues(self):
            return []

        def build_meta(self):
            return {}

    class FakeEngine:
        def process(self, *args, **kwargs):
            raise AssertionError('process() should not be called for --all-sheets')

        def process_all(self, pattern, data_file, **kwargs):
            calls['process_all'] = (pattern, data_file, kwargs)
            return {'sheet1': {'ok': True}}

    monkeypatch.setattr(cli, 'Logger', FakeLogger)
    monkeypatch.setattr(cli, 'Engine', FakeEngine)

    out_dir = tmp_path / 'json'
    args = SimpleNamespace(
        debug=False,
        verbose=0,
        files=['a.xlsx', 'b.xlsx'],
        log=None,
        format='legacy',
        all_sheets=True,
        max_size=5,
        max_uncompressed=50,
        max_cell_len=1000,
        max_rows=2048,
        max_columns=1024,
        sheet=None,
        output=str(out_dir),
    )

    ok, _ = cli._process_file('pattern.xlsx', 'a.xlsx', args, stem='a')

    assert ok is False   # has_warnings() → True so ok = not True = False
    assert (out_dir / 'a.json').exists()
    assert calls['process_all'][2]['output_format'] == 'legacy'
    assert 'JSON written to' in capsys.readouterr().err


def test_process_file_handles_engine_exception(monkeypatch, capsys):
    class FakeLogger:
        def __init__(self, **kwargs):
            self.closed = False
            self.last_stats = None

        def close(self):
            self.closed = True

        def has_errors(self):
            return False

        def has_warnings(self):
            return False

        def issues(self):
            return []

        def build_meta(self):
            return {}

    class FakeEngine:
        def process(self, pattern, data_file, **kwargs):
            raise RuntimeError('boom')

    monkeypatch.setattr(cli, 'Logger', FakeLogger)
    monkeypatch.setattr(cli, 'Engine', FakeEngine)

    args = SimpleNamespace(
        debug=False,
        verbose=0,
        files=['data.xlsx'],
        log=None,
        format='nested',
        all_sheets=False,
        max_size=5,
        max_uncompressed=50,
        max_cell_len=1000,
        max_rows=2048,
        max_columns=1024,
        sheet=None,
        output=None,
    )

    ok, _ = cli._process_file('pattern.xlsx', 'data.xlsx', args, stem='data')

    assert ok is False
    assert 'Unexpected error processing data.xlsx: boom' in capsys.readouterr().err


def test_run_docs_imports_generator_and_writes(monkeypatch, capsys):
    called = {}

    class FakeDocsGenerator:
        def write(self, path):
            called['path'] = path

    fake_mod = types.ModuleType('grepxcel.docs_generator')
    fake_mod.DocsGenerator = FakeDocsGenerator
    monkeypatch.setitem(__import__('sys').modules, 'grepxcel.docs_generator', fake_mod)

    code = cli._run_docs(SimpleNamespace(output='ref.xlsx'))

    assert code == 0
    assert called['path'] == 'ref.xlsx'
    assert 'Pattern reference written to: ref.xlsx' in capsys.readouterr().err


def _make_fake_drafter_mod(extra=None):
    """Return a fake grepxcel.drafter module with all names _run_draft imports."""
    fake_mod = types.ModuleType('grepxcel.drafter')
    # These four names are always imported at the top of _run_draft; the fake module
    # must expose all of them even when only one backend is exercised by the test.
    fake_mod.ClaudeBackend = object
    fake_mod.GitHubModelsBackend = object
    fake_mod.NvidiaBackend = object
    fake_mod.OpenAICompatBackend = object
    fake_mod.PatternDrafter = object
    if extra:
        for k, v in extra.items():
            setattr(fake_mod, k, v)
    return fake_mod


def test_run_draft_gemini_is_disabled(monkeypatch, capsys):
    monkeypatch.setitem(
        __import__('sys').modules, 'grepxcel.drafter', _make_fake_drafter_mod()
    )

    args = SimpleNamespace(
        backend='gemini',
        file='in.xlsx',
        output='out.xlsx',
        sheet=None,
        max_size=5,
        max_uncompressed=50,
        verbose=False,
        dry_run=False,
    )

    code = cli._run_draft(args)

    assert code == 1
    assert 'planned for a future release' in capsys.readouterr().err


def test_run_draft_claude_uses_backend_and_returns_run_code(monkeypatch, capsys):
    called = {}

    class FakeClaudeBackend:
        pass

    class FakePatternDrafter:
        def __init__(self, **kwargs):
            called['kwargs'] = kwargs

        def run(self):
            return 7

    monkeypatch.setitem(
        __import__('sys').modules,
        'grepxcel.drafter',
        _make_fake_drafter_mod({
            'ClaudeBackend': FakeClaudeBackend,
            'PatternDrafter': FakePatternDrafter,
        }),
    )

    args = SimpleNamespace(
        backend='claude',
        file='in.xlsx',
        output='out.xlsx',
        sheet='4',
        max_size=5,
        max_uncompressed=50,
        verbose=True,
        dry_run=True,
    )

    code = cli._run_draft(args)

    assert code == 7
    assert isinstance(called['kwargs']['backend'], FakeClaudeBackend)
    # _resolve_sheet returns the string unchanged; downstream engine converts to int
    assert called['kwargs']['sheet'] == '4'
    assert "Anthropic's API" in capsys.readouterr().err


def test_main_aliases_suggest_to_draft(monkeypatch):
    captured = {}

    class FakeParser:
        def parse_args(self, argv):
            captured['argv'] = argv
            return SimpleNamespace(command='draft')

    monkeypatch.setattr(cli, '_build_parser', lambda: FakeParser())
    monkeypatch.setattr(cli, '_run_draft', lambda args: 0)

    with pytest.raises(SystemExit) as exc:
        cli.main(['suggest', 'report.xlsx'])

    assert captured['argv'] == ['draft', 'report.xlsx']
    assert exc.value.code == 0


def test_main_extract_returns_nonzero_if_any_file_fails(monkeypatch):
    class FakeParser:
        def parse_args(self, argv):
            return SimpleNamespace(
                command='extract', pattern='p.xlsx', files=['a.xlsx', 'b.xlsx']
            )

    # _process_file now returns (ok, issue_fields); mock must match that signature.
    results = iter([(True, []), (False, [])])
    monkeypatch.setattr(cli, '_build_parser', lambda: FakeParser())
    monkeypatch.setattr(cli, '_output_stem', lambda data_file, all_files: data_file)
    monkeypatch.setattr(
        cli, '_process_file',
        lambda pattern, data_file, args, stem: next(results),
    )
    # Stub out expansion and pattern-check (imported locally in main) to keep
    # this test focused on the exit-code logic.
    monkeypatch.setattr(cli, '_expand_files', lambda paths, **kw: paths)
    import grepxcel.pattern_check as _pc
    from grepxcel.pattern_check import CheckResult
    monkeypatch.setattr(_pc, 'check_pattern', lambda path: CheckResult(path=path, valid=True))

    with pytest.raises(SystemExit) as exc:
        cli.main(['extract', '-p', 'p.xlsx', 'a.xlsx', 'b.xlsx'])

    assert exc.value.code == 1


# ── Grouped help (-h) ─────────────────────────────────────────────────────────

class TestGroupedHelp:
    def _help_text(self, capsys):
        with pytest.raises(SystemExit) as exc:
            cli.main(['-h'])
        assert exc.value.code == 0
        return capsys.readouterr().out

    def test_usage_line_contains_command(self, capsys):
        out = self._help_text(capsys)
        assert 'COMMAND' in out

    def test_all_six_sections_present(self, capsys):
        out = self._help_text(capsys)
        for section in ('extract & validate', 'onboarding', 'automation',
                        'AI & MCP', 'inspection', 'compliance & ops'):
            assert section in out, f'section missing: {section!r}'

    def test_all_commands_listed(self, capsys):
        out = self._help_text(capsys)
        for cmd in ('extract', 'validate-pattern', 'quickstart', 'web-wizard',
                    'generate-examples', 'docs', 'watch', 'test', 'draft',
                    'generate-skill', 'mcp', 'mcp-config', 'lint', 'schema',
                    'sbom', 'doctor'):
            assert cmd in out, f'command missing from -h output: {cmd!r}'

    def test_wizard_not_listed(self, capsys):
        out = self._help_text(capsys)
        # 'wizard' may appear inside 'web-wizard' — check no standalone entry
        wizard_lines = [l for l in out.splitlines()
                        if 'wizard' in l and 'web-wizard' not in l]
        assert not wizard_lines, f'unexpected wizard line(s): {wizard_lines}'

    def test_no_duplicate_command_listing(self, capsys):
        import re
        out = self._help_text(capsys)
        # command entries have multi-space padding after the name; the section
        # header 'extract & validate:' has only a single space → exclude it
        entries = re.findall(r'^\s+extract\s{2,}', out, re.MULTILINE)
        assert len(entries) == 1, f'expected 1 extract entry, found {len(entries)}'

    def test_per_command_help_hint_present(self, capsys):
        out = self._help_text(capsys)
        assert "grepxcel <command> --help" in out

    def test_synopsis_hint_present(self, capsys):
        out = self._help_text(capsys)
        assert 'grepxcel -h -h' in out


# ── Multi-level help (-h -h and -h -h -h) ────────────────────────────────────

class TestMultiLevelHelp:
    def test_synopsis_exits_0(self, capsys):
        with pytest.raises(SystemExit) as exc:
            cli.main(['-h', '-h'])
        assert exc.value.code == 0

    def test_synopsis_contains_usage_for_each_command(self, capsys):
        with pytest.raises(SystemExit):
            cli.main(['-h', '-h'])
        out = capsys.readouterr().out
        for cmd in ('extract', 'validate-pattern', 'watch', 'draft',
                    'web-wizard', 'lint', 'schema', 'sbom', 'doctor',
                    'quickstart', 'test'):
            assert cmd in out, f'{cmd} missing from synopsis'

    def test_synopsis_shows_required_flags(self, capsys):
        with pytest.raises(SystemExit):
            cli.main(['-h', '-h'])
        out = capsys.readouterr().out
        assert '-p FILE' in out  # extract / watch / test require --pattern

    def test_full_help_exits_0(self, capsys):
        with pytest.raises(SystemExit) as exc:
            cli.main(['-h', '-h', '-h'])
        assert exc.value.code == 0

    def test_full_help_contains_epilogs(self, capsys):
        with pytest.raises(SystemExit):
            cli.main(['-h', '-h', '-h'])
        out = capsys.readouterr().out
        assert 'examples:' in out.lower()

    def test_full_help_covers_all_commands(self, capsys):
        with pytest.raises(SystemExit):
            cli.main(['-h', '-h', '-h'])
        out = capsys.readouterr().out
        for cmd in ('extract', 'validate-pattern', 'watch', 'draft',
                    'web-wizard', 'lint', 'schema', 'sbom', 'doctor',
                    'quickstart', 'test'):
            assert f'grepxcel {cmd}' in out, f'{cmd} missing from full help'

    def test_help_help_also_accepts_long_flags(self, capsys):
        with pytest.raises(SystemExit) as exc:
            cli.main(['--help', '--help'])
        assert exc.value.code == 0

    def test_three_long_flags_triggers_full_help(self, capsys):
        with pytest.raises(SystemExit) as exc:
            cli.main(['--help', '--help', '--help'])
        out = capsys.readouterr().out
        assert exc.value.code == 0
        assert 'examples:' in out.lower()


# ── Wizard command removed ────────────────────────────────────────────────────

class TestWizardRemoved:
    def test_wizard_command_is_not_registered(self):
        import argparse
        p = cli._build_parser()
        sub_action = next(
            a for a in p._actions if isinstance(a, argparse._SubParsersAction)
        )
        assert 'wizard' not in sub_action.choices

    def test_wizard_argv_exits_nonzero(self, capsys):
        with pytest.raises(SystemExit) as exc:
            cli.main(['wizard', 'data.xlsx'])
        assert exc.value.code != 0
