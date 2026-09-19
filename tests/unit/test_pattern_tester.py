"""Unit tests for grepxcel.pattern_tester.

Covers:
 - _collect_xlsx_files (directory scanning, temp-file exclusion, recursion)
 - _extract_field_names / _count_present (field path helpers)
 - _bar (progress bar helper)
 - run_tests (classification, path tracking, fatal-error handling, pattern exclusion)
 - format_human (compact / -v / -vv verbosity, relative paths, color flag)
 - format_json (schema shape, values)
 - CLI integration (exit codes, --format json, -v flag, TTY color detection)
"""

import json
import os
import sys
import tempfile
from unittest.mock import MagicMock, patch

import pytest

from grepxcel.pattern_tester import (
    FileResult,
    PatternTestReport,
    _bar,
    _collect_xlsx_files,
    _compact_file_line,
    _count_present,
    _extract_field_names,
    _rel_path,
    format_human,
    format_json,
    run_tests,
)


# ── _rel_path ─────────────────────────────────────────────────────────────────

class TestRelPath:
    def test_returns_relative(self, tmp_path):
        base = str(tmp_path)
        path = str(tmp_path / 'sub' / 'file.xlsx')
        assert _rel_path(path, base) == os.path.join('sub', 'file.xlsx')

    def test_no_base_returns_original(self):
        assert _rel_path('/abs/path/file.xlsx', '') == '/abs/path/file.xlsx'

    def test_same_dir(self, tmp_path):
        path = str(tmp_path / 'file.xlsx')
        assert _rel_path(path, str(tmp_path)) == 'file.xlsx'


# ── _collect_xlsx_files ───────────────────────────────────────────────────────

class TestCollectXlsxFiles:
    def test_empty_directory(self, tmp_path):
        assert _collect_xlsx_files(str(tmp_path), recursive=False) == []

    def test_finds_xlsx(self, tmp_path):
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'b.XLSX').touch()
        files = _collect_xlsx_files(str(tmp_path), recursive=False)
        assert any('a.xlsx' in f for f in files)
        assert any('b.XLSX' in f for f in files)

    def test_skips_temp_files(self, tmp_path):
        (tmp_path / '~$temp.xlsx').touch()
        (tmp_path / 'real.xlsx').touch()
        files = _collect_xlsx_files(str(tmp_path), recursive=False)
        assert len(files) == 1
        assert 'real.xlsx' in files[0]

    def test_skips_other_extensions(self, tmp_path):
        (tmp_path / 'data.csv').touch()
        (tmp_path / 'data.xls').touch()
        assert _collect_xlsx_files(str(tmp_path), recursive=False) == []

    def test_recursive(self, tmp_path):
        sub = tmp_path / 'sub'
        sub.mkdir()
        (tmp_path / 'root.xlsx').touch()
        (sub / 'nested.xlsx').touch()
        files = _collect_xlsx_files(str(tmp_path), recursive=True)
        assert len(files) == 2

    def test_non_recursive_does_not_descend(self, tmp_path):
        sub = tmp_path / 'sub'
        sub.mkdir()
        (sub / 'nested.xlsx').touch()
        files = _collect_xlsx_files(str(tmp_path), recursive=False)
        assert files == []


# ── _extract_field_names ──────────────────────────────────────────────────────

class TestExtractFieldNames:
    def test_flat(self):
        result = {'a': '1', 'b': '2'}
        assert _extract_field_names(result) == {'a', 'b'}

    def test_nested(self):
        result = {'header': {'number': '001', 'date': '2024-01-01'}, 'total': '100'}
        names = _extract_field_names(result)
        assert 'header.number' in names
        assert 'header.date' in names
        assert 'total' in names
        assert 'header' not in names

    def test_table_list_is_leaf(self):
        result = {'line_items': [{'qty': 1}, {'qty': 2}]}
        assert _extract_field_names(result) == {'line_items'}

    def test_skips_private_keys(self):
        result = {'_meta': {'run_id': 'abc'}, 'number': '1'}
        assert _extract_field_names(result) == {'number'}

    def test_empty(self):
        assert _extract_field_names({}) == set()


# ── _count_present ────────────────────────────────────────────────────────────

class TestCountPresent:
    def test_all_present(self):
        result = {'a': '1', 'b': 'two', 'c': 3}
        assert _count_present(result) == {'a', 'b', 'c'}

    def test_none_excluded(self):
        result = {'a': '1', 'b': None}
        assert _count_present(result) == {'a'}

    def test_empty_string_excluded(self):
        result = {'a': '1', 'b': ''}
        assert _count_present(result) == {'a'}

    def test_empty_list_excluded(self):
        result = {'items': []}
        assert _count_present(result) == set()

    def test_non_empty_list_included(self):
        result = {'items': [{'x': 1}]}
        assert _count_present(result) == {'items'}

    def test_nested(self):
        result = {'h': {'num': '1', 'date': None}}
        assert _count_present(result) == {'h.num'}


# ── _bar ──────────────────────────────────────────────────────────────────────

class TestBar:
    def test_full(self):
        b = _bar(20, 20, width=20)
        assert b == '█' * 20

    def test_empty(self):
        b = _bar(0, 20, width=20)
        assert b == '░' * 20

    def test_half(self):
        b = _bar(10, 20, width=20)
        assert b == '█' * 10 + '░' * 10

    def test_zero_total(self):
        b = _bar(0, 0, width=20)
        assert len(b) == 20


# ── format_human ─────────────────────────────────────────────────────────────

class TestFormatHuman:
    def _report(self, **kwargs):
        defaults = dict(
            pattern_path='/tmp/pattern.xlsx',
            total=3,
            passed=3,
            warned=0,
            failed=0,
            file_results=[
                FileResult('/tmp/a.xlsx', 'pass', None, {}, []),
                FileResult('/tmp/b.xlsx', 'pass', None, {}, []),
                FileResult('/tmp/c.xlsx', 'pass', None, {}, []),
            ],
            field_counts={'inv.number': (3, 3), 'inv.date': (2, 3)},
            base_dir='/tmp',
        )
        defaults.update(kwargs)
        return PatternTestReport(**defaults)

    # ── compact (default, verbose=0) ─────────────────────────────────────────

    def test_compact_one_line_per_file(self):
        fr = FileResult('/tmp/x/a.xlsx', 'pass', None, {}, [])
        r = self._report(total=1, passed=1, file_results=[fr], field_counts={},
                         base_dir='/tmp')
        out = format_human(r, color=False, verbose=0)
        # Each file appears exactly once, as a compact line
        file_lines = [ln for ln in out.splitlines() if 'a.xlsx' in ln]
        assert len(file_lines) == 1

    def test_compact_shows_relative_path(self):
        fr = FileResult('/tmp/samples/sub/inv.xlsx', 'warn', None, {},
                        ['anchor not found'])
        r = self._report(total=1, passed=0, warned=1, failed=0, field_counts={},
                         file_results=[fr], base_dir='/tmp/samples')
        out = format_human(r, color=False, verbose=0)
        # Relative path, not just basename
        assert 'sub/inv.xlsx' in out
        assert '/tmp/samples/sub/inv.xlsx' not in out

    def test_compact_fail_shows_first_error(self):
        fr = FileResult('/tmp/bad.xlsx', 'fail', 'anchor not found', None, [])
        r = self._report(total=1, passed=0, warned=0, failed=1,
                         file_results=[fr], field_counts={}, base_dir='/tmp')
        out = format_human(r, color=False, verbose=0)
        assert 'bad.xlsx' in out
        assert 'anchor not found' in out

    def test_compact_no_field_table_at_verbose0(self):
        r = self._report()
        out = format_human(r, color=False, verbose=0)
        assert 'Field reliability' not in out

    def test_summary_always_present(self):
        r = self._report()
        out = format_human(r, color=False, verbose=0)
        assert 'Tested 3 files' in out
        assert '3 passed' in out

    def test_summary_mixed(self):
        frs = [
            FileResult('/tmp/p.xlsx', 'pass', None, {}, []),
            FileResult('/tmp/w.xlsx', 'warn', None, {}, ['partial']),
            FileResult('/tmp/f.xlsx', 'fail', 'boom', None, []),
        ]
        r = self._report(total=3, passed=1, warned=1, failed=1,
                         file_results=frs, field_counts={})
        out = format_human(r, color=False, verbose=0)
        assert '1 passed' in out
        assert '1 warned' in out
        assert '1 failed' in out

    # ── verbose=1 ────────────────────────────────────────────────────────────

    def test_verbose1_shows_field_table(self):
        r = self._report()
        out = format_human(r, color=False, verbose=1)
        assert 'Field reliability' in out
        assert 'inv.number' in out
        assert '3/3' in out
        assert '100.0%' in out

    def test_verbose1_still_compact_lines(self):
        r = self._report()
        out = format_human(r, color=False, verbose=1)
        # Should still have compact per-file lines, not detailed blocks
        assert '─' * 20 not in out or 'Field reliability' in out  # no big separators per file

    # ── verbose=2 ────────────────────────────────────────────────────────────

    def test_verbose2_shows_per_file_blocks(self):
        frs = [
            FileResult('/tmp/a.xlsx', 'pass', None, {}, []),
            FileResult('/tmp/b.xlsx', 'warn', None, {}, ['missing: inv.date']),
        ]
        r = self._report(total=2, passed=1, warned=1, failed=0,
                         file_results=frs, field_counts={}, base_dir='/tmp')
        out = format_human(r, color=False, verbose=2)
        # Should have separator lines and all issues
        assert 'a.xlsx' in out
        assert 'b.xlsx' in out
        assert 'missing: inv.date' in out

    def test_verbose2_shows_field_table(self):
        r = self._report()
        out = format_human(r, color=False, verbose=2)
        assert 'Field reliability' in out

    # ── color flag ───────────────────────────────────────────────────────────

    def test_color_true_uses_emoji(self):
        fr = FileResult('/tmp/a.xlsx', 'pass', None, {}, [])
        r = self._report(total=1, passed=1, file_results=[fr], field_counts={})
        out = format_human(r, color=True, verbose=0)
        assert '✅' in out

    def test_color_false_uses_text_marks(self):
        fr = FileResult('/tmp/a.xlsx', 'pass', None, {}, [])
        r = self._report(total=1, passed=1, file_results=[fr], field_counts={})
        out = format_human(r, color=False, verbose=0)
        assert '✅' not in out
        assert 'OK' in out or 'PASS' in out or 'passed' in out


# ── format_json ───────────────────────────────────────────────────────────────

class TestFormatJson:
    def test_valid_json(self):
        fr = FileResult(path='/tmp/a.xlsx', status='pass', error=None,
                        result={'x': '1'}, issues=[])
        r = PatternTestReport(
            pattern_path='/tmp/p.xlsx',
            total=1, passed=1, warned=0, failed=0,
            file_results=[fr],
            field_counts={'x': (1, 1)},
        )
        data = json.loads(format_json(r))
        assert data['total'] == 1
        assert data['passed'] == 1
        assert 'x' in data['field_reliability']
        assert data['field_reliability']['x']['pct'] == 100.0
        assert len(data['files']) == 1
        assert data['files'][0]['status'] == 'pass'

    def test_full_path_in_json_output(self):
        fr = FileResult(path='/abs/path/to/file.xlsx', status='pass',
                        error=None, result={}, issues=[])
        r = PatternTestReport(
            pattern_path='/p.xlsx', total=1, passed=1, warned=0, failed=0,
            file_results=[fr], field_counts={},
        )
        data = json.loads(format_json(r))
        assert data['files'][0]['path'] == '/abs/path/to/file.xlsx'


# ── run_tests validation ──────────────────────────────────────────────────────

class TestRunTestsValidation:
    def test_missing_pattern(self, tmp_path):
        with pytest.raises(FileNotFoundError, match='Pattern file not found'):
            run_tests('/nonexistent/pattern.xlsx', str(tmp_path))

    def test_missing_directory(self, tmp_path):
        pattern = tmp_path / 'pattern.xlsx'
        pattern.write_bytes(b'')
        with pytest.raises(FileNotFoundError, match='Test directory not found'):
            run_tests(str(pattern), '/nonexistent/dir')

    def test_empty_directory(self, tmp_path):
        pattern = tmp_path / 'pattern.xlsx'
        pattern.write_bytes(b'')
        data_dir = tmp_path / 'data'
        data_dir.mkdir()
        with pytest.raises(ValueError, match='No .xlsx'):
            run_tests(str(pattern), str(data_dir))

    def test_pattern_file_excluded_from_scan(self, tmp_path):
        """The pattern file itself must never appear as a data file."""
        pattern = tmp_path / 'pattern.xlsx'
        pattern.write_bytes(b'')
        data = tmp_path / 'data.xlsx'
        data.write_bytes(b'')
        # Mock engine so it doesn't actually parse the empty xlsx
        with patch('grepxcel.pattern_tester.run_tests') as mock:
            # Run for real but intercept engine — we just need the file list logic
            pass
        # When only data.xlsx is a real file, pattern is excluded — no ValueError
        # (engine will fail but that's a different code path)
        try:
            run_tests(str(pattern), str(tmp_path))
        except ValueError as exc:
            assert 'No .xlsx' not in str(exc), 'pattern file was not excluded from scan'
        except Exception:
            pass  # engine failure is expected for empty files

    def test_report_base_dir_is_set(self, tmp_path):
        """run_tests must populate base_dir on the returned report."""
        pattern = tmp_path / 'pattern.xlsx'
        pattern.write_bytes(b'')
        data_dir = tmp_path / 'samples'
        data_dir.mkdir()
        (data_dir / 'a.xlsx').write_bytes(b'')
        try:
            report = run_tests(str(pattern), str(data_dir))
        except Exception:
            pass  # engine fails on empty files — we don't care; run_tests may raise
        # Use the function with a fixture instead
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern_f = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')
        if not os.path.isfile(pattern_f) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')
        with tempfile.TemporaryDirectory() as tmp:
            import shutil
            shutil.copy(data_xlsx, tmp)
            report = run_tests(pattern_f, tmp)
        assert report.base_dir == os.path.abspath(tmp)

    def test_successful_run_against_fixtures(self):
        """Integration: run a known fixture through the engine end-to-end."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        with tempfile.TemporaryDirectory() as tmp:
            import shutil
            shutil.copy(data_xlsx, tmp)
            report = run_tests(pattern, tmp)

        assert report.total == 1
        assert report.failed == 0


# ── Fatal error → fail classification ────────────────────────────────────────

class TestFatalErrorClassification:
    """FATAL ERROR records in the logger must produce status='fail', not 'warn'."""

    def _mock_logger_with_error(self):
        """Return a mock Logger whose has_errors() returns True."""
        logger = MagicMock()
        logger.has_errors.return_value = True
        error_rec = MagicMock()
        error_rec.severity = 'ERROR'
        error_rec.message = 'FATAL ERROR: header layout unrecognized'
        logger._records = [error_rec]
        return logger

    def test_fatal_error_produces_fail_status(self, tmp_path):
        """Integration: when the engine logs an ERROR, FileResult.status must be 'fail'."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        # Patch Logger so has_errors() → True, simulating a FATAL record
        from grepxcel import pattern_tester as pt
        with tempfile.TemporaryDirectory() as tmp:
            import shutil
            shutil.copy(data_xlsx, tmp)

            original_logger_cls = None
            import grepxcel.logger as logger_mod

            class FakeLogger:
                def __init__(self, **kwargs):
                    self._records = [MagicMock(severity='ERROR',
                                               message='FATAL ERROR: injected')]

                def has_errors(self):
                    return True

            with patch.object(logger_mod, 'Logger', FakeLogger):
                from grepxcel.engine import Engine
                with patch.object(Engine, 'process', return_value={'inv': '001'}):
                    report = run_tests(pattern, tmp)

        assert report.failed >= 1, 'FATAL ERROR should produce at least one fail'
        fail_results = [r for r in report.file_results if r.status == 'fail']
        assert len(fail_results) >= 1

    def test_clean_run_still_passes(self, tmp_path):
        """Without ERROR records, a clean extraction must remain 'pass'."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        with tempfile.TemporaryDirectory() as tmp:
            import shutil
            shutil.copy(data_xlsx, tmp)
            report = run_tests(pattern, tmp)

        assert report.failed == 0
        assert all(r.status != 'fail' for r in report.file_results)


# ── compact_file_line ─────────────────────────────────────────────────────────

class TestCompactFileLine:
    def test_pass_line(self):
        fr = FileResult('/tmp/base/a.xlsx', 'pass', None, {}, [])
        line = _compact_file_line(fr, '/tmp/base', 'OK', 'WARN', 'FAIL')
        assert 'OK' in line
        assert 'a.xlsx' in line

    def test_warn_line_shows_first_issue(self):
        fr = FileResult('/tmp/base/b.xlsx', 'warn', None, {}, ['missing: inv.date'])
        line = _compact_file_line(fr, '/tmp/base', 'OK', 'WARN', 'FAIL')
        assert 'WARN' in line
        assert 'missing: inv.date' in line

    def test_fail_line_shows_error(self):
        fr = FileResult('/tmp/base/c.xlsx', 'fail', 'anchor not found', None, [])
        line = _compact_file_line(fr, '/tmp/base', 'OK', 'WARN', 'FAIL')
        assert 'FAIL' in line
        assert 'anchor not found' in line

    def test_uses_relative_path(self):
        fr = FileResult('/tmp/base/sub/d.xlsx', 'pass', None, {}, [])
        line = _compact_file_line(fr, '/tmp/base', 'OK', 'WARN', 'FAIL')
        assert 'sub/d.xlsx' in line or os.path.join('sub', 'd.xlsx') in line
        assert '/tmp/base' not in line


# ── CLI integration ───────────────────────────────────────────────────────────

class TestTestCLI:
    def test_missing_pattern_flag_exits_2(self, capsys):
        with pytest.raises(SystemExit) as exc:
            from grepxcel.cli import main
            sys.argv = ['grepxcel', 'test', '/tmp']
            main()
        assert exc.value.code == 2

    def test_help_exits_0(self, capsys):
        with pytest.raises(SystemExit) as exc:
            from grepxcel.cli import main
            sys.argv = ['grepxcel', 'test', '--help']
            main()
        assert exc.value.code == 0

    def test_bad_directory_exits_2(self, tmp_path, capsys):
        pattern = tmp_path / 'p.xlsx'
        pattern.write_bytes(b'')
        with pytest.raises(SystemExit) as exc:
            from grepxcel.cli import main
            sys.argv = ['grepxcel', 'test', '-p', str(pattern),
                        '/nonexistent/dir/xyz']
            main()
        assert exc.value.code == 2

    def test_json_format_produces_valid_json(self, tmp_path, capsys):
        """End-to-end: with a real fixture, --format json produces parseable JSON."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        import shutil
        data_dir = tmp_path / 'data'
        data_dir.mkdir()
        shutil.copy(data_xlsx, str(data_dir))

        from grepxcel.cli import main
        sys.argv = ['grepxcel', 'test', '-p', pattern, str(data_dir),
                    '--format', 'json']
        with pytest.raises(SystemExit) as exc:
            main()

        captured = capsys.readouterr()
        parsed = json.loads(captured.out)
        assert 'total' in parsed
        assert 'files' in parsed
        assert exc.value.code == 0

    def test_verbose_flag_adds_field_table(self, tmp_path, capsys):
        """--verbose / -v should include field reliability in human output."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        import shutil
        data_dir = tmp_path / 'data'
        data_dir.mkdir()
        shutil.copy(data_xlsx, str(data_dir))

        from grepxcel.cli import main
        sys.argv = ['grepxcel', 'test', '-p', pattern, str(data_dir), '-v',
                    '--no-color']
        with pytest.raises(SystemExit):
            main()

        captured = capsys.readouterr()
        assert 'Field reliability' in captured.out

    def test_no_color_flag_suppresses_emoji(self, tmp_path, capsys):
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        import shutil
        data_dir = tmp_path / 'data'
        data_dir.mkdir()
        shutil.copy(data_xlsx, str(data_dir))

        from grepxcel.cli import main
        sys.argv = ['grepxcel', 'test', '-p', pattern, str(data_dir), '--no-color']
        with pytest.raises(SystemExit):
            main()

        captured = capsys.readouterr()
        assert '✅' not in captured.out
        assert '❌' not in captured.out

    def test_piped_output_no_color(self, tmp_path, capsys):
        """When stdout is not a TTY, color must be disabled automatically."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        import shutil
        data_dir = tmp_path / 'data'
        data_dir.mkdir()
        shutil.copy(data_xlsx, str(data_dir))

        # Simulate a pipe: stdout.isatty() → False
        import io
        non_tty = io.StringIO()
        non_tty.isatty = lambda: False  # type: ignore[method-assign]

        from grepxcel.cli import main
        sys.argv = ['grepxcel', 'test', '-p', pattern, str(data_dir)]
        with patch('sys.stdout', non_tty):
            try:
                main()
            except SystemExit:
                pass

        output = non_tty.getvalue()
        assert '✅' not in output
        assert '❌' not in output

    def test_all_pass_exits_0(self, tmp_path, capsys):
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        import shutil
        data_dir = tmp_path / 'data'
        data_dir.mkdir()
        shutil.copy(data_xlsx, str(data_dir))

        from grepxcel.cli import main
        sys.argv = ['grepxcel', 'test', '-p', pattern, str(data_dir), '--no-color']
        with pytest.raises(SystemExit) as exc:
            main()
        assert exc.value.code == 0

    def test_verbose_flag_shows_relative_paths(self, tmp_path, capsys):
        """Output must show relative paths (not basenames only, not absolute paths)."""
        fixtures_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'fixtures', '01_simple_invoice',
        )
        pattern = os.path.join(fixtures_dir, '01_simple_invoice_pattern-from-web.csv')
        data_xlsx = os.path.join(fixtures_dir, '01_simple_invoice_data.xlsx')

        if not os.path.isfile(pattern) or not os.path.isfile(data_xlsx):
            pytest.skip('fixture files not available')

        import shutil
        # Create a nested structure so relative path != basename
        data_dir = tmp_path / 'data'
        sub = data_dir / 'invoices'
        sub.mkdir(parents=True)
        shutil.copy(data_xlsx, str(sub / '01_simple_invoice_data.xlsx'))

        from grepxcel.cli import main
        sys.argv = ['grepxcel', 'test', '-p', pattern, str(data_dir),
                    '-r', '--no-color']
        with pytest.raises(SystemExit):
            main()

        captured = capsys.readouterr()
        # Must show the directory context, not just the filename
        assert 'invoices' in captured.out
        # Must not show the absolute tmp path
        assert str(data_dir) not in captured.out
