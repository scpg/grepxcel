"""Unit tests for grepxcel.pattern_tester."""

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
    _count_present,
    _extract_field_names,
    format_human,
    format_json,
    run_tests,
)


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
            file_results=[],
            field_counts={'inv.number': (3, 3), 'inv.date': (2, 3)},
        )
        defaults.update(kwargs)
        return PatternTestReport(**defaults)

    def test_all_pass(self):
        r = self._report()
        out = format_human(r, color=False)
        assert 'All 3 passed' in out

    def test_shows_field_table(self):
        r = self._report()
        out = format_human(r, color=False)
        assert 'inv.number' in out
        assert '3/3' in out
        assert '100.0%' in out

    def test_shows_failures(self):
        fr = FileResult(
            path='/tmp/bad.xlsx',
            status='fail',
            error='anchor not found',
            result=None,
            issues=[],
        )
        r = self._report(total=1, passed=0, warned=0, failed=1, file_results=[fr])
        out = format_human(r, color=False)
        assert 'Failures' in out
        assert 'bad.xlsx' in out
        assert 'anchor not found' in out

    def test_no_color_flag(self):
        r = self._report()
        out_color = format_human(r, color=True)
        out_no_color = format_human(r, color=False)
        assert '✅' in out_color
        assert '✅' not in out_no_color


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
        with pytest.raises(ValueError, match='No .xlsx files found'):
            run_tests(str(pattern), str(data_dir))

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

        # Create a temp dir with just the one data file
        with tempfile.TemporaryDirectory() as tmp:
            import shutil
            shutil.copy(data_xlsx, tmp)
            report = run_tests(pattern, tmp)

        assert report.total == 1
        assert report.failed == 0


# ── CLI integration ───────────────────────────────────────────────────────────

class TestTestCLI:
    def test_missing_pattern_flag_exits_2(self, capsys):
        """--pattern is required; argparse exits 2 when it's missing."""
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
        # exit 0 = all pass
        assert exc.value.code == 0
