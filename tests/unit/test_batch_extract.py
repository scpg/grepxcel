"""Tests for batch/directory extraction (CLI _expand_files + -r flag)."""

import json
import os
import shutil

import pytest

from grepxcel.cli import _expand_files, _DEFAULT_MAX_FILES, main as cli_main
from tests.conftest import find_data_file, find_pattern_xlsx

FIXTURES = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


def _fixture(*parts):
    return os.path.join(FIXTURES, *parts)


# ── _expand_files unit tests ─────────────────────────────────────────────────


class TestExpandFiles:
    """The _expand_files helper expands directories to .xlsx files."""

    def test_plain_files_unchanged(self):
        files = ['a.xlsx', 'b.xlsx']
        assert _expand_files(files) == files

    def test_directory_lists_xlsx(self, tmp_path):
        (tmp_path / 'alpha.xlsx').touch()
        (tmp_path / 'beta.xlsx').touch()
        (tmp_path / 'readme.txt').touch()
        result = _expand_files([str(tmp_path)])
        assert len(result) == 2
        assert all(f.endswith('.xlsx') for f in result)

    def test_directory_sorted(self, tmp_path):
        (tmp_path / 'z.xlsx').touch()
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'm.xlsx').touch()
        result = _expand_files([str(tmp_path)])
        names = [os.path.basename(f) for f in result]
        assert names == ['a.xlsx', 'm.xlsx', 'z.xlsx']

    def test_directory_ignores_temp_files(self, tmp_path):
        (tmp_path / 'data.xlsx').touch()
        (tmp_path / '~$data.xlsx').touch()
        result = _expand_files([str(tmp_path)])
        assert len(result) == 1
        assert '~$' not in result[0]

    def test_directory_case_insensitive_extension(self, tmp_path):
        (tmp_path / 'upper.XLSX').touch()
        (tmp_path / 'mixed.Xlsx').touch()
        result = _expand_files([str(tmp_path)])
        assert len(result) == 2

    def test_empty_directory(self, tmp_path):
        result = _expand_files([str(tmp_path)])
        assert result == []

    def test_no_recursion_by_default(self, tmp_path):
        sub = tmp_path / 'sub'
        sub.mkdir()
        (tmp_path / 'top.xlsx').touch()
        (sub / 'nested.xlsx').touch()
        result = _expand_files([str(tmp_path)], recursive=False)
        assert len(result) == 1
        assert 'top.xlsx' in result[0]

    def test_recursive_finds_nested(self, tmp_path):
        sub = tmp_path / 'sub'
        sub.mkdir()
        deep = sub / 'deep'
        deep.mkdir()
        (tmp_path / 'top.xlsx').touch()
        (sub / 'mid.xlsx').touch()
        (deep / 'bottom.xlsx').touch()
        result = _expand_files([str(tmp_path)], recursive=True)
        assert len(result) == 3
        names = [os.path.basename(f) for f in result]
        assert 'top.xlsx' in names
        assert 'mid.xlsx' in names
        assert 'bottom.xlsx' in names

    def test_mixed_files_and_dirs(self, tmp_path):
        scan_dir = tmp_path / 'scan'
        scan_dir.mkdir()
        (scan_dir / 'dir_file.xlsx').touch()
        explicit = str(tmp_path / 'explicit.xlsx')
        open(explicit, 'w').close()
        result = _expand_files([explicit, str(scan_dir)])
        assert len(result) == 2

    def test_dir_without_xlsx_returns_empty(self, tmp_path):
        (tmp_path / 'readme.md').touch()
        (tmp_path / 'data.csv').touch()
        result = _expand_files([str(tmp_path)])
        assert result == []

    def test_recursive_ignores_empty_subdirs(self, tmp_path):
        (tmp_path / 'empty_a').mkdir()
        (tmp_path / 'empty_b').mkdir()
        has_files = tmp_path / 'has_files'
        has_files.mkdir()
        (has_files / 'report.xlsx').touch()
        result = _expand_files([str(tmp_path)], recursive=True)
        assert len(result) == 1
        assert 'report.xlsx' in result[0]

    def test_max_files_cap_aborts(self, tmp_path):
        for i in range(5):
            (tmp_path / f'file_{i}.xlsx').touch()
        with pytest.raises(SystemExit):
            _expand_files([str(tmp_path)], max_files=3)

    def test_max_files_at_limit_succeeds(self, tmp_path):
        for i in range(3):
            (tmp_path / f'file_{i}.xlsx').touch()
        result = _expand_files([str(tmp_path)], max_files=3)
        assert len(result) == 3

    def test_default_max_files_is_sensible(self):
        assert _DEFAULT_MAX_FILES == 10_000

    def test_symlink_file_skipped(self, tmp_path, capsys):
        real = tmp_path / 'real.xlsx'
        real.touch()
        link = tmp_path / 'link.xlsx'
        link.symlink_to(real)
        result = _expand_files([str(tmp_path)])
        assert len(result) == 1
        assert 'real.xlsx' in result[0]
        err = capsys.readouterr().err
        assert 'symlink' in err.lower()

    def test_symlink_dir_skipped_in_recursive(self, tmp_path, capsys):
        real_dir = tmp_path / 'real'
        real_dir.mkdir()
        (real_dir / 'data.xlsx').touch()
        link_dir = tmp_path / 'link_dir'
        link_dir.symlink_to(real_dir)
        result = _expand_files([str(tmp_path)], recursive=True)
        assert len(result) == 1
        assert 'real' in result[0]
        err = capsys.readouterr().err
        assert 'symlink' in err.lower()

    def test_symlink_as_top_level_argument_skipped(self, tmp_path, capsys):
        real = tmp_path / 'real_dir'
        real.mkdir()
        (real / 'data.xlsx').touch()
        link = tmp_path / 'link_arg'
        link.symlink_to(real)
        result = _expand_files([str(link)])
        assert result == []
        err = capsys.readouterr().err
        assert 'symlink' in err.lower()


# ── CLI integration tests ────────────────────────────────────────────────────


class TestBatchExtractCLI:
    """End-to-end tests for directory-based extraction via CLI."""

    def test_directory_argument(self, tmp_path):
        pattern = find_pattern_xlsx(_fixture('01_simple_invoice'))
        data_src = find_data_file(_fixture('01_simple_invoice'))
        data_dst = tmp_path / 'data.xlsx'
        shutil.copy2(data_src, data_dst)

        out_dir = str(tmp_path / 'output')
        with pytest.raises(SystemExit) as exc_info:
            cli_main([
                'extract', '-p', pattern,
                str(tmp_path),
                '-o', out_dir,
            ])
        assert exc_info.value.code == 0
        jsons = [f for f in os.listdir(out_dir) if f.endswith('.json')]
        assert len(jsons) == 1

    def test_recursive_flag(self, tmp_path):
        pattern = find_pattern_xlsx(_fixture('01_simple_invoice'))
        data_src = find_data_file(_fixture('01_simple_invoice'))

        sub = tmp_path / 'sub'
        sub.mkdir()
        shutil.copy2(data_src, tmp_path / 'top.xlsx')
        shutil.copy2(data_src, sub / 'nested.xlsx')

        out_dir = str(tmp_path / 'output')
        with pytest.raises(SystemExit) as exc_info:
            cli_main([
                'extract', '-p', pattern,
                str(tmp_path), '-r',
                '-o', out_dir,
            ])
        assert exc_info.value.code == 0
        jsons = [f for f in os.listdir(out_dir) if f.endswith('.json')]
        assert len(jsons) == 2

    def test_empty_directory_exits_nonzero(self, tmp_path):
        pattern = find_pattern_xlsx(_fixture('01_simple_invoice'))
        empty = tmp_path / 'empty'
        empty.mkdir()
        with pytest.raises(SystemExit) as exc_info:
            cli_main([
                'extract', '-p', pattern,
                str(empty),
            ])
        assert exc_info.value.code == 1

    def test_temp_files_skipped_in_directory(self, tmp_path):
        pattern = find_pattern_xlsx(_fixture('01_simple_invoice'))
        data_src = find_data_file(_fixture('01_simple_invoice'))
        shutil.copy2(data_src, tmp_path / 'data.xlsx')
        shutil.copy2(data_src, tmp_path / '~$data.xlsx')

        out_dir = str(tmp_path / 'output')
        with pytest.raises(SystemExit) as exc_info:
            cli_main([
                'extract', '-p', pattern,
                str(tmp_path),
                '-o', out_dir,
            ])
        assert exc_info.value.code == 0
        jsons = [f for f in os.listdir(out_dir) if f.endswith('.json')]
        assert len(jsons) == 1
