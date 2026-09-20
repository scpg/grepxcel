"""Tests for the generate-examples command."""

import json
import os
import pathlib
import sys

import pytest

import grepxcel
from grepxcel.cli import main as cli_main
from grepxcel.examples_generator import EXAMPLES, generate_examples
from tests.conftest import find_data_file, find_pattern_xlsx


FIXTURES_DIR = pathlib.Path(__file__).resolve().parent.parent / 'fixtures'


class TestExamplesMetadata:
    """EXAMPLES registry is well-formed."""

    def test_examples_has_four_entries(self):
        assert len(EXAMPLES) == 4

    def test_each_example_has_required_keys(self):
        for ex in EXAMPLES:
            assert 'name' in ex
            assert 'description' in ex
            assert 'pattern' in ex
            assert 'data' in ex

    def test_bundled_files_exist(self):
        for ex in EXAMPLES:
            assert os.path.isfile(ex['pattern']), f"missing: {ex['pattern']}"
            assert os.path.isfile(ex['data']), f"missing: {ex['data']}"


class TestGenerateExamples:
    """generate_examples() creates the expected directory structure."""

    def test_creates_output_directory(self, tmp_path):
        out = tmp_path / 'examples'
        generate_examples(str(out))
        assert out.is_dir()

    def test_creates_one_subdirectory_per_example(self, tmp_path):
        out = tmp_path / 'examples'
        generate_examples(str(out))
        subdirs = sorted(d.name for d in out.iterdir() if d.is_dir())
        assert len(subdirs) == 4

    def test_each_example_has_pattern_and_data(self, tmp_path):
        out = tmp_path / 'examples'
        generate_examples(str(out))
        for d in out.iterdir():
            if not d.is_dir():
                continue
            assert (d / 'pattern.xlsx').is_file(), f"missing pattern.xlsx in {d.name}"
            assert (d / 'data.xlsx').is_file(), f"missing data.xlsx in {d.name}"

    def test_each_example_has_readme(self, tmp_path):
        out = tmp_path / 'examples'
        generate_examples(str(out))
        for d in out.iterdir():
            if not d.is_dir():
                continue
            readme = d / 'README.txt'
            assert readme.is_file(), f"missing README.txt in {d.name}"
            text = readme.read_text()
            assert 'grepxcel extract' in text

    def test_pattern_files_are_valid_xlsx(self, tmp_path):
        import openpyxl
        out = tmp_path / 'examples'
        generate_examples(str(out))
        for d in out.iterdir():
            if not d.is_dir():
                continue
            wb = openpyxl.load_workbook(d / 'pattern.xlsx')
            assert wb.sheetnames

    def test_overwrites_existing_directory(self, tmp_path):
        """generate_examples() no longer guards; it overwrites freely."""
        out = tmp_path / 'examples'
        out.mkdir()
        (out / 'somefile.txt').write_text('block')
        generate_examples(str(out))  # must not raise
        assert any(out.iterdir())

    def test_cli_refuses_existing_directory_without_force(self, tmp_path):
        out = tmp_path / 'examples'
        out.mkdir()
        (out / 'somefile.txt').write_text('block')
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['generate-examples', '-o', str(out)])
        assert exc_info.value.code != 0

    def test_cli_force_overwrites_existing_directory(self, tmp_path):
        out = tmp_path / 'examples'
        out.mkdir()
        (out / 'somefile.txt').write_text('block')
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['generate-examples', '-o', str(out), '--force'])
        assert exc_info.value.code == 0
        assert (out / '01_simple_invoice').is_dir()

    def test_returns_output_path(self, tmp_path):
        out = tmp_path / 'examples'
        result = generate_examples(str(out))
        assert result == str(out)


class TestExtractionWorks:
    """Each bundled example produces valid JSON when extracted."""

    def test_extract_each_example(self, tmp_path):
        out = tmp_path / 'examples'
        generate_examples(str(out))
        for d in sorted(out.iterdir()):
            if not d.is_dir():
                continue
            pattern = str(d / 'pattern.xlsx')
            data = str(d / 'data.xlsx')
            result = grepxcel.extract(pattern, data)
            assert isinstance(result, dict)
            assert len(result) > 0, f"{d.name} produced empty output"


class TestCLISubcommand:
    """The generate-examples subcommand works via cli.main()."""

    def test_generate_examples_creates_directory(self, tmp_path):
        out = tmp_path / 'grepxcel-examples'
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['generate-examples', '-o', str(out)])
        assert exc_info.value.code == 0
        assert out.is_dir()
        subdirs = [d for d in out.iterdir() if d.is_dir()]
        assert len(subdirs) == 4

    def test_generate_examples_default_output(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['generate-examples'])
        assert exc_info.value.code == 0
        default_dir = tmp_path / 'grepxcel-examples'
        assert default_dir.is_dir()

    def test_help_text(self, capsys):
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['generate-examples', '--help'])
        assert exc_info.value.code == 0
        captured = capsys.readouterr()
        assert 'example' in captured.out.lower()


class TestBundledExamplesMatchFixtures:
    """Catch drift: bundled examples must stay in sync with source fixtures."""

    @pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
    def test_data_xlsx_matches_fixture(self, ex):
        fixture_data = pathlib.Path(find_data_file(str(FIXTURES_DIR / ex['source_fixture'])))
        bundled_data = pathlib.Path(ex['data'])
        src = ex['source_fixture']
        assert fixture_data.read_bytes() == bundled_data.read_bytes(), (
            f"Bundled {ex['name']}/data.xlsx differs from "
            f"tests/fixtures/{src}/{src}_data.xlsx — "
            f"copy the updated fixture into grepxcel/examples/{ex['name']}/"
        )

    @pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
    def test_pattern_xlsx_matches_fixture(self, ex):
        fixture_pattern = pathlib.Path(
            find_pattern_xlsx(str(FIXTURES_DIR / ex['source_fixture']))
        )
        bundled_pattern = pathlib.Path(ex['pattern'])
        src = ex['source_fixture']
        assert fixture_pattern.read_bytes() == bundled_pattern.read_bytes(), (
            f"Bundled {ex['name']}/pattern.xlsx differs from "
            f"tests/fixtures/{src}/{src}_pattern-from-draft.xlsx — "
            f"copy the updated fixture into grepxcel/examples/{ex['name']}/"
        )
