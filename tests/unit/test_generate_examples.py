"""Tests for the generate-examples command."""

import json
import os
import pathlib
import subprocess
import sys

import pytest

from grepxcel.examples_generator import EXAMPLES, generate_examples


GREPXCEL = [sys.executable, '-m', 'grepxcel']

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

    def test_refuses_existing_directory(self, tmp_path):
        out = tmp_path / 'examples'
        out.mkdir()
        (out / 'somefile.txt').write_text('block')
        with pytest.raises(SystemExit):
            generate_examples(str(out))

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
            result = subprocess.run(
                [*GREPXCEL,'extract', '-p', pattern, data],
                capture_output=True, text=True, timeout=30,
            )
            assert result.returncode == 0, (
                f"{d.name} extraction failed:\n{result.stderr}"
            )
            parsed = json.loads(result.stdout)
            assert isinstance(parsed, dict)
            assert len(parsed) > 0, f"{d.name} produced empty output"


class TestCLISubcommand:
    """The generate-examples subcommand works end-to-end."""

    def test_generate_examples_creates_directory(self, tmp_path):
        out = tmp_path / 'grepxcel-examples'
        result = subprocess.run(
            [*GREPXCEL,'generate-examples', '-o', str(out)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        assert out.is_dir()
        subdirs = [d for d in out.iterdir() if d.is_dir()]
        assert len(subdirs) == 4

    def test_generate_examples_default_output(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        result = subprocess.run(
            [*GREPXCEL,'generate-examples'],
            capture_output=True, text=True, timeout=30,
            cwd=str(tmp_path),
        )
        assert result.returncode == 0
        default_dir = tmp_path / 'grepxcel-examples'
        assert default_dir.is_dir()

    def test_help_text(self):
        result = subprocess.run(
            [*GREPXCEL,'generate-examples', '--help'],
            capture_output=True, text=True, timeout=10,
        )
        assert result.returncode == 0
        assert 'example' in result.stdout.lower()


class TestBundledExamplesMatchFixtures:
    """Catch drift: bundled examples must stay in sync with source fixtures."""

    @pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
    def test_data_xlsx_matches_fixture(self, ex):
        fixture_data = FIXTURES_DIR / ex['source_fixture'] / 'data.xlsx'
        bundled_data = pathlib.Path(ex['data'])
        assert fixture_data.read_bytes() == bundled_data.read_bytes(), (
            f"Bundled {ex['name']}/data.xlsx differs from "
            f"tests/fixtures/{ex['source_fixture']}/data.xlsx — "
            f"copy the updated fixture into grepxcel/examples/{ex['name']}/"
        )

    @pytest.mark.parametrize('ex', EXAMPLES, ids=[e['name'] for e in EXAMPLES])
    def test_pattern_xlsx_matches_fixture(self, ex):
        fixture_pattern = (
            FIXTURES_DIR / ex['source_fixture'] / 'pattern-from-draft.xlsx'
        )
        bundled_pattern = pathlib.Path(ex['pattern'])
        assert fixture_pattern.read_bytes() == bundled_pattern.read_bytes(), (
            f"Bundled {ex['name']}/pattern.xlsx differs from "
            f"tests/fixtures/{ex['source_fixture']}/pattern-from-draft.xlsx — "
            f"copy the updated fixture into grepxcel/examples/{ex['name']}/"
        )
