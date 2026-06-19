"""Tests for grepxcel.cli._load_dotenv — scoped .env discovery.

The loader must not silently inherit a .env from an unrelated ancestor project:
discovery stops at the project root (a dir with .git or pyproject.toml), and a
.env found in an ancestor (not the cwd) is announced on stderr.
"""
import os

import pytest

from grepxcel.cli import _load_dotenv
import grepxcel.cli as cli


def _write(path, text):
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(text)


def _isolate_config_dir(monkeypatch, tmp_path):
    """Point the per-user config dir at an empty temp dir (hermetic)."""
    cfg = tmp_path / 'cfgdir'
    cfg.mkdir()
    monkeypatch.setattr(cli, '_config_dir', lambda: str(cfg))
    return cfg


def test_loads_env_from_cwd(tmp_path, monkeypatch):
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=from_cwd\n')
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'from_cwd'


def test_real_env_var_wins(tmp_path, monkeypatch):
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=from_file\n')
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv('GREPXCEL_TEST_KEY', 'from_real_env')
    _load_dotenv()
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'from_real_env'


def test_finds_env_in_project_root_from_subdir(tmp_path, monkeypatch):
    """A .env at the project root is found from a subdirectory."""
    (tmp_path / '.git').mkdir()
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=from_root\n')
    sub = tmp_path / 'a' / 'b'
    sub.mkdir(parents=True)
    monkeypatch.chdir(sub)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'from_root'


def test_does_not_escape_project_root(tmp_path, monkeypatch):
    """A .env ABOVE the project root must not be loaded — discovery stops at the
    boundary marked by .git/pyproject.toml."""
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=unrelated_ancestor\n')
    project = tmp_path / 'project'
    project.mkdir()
    (project / 'pyproject.toml').write_text('[project]\n')   # project-root marker
    monkeypatch.chdir(project)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    assert os.environ.get('GREPXCEL_TEST_KEY') is None   # ancestor .env ignored


def test_announces_ancestor_env(tmp_path, monkeypatch, capsys):
    """A .env found in an ancestor (not cwd) is announced on stderr."""
    (tmp_path / '.git').mkdir()
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=from_root\n')
    sub = tmp_path / 'sub'
    sub.mkdir()
    monkeypatch.chdir(sub)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    err = capsys.readouterr().err
    assert '.env' in err


def test_no_announcement_for_cwd_env(tmp_path, monkeypatch, capsys):
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=from_cwd\n')
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    err = capsys.readouterr().err
    assert err == ''


# ── config-dir .env (sanctioned global, lowest priority) ─────────────────────

def test_config_dir_env_loaded_as_fallback(tmp_path, monkeypatch):
    cfg = _isolate_config_dir(monkeypatch, tmp_path)
    _write(cfg / '.env', 'GREPXCEL_TEST_KEY=from_config_dir\n')
    work = tmp_path / 'work'         # loose cwd, no project .env
    work.mkdir()
    monkeypatch.chdir(work)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'from_config_dir'


def test_project_env_wins_over_config_dir(tmp_path, monkeypatch):
    cfg = _isolate_config_dir(monkeypatch, tmp_path)
    _write(cfg / '.env', 'GREPXCEL_TEST_KEY=from_config_dir\n')
    proj = tmp_path / 'proj'
    proj.mkdir()
    _write(proj / '.env', 'GREPXCEL_TEST_KEY=from_project\n')
    monkeypatch.chdir(proj)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv()
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'from_project'


# ── --strict-env / GREPXCEL_STRICT_ENV ───────────────────────────────────────

def test_strict_refuses_out_of_project_env(tmp_path, monkeypatch):
    _isolate_config_dir(monkeypatch, tmp_path)
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=ancestor\n')  # loose ancestor
    sub = tmp_path / 'sub'
    sub.mkdir()
    monkeypatch.chdir(sub)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    with pytest.raises(SystemExit):
        _load_dotenv(strict=True)


def test_strict_env_var_also_triggers(tmp_path, monkeypatch):
    _isolate_config_dir(monkeypatch, tmp_path)
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=ancestor\n')
    sub = tmp_path / 'sub'
    sub.mkdir()
    monkeypatch.chdir(sub)
    monkeypatch.setenv('GREPXCEL_STRICT_ENV', '1')
    with pytest.raises(SystemExit):
        _load_dotenv()


def test_strict_allows_config_dir_env(tmp_path, monkeypatch):
    """The sanctioned config-dir .env is never refused, even under strict."""
    cfg = _isolate_config_dir(monkeypatch, tmp_path)
    _write(cfg / '.env', 'GREPXCEL_TEST_KEY=from_config_dir\n')
    work = tmp_path / 'work'
    work.mkdir()
    monkeypatch.chdir(work)
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv(strict=True)        # must NOT raise
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'from_config_dir'


def test_strict_allows_project_local_env(tmp_path, monkeypatch):
    _isolate_config_dir(monkeypatch, tmp_path)
    _write(tmp_path / '.env', 'GREPXCEL_TEST_KEY=local\n')
    monkeypatch.chdir(tmp_path)       # .env is in cwd → in-project
    monkeypatch.delenv('GREPXCEL_TEST_KEY', raising=False)
    _load_dotenv(strict=True)        # must NOT raise
    assert os.environ.get('GREPXCEL_TEST_KEY') == 'local'
