"""Unit tests for the grepxcel.extract() public facade."""

import os

import pytest

import grepxcel
from grepxcel import Engine, Logger, VerbosityLevel

_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '01_simple_invoice')
_PATTERN = os.path.join(_FIX, 'pattern.xlsx')
_DATA = os.path.join(_FIX, 'data.xlsx')


def test_extract_is_exported():
    assert hasattr(grepxcel, 'extract')
    assert 'extract' in grepxcel.__all__


def test_extract_matches_engine_process():
    """The facade returns the same result as the underlying Engine."""
    facade = grepxcel.extract(_PATTERN, _DATA)
    engine = Engine().process(
        _PATTERN, _DATA, logger=Logger(level=VerbosityLevel.QUIET),
    )
    assert facade == engine
    assert isinstance(facade, dict)


def test_extract_accepts_pathlib(tmp_path):
    from pathlib import Path
    result = grepxcel.extract(Path(_PATTERN), Path(_DATA))
    assert isinstance(result, dict) and result


def test_extract_sheet_and_all_sheets_are_mutually_exclusive():
    with pytest.raises(ValueError, match='not both'):
        grepxcel.extract(_PATTERN, _DATA, sheet='Sheet1', all_sheets=True)


def test_extract_all_sheets_returns_sheet_keyed_dict():
    result = grepxcel.extract(_PATTERN, _DATA, all_sheets=True)
    assert isinstance(result, dict)
    # the single-sheet invoice fixture -> one sheet key
    assert len(result) >= 1


def test_extract_is_silent_by_default(capsys):
    grepxcel.extract(_PATTERN, _DATA)
    captured = capsys.readouterr()
    assert captured.err == ''  # QUIET logger writes nothing to stderr
