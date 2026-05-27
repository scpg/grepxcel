"""
Unit tests for cell:next and cell:A1 (absolute reference) addressing in Engine.
Uses in-memory xlsx workbooks to avoid fixture file dependencies.
"""

import openpyxl
import pytest

from engine.engine import Engine
from engine.logger import EngineError, Logger


def _make_data(cells: dict, tmp_path, filename='data.xlsx') -> str:
    """Create an xlsx with the given {coordinate: value} dict."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for coord, val in cells.items():
        ws[coord] = val
    path = str(tmp_path / filename)
    wb.save(path)
    return path


def _make_pattern(rows: list, tmp_path, filename='pattern.xlsx') -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / filename)
    wb.save(path)
    return path


def _run(pattern_rows, data_cells, tmp_path):
    pat = _make_pattern(pattern_rows, tmp_path)
    dat = _make_data(data_cells, tmp_path)
    return Engine().process(pat, dat)


# ── cell:next ─────────────────────────────────────────────────────────────────

class TestCellNext:
    def test_next_is_alias_for_cell1(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:next', 'x.val'], ['END:']],
            {'A1': 'hello'},
            tmp_path,
        )
        assert result == {'x': {'val': 'hello'}}

    def test_next_skips_empty_cells(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:next', 'x.val'], ['END:']],
            {'B3': 'found it'},  # A1, A2, B1, B2 are empty
            tmp_path,
        )
        assert result == {'x': {'val': 'found it'}}

    def test_cell1_still_works(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:1', 'x.val'], ['END:']],
            {'A1': 'old style'},
            tmp_path,
        )
        assert result == {'x': {'val': 'old style'}}


# ── cell:A1 — basic extraction ────────────────────────────────────────────────

class TestCellAbsoluteRef:
    def test_reads_specific_cell(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:B3', 'x.val'], ['END:']],
            {'B3': 'targeted', 'A1': 'not this one'},
            tmp_path,
        )
        assert result == {'x': {'val': 'targeted'}}

    def test_ignore_with_absolute_ref(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'IGNORE'],
             ['cell:next', 'x.val'],
             ['END:']],
            {'A1': 'skip me', 'B1': 'keep me'},
            tmp_path,
        )
        assert result == {'x': {'val': 'keep me'}}

    def test_cursor_advances_after_absolute_ref(self, tmp_path):
        # After cell:A1, cell:next should pick up the next non-empty cell after A1
        result = _run(
            [['var:', 'a.first', 'string', '.*'],
             ['var:', 'b.second', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'a.first'],
             ['cell:next', 'b.second'],
             ['END:']],
            {'A1': 'first', 'B1': 'second', 'C1': 'ignored'},
            tmp_path,
        )
        assert result == {'a': {'first': 'first'}, 'b': {'second': 'second'}}

    def test_mixing_absolute_and_sequential(self, tmp_path):
        result = _run(
            [['var:', 'a.v1', 'string', '.*'],
             ['var:', 'b.v2', 'string', '.*'],
             ['var:', 'c.v3', 'string', '.*'],
             ['START:'],
             ['cell:next', 'a.v1'],   # picks up A1
             ['cell:C1',   'b.v2'],   # jump to C1
             ['cell:next', 'c.v3'],   # next after C1 → D1
             ['END:']],
            {'A1': 'seq', 'B1': 'skipped', 'C1': 'jump', 'D1': 'after'},
            tmp_path,
        )
        assert result == {
            'a': {'v1': 'seq'},
            'b': {'v2': 'jump'},
            'c': {'v3': 'after'},
        }

    def test_lowercase_ref_normalized(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:b3', 'x.val'], ['END:']],
            {'B3': 'lower case ref'},
            tmp_path,
        )
        assert result == {'x': {'val': 'lower case ref'}}


# ── cell:A1 — empty cell handling ────────────────────────────────────────────

class TestCellAbsoluteEmpty:
    def test_var_empty_records_null(self, tmp_path):
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'], ['cell:C5', 'x.val'], ['END:']],
            {'A1': 'something else'},  # C5 is empty
            tmp_path,
        )
        assert result == {'x': {'val': None}}

    def test_lbl_empty_causes_fatal(self, tmp_path):
        pat = _make_pattern(
            [['lbl:', 'my_label', 'string', 'Expected Label'],
             ['START:'], ['cell:C5', 'my_label'], ['END:']],
            tmp_path,
        )
        dat = _make_data({'A1': 'something'}, tmp_path)
        logger = Logger()
        Engine().process(pat, dat, logger=logger)
        assert logger.has_errors()

    def test_ignore_empty_is_silent(self, tmp_path):
        # IGNORE on empty absolute ref should not error
        result = _run(
            [['var:', 'x.val', 'string', '.*'],
             ['START:'],
             ['cell:A1', 'IGNORE'],  # A1 is empty
             ['cell:next', 'x.val'],
             ['END:']],
            {'B1': 'value'},
            tmp_path,
        )
        assert result == {'x': {'val': 'value'}}


# ── cell:A1 — unreachable cell (cursor past target) ──────────────────────────

class TestCellAbsoluteUnreachable:
    def test_cursor_past_target_is_fatal(self, tmp_path):
        # cell:next consumes A1, then cell:A1 references A1 which is already past
        pat = _make_pattern(
            [['var:', 'x.first', 'string', '.*'],
             ['var:', 'y.second', 'string', '.*'],
             ['START:'],
             ['cell:next', 'x.first'],  # consumes A1; cursor moves to after A1
             ['cell:A1',   'y.second'],  # A1 is already before cursor → fatal
             ['END:']],
            tmp_path,
        )
        dat = _make_data({'A1': 'value', 'B1': 'other'}, tmp_path)
        logger = Logger()
        Engine().process(pat, dat, logger=logger)
        assert logger.has_errors()
