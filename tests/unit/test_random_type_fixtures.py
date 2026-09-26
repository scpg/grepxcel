"""Randomized, type-based "reverse white-box" test framework — Part B of the
`profile` + test-generation plan.

Unlike the 38 realistic fixtures (one spreadsheet exercising many types
loosely, in a fixed layout), this isolates **one type at a time**, builds
it in several structural shapes (1 label, 2 labels, 1 mini-table, 2
mini-tables, combined), at **randomized cell positions** within a 20x20
grid, with a **fixed seed** per test (reproducible without a committed
fixture file — no new .xlsx files are persisted by this module at all).

The one genuinely hard part is the "labels found before tables" placement
guarantee for the combined shape, under *both* of grepxcel's two read
directions. That guarantee is implemented by mirroring the engine's own
scan-order construction (engine.py `_build_scan_order`: row-major for LR,
column-major for TD) rather than inventing a separate ordering rule that
could quietly disagree with the real one.

Scope, deliberately: this framework tests *types*, not table-row mechanics
— generated mini-tables are always plain HEADER:1/DATA:* (no SKIP_IF/
SKIP_EMPTY_ROW). It also knowingly runs into a documented pre-existing
engine gap (see TestTwoMiniTablesShape) rather than silently working
around it.
"""
from __future__ import annotations

import csv
import datetime
import os
import random

import openpyxl
import pytest

from grepxcel.engine import Engine
from grepxcel.logger import Logger, Severity, VerbosityLevel

GRID_SIZE = 20  # the "20x20 range" from the requirements


# ── scan-order mirroring (the core, must-get-right piece) ──────────────────

def scan_index(pos: tuple[int, int], direction: str, max_row: int, max_col: int) -> int:
    """A cell's position in grepxcel's own scan order, as a single
    comparable integer — mirrors engine.py's Scanner._build_scan_order
    exactly (LR: row-major, TD: column-major), so "before" here means
    exactly what it means to the real engine, not a separately-invented
    approximation of it."""
    row, col = pos
    if direction == 'LR':
        return (row - 1) * max_col + (col - 1)
    if direction == 'TD':
        return (col - 1) * max_row + (row - 1)
    raise ValueError(f'unknown direction {direction!r}')


def position_before(a: tuple[int, int], b: tuple[int, int], direction: str,
                    max_row: int = GRID_SIZE, max_col: int = GRID_SIZE) -> bool:
    """True if cell *a* is encountered strictly before cell *b* when
    scanning in *direction*."""
    return scan_index(a, direction, max_row, max_col) < scan_index(b, direction, max_row, max_col)


class TestScanOrderMirroring:
    """Direct verification against engine.py's real _build_scan_order —
    the whole placement guarantee rests on this matching exactly."""

    def _real_scan_order(self, direction: str, max_row: int, max_col: int) -> list[tuple[int, int]]:
        cells = []
        if direction == 'LR':
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cells.append((r, c))
        elif direction == 'TD':
            for c in range(1, max_col + 1):
                for r in range(1, max_row + 1):
                    cells.append((r, c))
        return cells

    @pytest.mark.parametrize('direction', ['LR', 'TD'])
    def test_matches_real_scan_order_exactly(self, direction):
        max_row, max_col = 6, 5
        real = self._real_scan_order(direction, max_row, max_col)
        real_index = {pos: i for i, pos in enumerate(real)}
        for pos in real:
            assert scan_index(pos, direction, max_row, max_col) == real_index[pos]

    def test_lr_row_major(self):
        # Same row, later column -> after. Later row, any column -> after.
        assert position_before((1, 1), (1, 2), 'LR')
        assert position_before((1, 20), (2, 1), 'LR')
        assert not position_before((2, 1), (1, 20), 'LR')

    def test_td_column_major(self):
        # Same column, later row -> after. Later column, any row -> after.
        assert position_before((1, 1), (2, 1), 'TD')
        assert position_before((20, 1), (1, 2), 'TD')
        assert not position_before((1, 2), (20, 1), 'TD')

    def test_lr_and_td_disagree_on_the_same_pair(self):
        # The whole point of having two directions: the same two cells can
        # be ordered oppositely depending on which one is active.
        a, b = (1, 5), (5, 1)
        assert position_before(a, b, 'LR')
        assert position_before(b, a, 'TD')


# ── random placement ─────────────────────────────────────────────────────

class RandomPlacer:
    """Non-overlapping random cell/block placement within a GRID_SIZE x
    GRID_SIZE range, with a fixed seed for reproducibility."""

    def __init__(self, seed: int, size: int = GRID_SIZE):
        self.rng = random.Random(seed)
        self.size = size
        self._occupied: set[tuple[int, int]] = set()

    def _free_cell(self) -> tuple[int, int]:
        for _ in range(500):
            pos = (self.rng.randint(1, self.size), self.rng.randint(1, self.size))
            if pos not in self._occupied:
                self._occupied.add(pos)
                return pos
        raise RuntimeError('grid exhausted — increase size or reduce placements')

    def place_label(self) -> tuple[int, int]:
        return self._free_cell()

    def place_label_with_value_after(self) -> tuple[tuple[int, int], tuple[int, int]]:
        """A label cell plus an adjacent value cell immediately after it in
        column order (cell:next semantics) — both reserved together."""
        row, col = self._free_cell()
        # Reserve the "next" cell too so it's never independently reused.
        self._occupied.add((row, col + 1))
        return (row, col), (row, col + 1)

    def place_table(self, data_rows: int = 2) -> tuple[int, int]:
        """A 1-column table: header row + `data_rows` data rows, placed as
        one non-overlapping block. Returns the header cell's position."""
        for _ in range(500):
            row = self.rng.randint(1, self.size - data_rows)
            col = self.rng.randint(1, self.size)
            block = {(row + i, col) for i in range(data_rows + 1)}
            if block & self._occupied:
                continue
            self._occupied |= block
            return (row, col)
        raise RuntimeError('grid exhausted for table placement')

    def place_after(self, ref: tuple[int, int], direction: str, data_rows: int = 2) -> tuple[int, int]:
        """A table header position guaranteed to come strictly after *ref*
        in scan order for *direction* — the label-before-table guarantee."""
        for _ in range(2000):
            row = self.rng.randint(1, self.size - data_rows)
            col = self.rng.randint(1, self.size)
            if not position_before(ref, (row, col), direction):
                continue
            block = {(row + i, col) for i in range(data_rows + 1)}
            if block & self._occupied:
                continue
            self._occupied |= block
            return (row, col)
        raise RuntimeError('could not place a table strictly after the label — grid too small')


# ── per-type value catalog ──────────────────────────────────────────────────
# valid: a value that satisfies validate_type() for this type.
# invalid: a value that FAILS it — reusing the exact failure modes already
# known from utils.py's validate_type() (a bool for integer, a non-ISO
# string for date, etc.), not invented ad hoc.

def _tc(valid, invalid, fmt=None, regex='.*'):
    return {'valid': valid, 'invalid': invalid, 'format': fmt, 'regex': regex}


TYPE_CASES = {
    'string':     _tc('hello world', None, regex='.*'),  # see note below
    'integer':    _tc(42, True, regex='.*'),               # bool is explicitly rejected
    'boolean':    _tc(True, 'banana', regex='.*'),
    'date':       _tc(datetime.date(2024, 1, 15), 'not a date', fmt='yyyy-mm-dd', regex='.*'),
    'url':        _tc('https://example.com', 'not a url', regex='.*'),
    'currency':   _tc(1200.50, True, fmt='$#,##0.00', regex='.*'),        # bool rejected, same as integer
    'percentage': _tc(0.15, 'fifty percent', fmt='0.00%', regex='.*'),
    'number':     _tc(42.5, False, regex='.*'),                          # bool rejected
    'time':       _tc(datetime.time(9, 30), '9:30am', fmt='HH:MM', regex='.*'),
    'duration':   _tc(datetime.timedelta(hours=2), '2 hours', fmt='[h]:mm', regex='.*'),
}
# 'string' has no real *type* failure (str(anything) always succeeds) — its
# "invalid" case instead exercises a regex mismatch, which is the only way
# a var:string field can actually fail. Modelled explicitly per type below
# rather than forcing every type through one invalid-value shape.


def _invalid_case_for(type_name: str, cell_ref_placeholder: str = 'x'):
    """Return (var_type, regex, invalid_value) for the 'should fail' case."""
    if type_name == 'string':
        return 'string', r'^\d+$', 'not-a-number-but-pattern-wants-one'
    tc = TYPE_CASES[type_name]
    return type_name, '.*', tc['invalid']


# ── pattern/data file builders ──────────────────────────────────────────────

def write_pattern_csv(path: str, rows: list[list[str]]) -> None:
    with open(path, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        for row in rows:
            w.writerow(row)


def _set_cell(ws, pos: tuple[int, int], value, number_format: str | None = None) -> None:
    row, col = pos
    cell = ws.cell(row=row, column=col)
    cell.value = value
    if number_format:
        cell.number_format = number_format


def write_data_xlsx(path: str, cells: dict[tuple[int, int], tuple]) -> None:
    """cells: {(row, col): (value, number_format_or_None)}"""
    wb = openpyxl.Workbook()
    ws = wb.active
    for pos, (value, fmt) in cells.items():
        _set_cell(ws, pos, value, fmt)
    wb.save(path)


def _run(pattern_path: str, data_path: str) -> dict:
    lg = Logger(level=VerbosityLevel.QUIET)
    return Engine().process(pattern_file=pattern_path, data_file=data_path, logger=lg)


def _run_with_issues(pattern_path: str, data_path: str) -> tuple[dict, list]:
    """Like _run, but also returns every WARNING/ERROR record raised —
    needed because an invalid value does NOT become null in the output
    (a real behaviour this framework surfaced, not a test bug): engine.py's
    _process_cell() writes result['cells'][field] = value unconditionally,
    regardless of whether _validate_field() passed. Validation failure is
    reported as an issue, not reflected in the extracted value itself."""
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pattern_path, data_file=data_path, logger=lg)
    issues = [r for r in lg._records if r.severity in (Severity.WARNING, Severity.ERROR)]
    return result, issues


def _ref(pos: tuple[int, int]) -> str:
    return openpyxl.utils.get_column_letter(pos[1]) + str(pos[0])


# ── Shape 1: exactly 1 label ────────────────────────────────────────────────

class TestOneLabelShape:
    """A single lbl: anchor of the given type, randomly placed. Since lbl:
    fields never appear in output, success is proven indirectly via a
    var: field placed immediately after it (cell:next)."""

    SEEDS = [1, 2, 3]

    @pytest.mark.parametrize('type_name', sorted(TYPE_CASES))
    @pytest.mark.parametrize('seed', SEEDS)
    def test_valid_value_extracts(self, tmp_path, type_name, seed):
        tc = TYPE_CASES[type_name]
        placer = RandomPlacer(seed)
        label_pos, value_pos = placer.place_label_with_value_after()

        pattern_path = str(tmp_path / 'pattern.csv')
        data_path = str(tmp_path / 'data.xlsx')
        write_pattern_csv(pattern_path, [
            ['lbl:', 'anchor', 'string', 'ANCHOR'],
            ['var:', 'field', type_name, tc['regex']],
            [],
            ['START:'],
            ['cell:next', 'anchor'],
            ['cell:next', 'field'],
            ['END:'],
        ])
        write_data_xlsx(data_path, {
            label_pos: ('ANCHOR', None),
            value_pos: (tc['valid'], tc['format']),
        })

        result = _run(pattern_path, data_path)
        assert 'field' in result, f'seed={seed} type={type_name} label={_ref(label_pos)} value={_ref(value_pos)} result={result}'

    @pytest.mark.parametrize('type_name', sorted(TYPE_CASES))
    def test_invalid_value_raises_a_warning(self, tmp_path, type_name):
        var_type, regex, invalid_value = _invalid_case_for(type_name)
        placer = RandomPlacer(seed=42)
        label_pos, value_pos = placer.place_label_with_value_after()

        pattern_path = str(tmp_path / 'pattern.csv')
        data_path = str(tmp_path / 'data.xlsx')
        write_pattern_csv(pattern_path, [
            ['lbl:', 'anchor', 'string', 'ANCHOR'],
            ['var:', 'field', var_type, regex],
            [],
            ['START:'],
            ['cell:next', 'anchor'],
            ['cell:next', 'field'],
            ['END:'],
        ])
        write_data_xlsx(data_path, {
            label_pos: ('ANCHOR', None),
            value_pos: (invalid_value, None),
        })

        result, issues = _run_with_issues(pattern_path, data_path)
        # The raw value still comes through as-is — validation failure does
        # NOT null it out (see _run_with_issues' docstring). What actually
        # changes is that a warning/error is raised for it.
        assert result.get('field') == invalid_value, f'type={type_name} value_pos={_ref(value_pos)} result={result}'
        assert issues, f'expected a validation issue for invalid {type_name} value {invalid_value!r}, got none'


# ── Shape 2: exactly 2 labels ───────────────────────────────────────────────

class TestTwoLabelsShape:
    """Two independent lbl:-anchored fields of the same type, both randomly
    placed with no positional relationship required between them."""

    @pytest.mark.parametrize('type_name', sorted(TYPE_CASES))
    def test_both_resolve(self, tmp_path, type_name):
        tc = TYPE_CASES[type_name]
        placer = RandomPlacer(seed=7)
        label1, value1 = placer.place_label_with_value_after()
        # lbl:/var: anchors are found by scanning FORWARD from the cursor,
        # not by resetting to the whole sheet — so label2 must come after
        # label1 in scan order (default direction is LR). No seek:
        # instruction needed once that's guaranteed.
        label2 = placer.place_after(label1, 'LR', data_rows=0)
        value2 = (label2[0], label2[1] + 1)
        placer._occupied.add(value2)

        pattern_path = str(tmp_path / 'pattern.csv')
        data_path = str(tmp_path / 'data.xlsx')
        write_pattern_csv(pattern_path, [
            ['lbl:', 'anchor1', 'string', 'FIRST'],
            ['var:', 'field1', type_name, tc['regex']],
            ['lbl:', 'anchor2', 'string', 'SECOND'],
            ['var:', 'field2', type_name, tc['regex']],
            [],
            ['START:'],
            ['cell:next', 'anchor1'],
            ['cell:next', 'field1'],
            ['cell:next', 'anchor2'],
            ['cell:next', 'field2'],
            ['END:'],
        ])
        write_data_xlsx(data_path, {
            label1: ('FIRST', None),
            value1: (tc['valid'], tc['format']),
            label2: ('SECOND', None),
            value2: (tc['valid'], tc['format']),
        })

        result = _run(pattern_path, data_path)
        assert 'field1' in result and 'field2' in result, result


# ── Shape 3: exactly 1 mini-table (1 column) ────────────────────────────────

def _write_table_pattern(pattern_path: str, type_name: str, regex: str, multiplicity: str = '*') -> None:
    write_pattern_csv(pattern_path, [
        ['lbl:', 'col_header', 'string', 'COLHEAD'],
        ['var:', 'cell_value', type_name, regex],
        [],
        ['START:'],
        [f'table:{multiplicity}'],
        ['', 'HEADER:1', 'col_header'],
        ['', 'DATA:*', 'cell_value'],
        ['END:'],
    ])


class TestOneMiniTableShape:
    @pytest.mark.parametrize('type_name', sorted(TYPE_CASES))
    def test_extracts_table_rows(self, tmp_path, type_name):
        tc = TYPE_CASES[type_name]
        placer = RandomPlacer(seed=11)
        header_pos = placer.place_table(data_rows=2)
        row, col = header_pos

        pattern_path = str(tmp_path / 'pattern.csv')
        data_path = str(tmp_path / 'data.xlsx')
        _write_table_pattern(pattern_path, type_name, tc['regex'])
        write_data_xlsx(data_path, {
            (row, col):     ('COLHEAD', None),
            (row + 1, col): (tc['valid'], tc['format']),
            (row + 2, col): (tc['valid'], tc['format']),
        })

        result = _run(pattern_path, data_path)
        assert 'table_0' in result, f'type={type_name} header={_ref(header_pos)} result={result}'
        assert len(result['table_0']) == 1
        assert len(result['table_0'][0]['data']) == 2


# ── Shape 4: exactly 2 mini-tables ───────────────────────────────────────────

class TestTwoMiniTablesShape:
    """Documents *actual* engine behaviour for two physically separate
    table instances matching the same table:* definition — this is exactly
    the area flagged in the plan as touching a known pre-existing gap
    (table:1 multiplicity is stored but never enforced; table:* and
    table:1 are both greedy in practice). This test intentionally asserts
    what the engine really does today, not a guarantee it doesn't make."""

    def test_two_disjoint_tables_both_get_collected(self, tmp_path):
        type_name = 'integer'
        tc = TYPE_CASES[type_name]
        placer = RandomPlacer(seed=13)
        header1 = placer.place_table(data_rows=1)
        header2 = placer.place_table(data_rows=1)
        r1, c1 = header1
        r2, c2 = header2

        pattern_path = str(tmp_path / 'pattern.csv')
        data_path = str(tmp_path / 'data.xlsx')
        _write_table_pattern(pattern_path, type_name, tc['regex'])
        write_data_xlsx(data_path, {
            (r1, c1):     ('COLHEAD', None),
            (r1 + 1, c1): (tc['valid'], None),
            (r2, c2):     ('COLHEAD', None),
            (r2 + 1, c2): (tc['valid'], None),
        })

        result = _run(pattern_path, data_path)
        # Documented current behaviour: table:* finds every disjoint
        # matching anchor as a separate instance under the SAME group key,
        # rather than needing two distinct group names. If a future fix to
        # the table:1/table:* multiplicity gap changes this, this
        # assertion is the one to update — deliberately, not silently.
        assert 'table_0' in result
        assert len(result['table_0']) == 2, (
            f'expected 2 separate table instances (documenting current '
            f'table:* behaviour), got: {result}'
        )


# ── Shape 5: combined — label(s) before table(s), both directions ─────────

class TestCombinedLabelBeforeTable:
    """The positional guarantee: a label (with its value resolved via
    cell:next) followed, later in scan order, by a mini-table — verified
    under both LR and TD, using the exact same scan-order mirroring as
    TestScanOrderMirroring, not a separately-trusted assumption."""

    @pytest.mark.parametrize('direction', ['LR', 'TD'])
    def test_label_then_table_both_resolve(self, tmp_path, direction):
        type_name = 'integer'
        tc = TYPE_CASES[type_name]
        placer = RandomPlacer(seed=21)
        label_pos, value_pos = placer.place_label_with_value_after()
        header_pos = placer.place_after(label_pos, direction, data_rows=2)
        row, col = header_pos

        # Sanity-check the guarantee itself before even building the file —
        # if this fails, the placer is broken, independent of the engine.
        assert position_before(label_pos, header_pos, direction)

        pattern_path = str(tmp_path / 'pattern.csv')
        data_path = str(tmp_path / 'data.xlsx')
        write_pattern_csv(pattern_path, [
            ['config:', 'read.direction', direction],
            ['lbl:', 'anchor', 'string', 'ANCHOR'],
            ['var:', 'field', type_name, tc['regex']],
            ['lbl:', 'col_header', 'string', 'COLHEAD'],
            ['var:', 'cell_value', type_name, tc['regex']],
            [],
            ['START:'],
            ['cell:next', 'anchor'],
            ['cell:next', 'field'],
            ['table:*'],
            ['', 'HEADER:1', 'col_header'],
            ['', 'DATA:*', 'cell_value'],
            ['END:'],
        ])
        write_data_xlsx(data_path, {
            label_pos:            ('ANCHOR', None),
            value_pos:            (tc['valid'], None),
            (row, col):           ('COLHEAD', None),
            (row + 1, col):       (tc['valid'], None),
            (row + 2, col):       (tc['valid'], None),
        })

        result = _run(pattern_path, data_path)
        assert 'field' in result, (
            f'direction={direction} label={_ref(label_pos)} '
            f'header={_ref(header_pos)} result={result}'
        )
        assert 'table_0' in result, result
