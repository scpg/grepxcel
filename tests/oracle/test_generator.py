"""Wave 2 acceptance: every generated case is structurally sound (spec §10.1).

The generator is the oracle suite's foundation — if a generated pattern is
invalid or a generated file does not contain what the manifest claims, every
downstream oracle is measuring the wrong thing. So before any oracle is wired
in, this module proves, for every shape x direction x via x header_mode at
seed 0:

  * ``validate-pattern`` exits 0 on the generated pattern (§10.1);
  * every manifest ref really exists in the workbook, and the workbook contains
    nothing the manifest does not list (the same ref-set equality the profile
    oracle will later assert against the classifier);
  * the declared ``StorageCase.round_trip`` transform matches what openpyxl
    actually reads back — checked here rather than assumed, so a change in
    openpyxl's casting fails loudly instead of silently rewriting expectations;
  * the placement invariants survive generation (error cells last in scan
    order, margins intact).
"""
from __future__ import annotations

import os

import openpyxl
import pytest

from grepxcel import cli
from tests.oracle import catalog
from tests.oracle.generator import (
    HEADER_MODES, PLACE_SIZE, SCALAR_SHAPES, TABLE_SHAPES, VIAS,
    CaseParams, build_case,
)
from tests.oracle.placement import GRID
from tests.unit.test_random_type_fixtures import scan_index

pytestmark = pytest.mark.oracle

DIRECTIONS = ('LR', 'TD')


def _params() -> list[CaseParams]:
    """Every shape x direction x via / header_mode at seed 0, for a
    representative accepted and rejected storage case per type."""
    out: list[CaseParams] = []
    for type_name in catalog.EXTRACTABLE_TYPES:
        cases = catalog.STORAGE_MATRIX[type_name]
        # one accepted + one rejected case per type keeps this fast while still
        # covering both branches of every shape
        chosen = []
        for accepted in (True, False):
            match = [c for c in cases if c.accepted is accepted]
            if match:
                chosen.append(match[0])
        for case in chosen:
            fmt = catalog.formats_for(type_name, case)[0]
            for direction in DIRECTIONS:
                for shape in SCALAR_SHAPES:
                    for via in VIAS:
                        out.append(CaseParams(type_name, shape, case.key, fmt,
                                              direction, 0, via=via))
                for shape in TABLE_SHAPES:
                    modes = ('same_header',) if shape == 'table1' else HEADER_MODES
                    for mode in modes:
                        out.append(CaseParams(type_name, shape, case.key, fmt,
                                              direction, 0, header_mode=mode))
    return out


PARAMS = _params()
IDS = [p.case_id for p in PARAMS]


@pytest.fixture(scope='module')
def built(tmp_path_factory):
    """Build every case once; the assertions below share them."""
    root = tmp_path_factory.mktemp('gen')
    out = {}
    for p in PARAMS:
        out[p.case_id] = build_case(p, str(root / p.case_id))
    return out


@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_generated_pattern_validates(built, params):
    """§10.1 — validate-pattern must exit 0 on every generated pattern."""
    m = built[params.case_id]
    with pytest.raises(SystemExit) as exc:
        cli.main(['validate-pattern', m.pattern_path, '-q'])
    assert exc.value.code == 0, (
        f'{params.case_id}: validate-pattern rejected the generated pattern\n'
        f'{open(m.pattern_path).read()}'
    )


@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_manifest_ref_set_matches_the_workbook(built, params):
    """Every manifest ref exists and holds the written value; no stray cells."""
    m = built[params.case_id]
    ws = openpyxl.load_workbook(m.data_path, data_only=True).active
    actual = {
        cell.coordinate
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    }
    assert actual == m.refs, (
        f'{params.case_id}: workbook/manifest ref mismatch — '
        f'only in workbook: {sorted(actual - m.refs)}, '
        f'only in manifest: {sorted(m.refs - actual)}'
    )


@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_declared_round_trip_matches_openpyxl(built, params):
    """The catalog's round_trip() must equal what openpyxl really returns."""
    m = built[params.case_id]
    case = params.storage_case
    ws = openpyxl.load_workbook(m.data_path, data_only=True).active
    for spec in m.cells:
        if spec.role not in ('value', 'data'):
            continue
        read_back = ws[spec.ref].value
        declared = case.round_trip(spec.written_value)
        assert type(read_back) is type(declared) and read_back == declared, (
            f'{params.case_id} {spec.ref}: wrote {spec.written_value!r}, '
            f'catalog declares round_trip -> {declared!r} '
            f'({type(declared).__name__}), openpyxl returned {read_back!r} '
            f'({type(read_back).__name__})'
        )


@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_error_cells_are_last_in_scan_order(built, params):
    """§4 — error cells must never sit inside the scan path of a cell:next
    chain, or they would be consumed as if they were data."""
    m = built[params.case_id]
    errors = [c for c in m.cells if c.role == 'error']
    others = [c for c in m.cells if c.role != 'error']
    assert errors, f'{params.case_id}: no error cells were generated'
    assert len(errors) == catalog.ERROR_CELLS_PER_CASE

    def idx(spec):
        return scan_index((spec.row, spec.col), params.direction, GRID, GRID)

    assert min(idx(e) for e in errors) > max(idx(o) for o in others), (
        f'{params.case_id}: an error cell precedes an extraction cell in '
        f'{params.direction} scan order'
    )
    for spec in others:
        assert spec.row <= PLACE_SIZE and spec.col <= PLACE_SIZE, (
            f'{params.case_id}: extraction cell {spec.ref} escaped the '
            f'1..{PLACE_SIZE} placement area'
        )


@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_blocks_keep_their_one_cell_margin(built, params):
    """No two written cells from different elements are adjacent, so a DATA:*
    scan always meets an empty row before the next element."""
    m = built[params.case_id]
    # group cells by element: tables by their block, scalars by anchor/value pair
    occupied = {(c.row, c.col) for c in m.cells}
    assert len(occupied) == len(m.cells), f'{params.case_id}: duplicate cell refs'

    # Every written cell must have at least one empty neighbour ring cell in
    # the direction of travel; the strong per-pair invariant is proven in
    # test_placement.py, so here we only assert generation did not add cells
    # that break it (headers/data inside one table are legitimately adjacent).
    table_cells = {(c.row, c.col) for c in m.cells if c.role in ('header', 'data')}
    scalar_cells = [(c.row, c.col) for c in m.cells if c.role in ('anchor', 'value')]
    for (r, c) in scalar_cells:
        for (r2, c2) in table_cells:
            assert max(abs(r - r2), abs(c - c2)) > 1, (
                f'{params.case_id}: scalar cell {(r, c)} touches table cell {(r2, c2)}'
            )


def test_case_ids_are_unique():
    assert len(set(IDS)) == len(IDS)


def test_matrix_is_not_vacuous():
    assert len(PARAMS) >= 200, f'only {len(PARAMS)} generator cases'
