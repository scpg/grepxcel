"""Unit tests for `tests/oracle/placement.py` — the placement guarantees that
the rest of the oracle type-matrix suite relies on (spec section 9).

What is asserted here:
  * blocks never overlap;
  * the one-cell margin holds, plus the two derived properties that make the
    spec's "one fully empty row and one fully empty column between blocks"
    wording true for any pair of blocks;
  * `sort_blocks` really orders by the engine's scan index, per direction;
  * `value_pos_for` is the scan-order successor of the anchor, cross-checked
    against `position_before`;
  * m=10 (the worst count the generator asks for) is placeable for seeds 0..19,
    both for scalars in both directions and for 6x3 mini-tables.
"""
from __future__ import annotations

import pytest

from tests.oracle.placement import (
    GRID,
    Block,
    OraclePlacer,
    sort_blocks,
    value_pos_for,
)
from tests.unit.test_random_type_fixtures import position_before, scan_index

pytestmark = pytest.mark.oracle

DIRECTIONS = ['LR', 'TD']
CAPACITY_SEEDS = list(range(20))


def _inflate(block: Block) -> set[tuple[int, int]]:
    """Block cells grown by 1 in all four directions, unclipped, so the test
    checks the margin independently of placement.py's own clipping."""
    return {
        (r, c)
        for r in range(block.top - 1, block.bottom + 2)
        for c in range(block.left - 1, block.right + 2)
    }


def _ranges_separated(lo_a: int, hi_a: int, lo_b: int, hi_b: int) -> bool:
    """True if [lo_a, hi_a] and [lo_b, hi_b] have at least one index between
    them (they neither overlap nor merely touch)."""
    return hi_a + 1 < lo_b or hi_b + 1 < lo_a


def _ranges_overlap_or_touch(lo_a: int, hi_a: int, lo_b: int, hi_b: int) -> bool:
    return not _ranges_separated(lo_a, hi_a, lo_b, hi_b)


def _mixed_placement(seed: int, direction: str) -> list[Block]:
    """A deliberately mixed set: anchored scalars, bare value cells and
    mini-tables of every shape the generator can ask for, all from one
    placer so the invariant is exercised across kinds."""
    placer = OraclePlacer(seed)
    blocks = [
        placer.place_scalar(direction, True),
        placer.place_scalar(direction, False),
        placer.place_scalar(direction, True),
        placer.place_table(1, 1),
        placer.place_table(5, 3),
        placer.place_table(3, 2),
        placer.place_scalar(direction, False),
        placer.place_table(2, 3),
    ]
    assert placer.blocks == blocks
    return blocks


# ── 1. no overlap ──────────────────────────────────────────────────────────

@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_no_cell_claimed_twice(seed, direction):
    blocks = _mixed_placement(seed, direction)
    claimed: set[tuple[int, int]] = set()
    for block in blocks:
        assert not (block.cells & claimed), f'seed={seed} {direction}: overlap at {block}'
        claimed |= block.cells
    assert len(claimed) == sum(b.height * b.width for b in blocks)


@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_every_block_inside_the_grid(seed, direction):
    for block in _mixed_placement(seed, direction):
        assert 1 <= block.top <= block.bottom <= GRID, f'seed={seed}: {block}'
        assert 1 <= block.left <= block.right <= GRID, f'seed={seed}: {block}'


# ── 2. the margin invariant and its two derived properties ─────────────────

@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_margin_halo_never_touches_another_block(seed, direction):
    blocks = _mixed_placement(seed, direction)
    for i, a in enumerate(blocks):
        for j, b in enumerate(blocks):
            if i == j:
                continue
            assert not (_inflate(a) & b.cells), (
                f'seed={seed} {direction}: block {i} halo touches block {j} '
                f'({a} vs {b})'
            )


@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_row_overlap_implies_an_empty_column_between(seed, direction):
    """If two blocks share (or touch on) rows, an empty column separates them —
    so a horizontal neighbour can never be read as part of the other block."""
    blocks = _mixed_placement(seed, direction)
    for i, a in enumerate(blocks):
        for b in blocks[i + 1:]:
            if _ranges_overlap_or_touch(a.top, a.bottom, b.top, b.bottom):
                assert _ranges_separated(a.left, a.right, b.left, b.right), (
                    f'seed={seed} {direction}: {a} and {b} share rows but are not '
                    f'separated by an empty column'
                )


@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_column_overlap_implies_an_empty_row_between(seed, direction):
    """If two blocks share (or touch on) columns, an empty row separates them —
    this is what stops a `DATA:*` scan from running into the block below."""
    blocks = _mixed_placement(seed, direction)
    for i, a in enumerate(blocks):
        for b in blocks[i + 1:]:
            if _ranges_overlap_or_touch(a.left, a.right, b.left, b.right):
                assert _ranges_separated(a.top, a.bottom, b.top, b.bottom), (
                    f'seed={seed} {direction}: {a} and {b} share columns but are not '
                    f'separated by an empty row'
                )


def test_edge_placement_is_legal():
    """A block flush against the grid edge is accepted — the halo is clipped,
    not treated as out of bounds."""
    for corner in [
        Block(1, 1, 1, 2, 'scalar'),
        Block(1, GRID - 1, 1, 2, 'scalar'),
        Block(GRID - 5, GRID - 2, 6, 3, 'table'),
        Block(GRID, GRID, 1, 1, 'scalar'),
    ]:
        placer = OraclePlacer(0)
        assert placer._fits(corner), f'{corner} should be placeable'


def test_block_outside_the_grid_is_rejected():
    placer = OraclePlacer(0)
    assert not placer._fits(Block(0, 1, 1, 1, 'scalar'))
    assert not placer._fits(Block(1, 0, 1, 1, 'scalar'))
    assert not placer._fits(Block(GRID, 1, 2, 1, 'scalar'))
    assert not placer._fits(Block(1, GRID, 1, 2, 'scalar'))


def test_diagonally_adjacent_block_is_rejected():
    """Chebyshev distance 1 includes diagonals: a block whose corner merely
    kisses another's corner must be refused."""
    placer = OraclePlacer(0)
    placer._commit(Block(10, 10, 1, 1, 'scalar'))
    assert not placer._fits(Block(11, 11, 1, 1, 'scalar'))  # diagonal touch
    assert not placer._fits(Block(11, 10, 1, 1, 'scalar'))  # directly below
    assert not placer._fits(Block(10, 11, 1, 1, 'scalar'))  # directly right
    assert placer._fits(Block(12, 12, 1, 1, 'scalar'))      # one gap each way


def test_oversized_block_raises_with_the_seed():
    placer = OraclePlacer(7)
    with pytest.raises(RuntimeError, match='seed=7'):
        placer.place_table(GRID, GRID)


def test_exhausted_grid_raises_with_the_seed():
    """A tiny grid genuinely cannot hold two margined blocks — the failure names
    the seed and what failed to place."""
    placer = OraclePlacer(3, size=3)
    placer.place_table(1, 3)  # fills rows 1-2 across the full width
    with pytest.raises(RuntimeError) as excinfo:
        placer.place_table(1, 3)
    assert 'seed=3' in str(excinfo.value)
    assert 'mini-table' in str(excinfo.value)


# ── 3. scan-order sort ─────────────────────────────────────────────────────

@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_sort_blocks_is_ascending_scan_index(seed, direction):
    blocks = _mixed_placement(seed, direction)
    ordered = sort_blocks(blocks, direction)
    indices = [scan_index(b.anchor, direction, GRID, GRID) for b in ordered]
    assert indices == sorted(indices), f'seed={seed} {direction}: {indices}'
    for a, b in zip(ordered, ordered[1:]):
        assert position_before(a.anchor, b.anchor, direction, GRID, GRID)


@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_sort_blocks_keeps_every_block(seed, direction):
    blocks = _mixed_placement(seed, direction)
    ordered = sort_blocks(blocks, direction)
    assert len(ordered) == len(blocks)
    assert {b.anchor for b in ordered} == {b.anchor for b in blocks}


def test_sort_blocks_honours_direction():
    """A hand-built case where LR and TD genuinely disagree, proving the
    direction argument is used and not incidentally irrelevant."""
    a = Block(1, 5, 1, 1, 'scalar')   # row 1, col 5
    b = Block(5, 1, 1, 1, 'scalar')   # row 5, col 1
    assert [blk.anchor for blk in sort_blocks([b, a], 'LR')] == [(1, 5), (5, 1)]
    assert [blk.anchor for blk in sort_blocks([a, b], 'TD')] == [(5, 1), (1, 5)]


# ── 4. value_pos_for ───────────────────────────────────────────────────────

@pytest.mark.parametrize('anchor', [(1, 1), (1, 5), (7, 3), (GRID - 1, GRID - 1)])
def test_value_pos_for_is_the_scan_order_successor(anchor):
    for direction in DIRECTIONS:
        value = value_pos_for(anchor, direction)
        expected = (anchor[0], anchor[1] + 1) if direction == 'LR' else (anchor[0] + 1, anchor[1])
        assert value == expected
        assert position_before(anchor, value, direction, GRID, GRID)
        assert (
            scan_index(value, direction, GRID, GRID)
            == scan_index(anchor, direction, GRID, GRID) + 1
        ), f'{direction}: {anchor} -> {value} is not the immediate successor'


def test_value_pos_for_rejects_unknown_direction():
    with pytest.raises(ValueError, match='unknown direction'):
        value_pos_for((1, 1), 'XX')


@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
def test_anchored_scalar_block_holds_anchor_and_its_successor(seed, direction):
    placer = OraclePlacer(seed)
    for _ in range(5):
        block = placer.place_scalar(direction, True)
        value = value_pos_for(block.anchor, direction)
        assert block.cells == {block.anchor, value}
        assert position_before(block.anchor, value, direction, GRID, GRID)
        if direction == 'LR':
            assert (block.height, block.width) == (1, 2)
        else:
            assert (block.height, block.width) == (2, 1)


@pytest.mark.parametrize('direction', DIRECTIONS)
def test_unanchored_scalar_is_a_single_cell(direction):
    block = OraclePlacer(1).place_scalar(direction, False)
    assert (block.height, block.width) == (1, 1)
    assert block.cells == {block.anchor}


def test_place_scalar_rejects_unknown_direction():
    with pytest.raises(ValueError, match='unknown direction'):
        OraclePlacer(0).place_scalar('sideways', True)


@pytest.mark.parametrize('data_rows', [1, 2, 3, 4, 5])
@pytest.mark.parametrize('cols', [1, 2, 3])
def test_table_block_geometry(data_rows, cols):
    block = OraclePlacer(data_rows * 10 + cols).place_table(data_rows, cols)
    assert block.kind == 'table'
    assert (block.height, block.width) == (1 + data_rows, cols)
    assert len(block.cells) == (1 + data_rows) * cols
    assert block.anchor == (block.top, block.left)


# ── 5. m=10 capacity ───────────────────────────────────────────────────────

@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
@pytest.mark.parametrize('direction', DIRECTIONS)
@pytest.mark.parametrize('with_anchor', [True, False])
def test_ten_scalars_place_successfully(seed, direction, with_anchor):
    placer = OraclePlacer(seed)
    blocks = [placer.place_scalar(direction, with_anchor) for _ in range(10)]
    assert len(blocks) == 10
    assert len({b.anchor for b in blocks}) == 10


@pytest.mark.parametrize('seed', CAPACITY_SEEDS)
def test_ten_worst_case_tables_place_successfully(seed):
    """data_rows=5, cols=3 — the largest mini-table the generator asks for."""
    placer = OraclePlacer(seed)
    blocks = [placer.place_table(5, 3) for _ in range(10)]
    assert len(blocks) == 10
    claimed: set[tuple[int, int]] = set()
    for block in blocks:
        assert not (_inflate(block) & claimed)
        claimed |= block.cells
    assert len(claimed) == 10 * 6 * 3
