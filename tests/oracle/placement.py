"""Shapes and placement for the oracle type-matrix suite (spec section 9).

Random, non-overlapping placement of *blocks* on a 30x30 grid, where a block is
either a scalar element (an anchor cell plus its value cell, or a lone value
cell) or a mini-table rectangle (one header row plus `data_rows` data rows).

Two invariants make randomly-located elements safe for a forward-scanning
engine:

1. **Margin.** No cell of a new block may lie within Chebyshev distance 1 of an
   already-occupied cell — i.e. the new block's rectangle, inflated by 1 in all
   four directions, must contain no previously-occupied cell. This is the
   spec's "at least one fully empty row and one fully empty column between
   blocks" requirement in the form that is both satisfiable at m=10 and
   sufficient for the engine: `engine._row_is_end_of_data` only inspects a
   table's own column span, so two blocks side by side with one empty column
   between them do not bleed into each other.
2. **Scan order.** `sort_blocks` orders blocks by the engine's own scan index
   (mirrored by `scan_index`, imported from the existing random-type-fixture
   framework rather than reinvented), so pattern instructions can be emitted in
   forward scan order.

Blocks may touch the grid edge; the halo is simply clipped by the boundary.
"""
from __future__ import annotations

import random
from dataclasses import dataclass

from tests.unit.test_random_type_fixtures import position_before, scan_index

__all__ = [
    'GRID',
    'MAX_DRAWS',
    'Block',
    'OraclePlacer',
    'value_pos_for',
    'sort_blocks',
]

GRID = 30  # 30 rows x 30 cols
MAX_DRAWS = 2000  # random draws before giving up on a single block


@dataclass(frozen=True)
class Block:
    """A rectangle of cells occupied by one generated element."""

    top: int  # 1-based row of the top-left cell
    left: int  # 1-based col of the top-left cell
    height: int
    width: int
    kind: str  # 'scalar' or 'table'

    @property
    def anchor(self) -> tuple[int, int]:
        """The anchor / header cell — always the block's top-left cell."""
        return (self.top, self.left)

    @property
    def bottom(self) -> int:
        return self.top + self.height - 1

    @property
    def right(self) -> int:
        return self.left + self.width - 1

    @property
    def cells(self) -> set[tuple[int, int]]:
        """Every (row, col) the block occupies."""
        return {
            (r, c)
            for r in range(self.top, self.top + self.height)
            for c in range(self.left, self.left + self.width)
        }

    def halo(self, size: int = GRID) -> set[tuple[int, int]]:
        """`cells` inflated by 1 in all four directions, clipped to the grid."""
        return {
            (r, c)
            for r in range(max(1, self.top - 1), min(size, self.bottom + 1) + 1)
            for c in range(max(1, self.left - 1), min(size, self.right + 1) + 1)
        }


def value_pos_for(anchor: tuple[int, int], direction: str) -> tuple[int, int]:
    """The scan-order successor of *anchor*: (r, c+1) for LR, (r+1, c) for TD."""
    row, col = anchor
    if direction == 'LR':
        return (row, col + 1)
    if direction == 'TD':
        return (row + 1, col)
    raise ValueError(f'unknown direction {direction!r}')


def sort_blocks(blocks: list[Block], direction: str, size: int = GRID) -> list[Block]:
    """Blocks sorted ascending by the scan index of their anchor cell."""
    return sorted(blocks, key=lambda b: scan_index(b.anchor, direction, size, size))


class OraclePlacer:
    """Seeded, non-overlapping block placement with a one-cell margin."""

    def __init__(self, seed: int, size: int = GRID):
        self.seed = seed
        self.size = size
        self.rng = random.Random(seed)
        self._occupied: set[tuple[int, int]] = set()
        self.blocks: list[Block] = []

    # ── invariant ──────────────────────────────────────────────────────────

    def _fits(self, block: Block) -> bool:
        """True if *block* lies fully inside the grid and its inflated
        rectangle contains no already-occupied cell."""
        if block.top < 1 or block.left < 1:
            return False
        if block.bottom > self.size or block.right > self.size:
            return False
        return not (block.halo(self.size) & self._occupied)

    def _commit(self, block: Block) -> Block:
        self._occupied |= block.cells
        self.blocks.append(block)
        return block

    def _place(self, height: int, width: int, kind: str, what: str) -> Block:
        if height > self.size or width > self.size:
            raise RuntimeError(
                f'seed={self.seed}: cannot place {what} '
                f'({height}x{width}) — larger than the {self.size}x{self.size} grid'
            )
        max_top = self.size - height + 1
        max_left = self.size - width + 1
        for _ in range(MAX_DRAWS):
            candidate = Block(
                top=self.rng.randint(1, max_top),
                left=self.rng.randint(1, max_left),
                height=height,
                width=width,
                kind=kind,
            )
            if self._fits(candidate):
                return self._commit(candidate)
        raise RuntimeError(
            f'seed={self.seed}: could not place {what} ({height}x{width}) '
            f'after {MAX_DRAWS} draws — {len(self.blocks)} block(s) already placed '
            f'on a {self.size}x{self.size} grid'
        )

    # ── public API ─────────────────────────────────────────────────────────

    def place_scalar(self, direction: str, with_anchor: bool) -> Block:
        """A scalar element.

        with_anchor=True  (pattern via='next'): anchor cell + its value cell,
            where the value cell is the scan-order successor of the anchor —
            'LR' -> (r, c+1) so the block is 1x2; 'TD' -> (r+1, c) so the block
            is 2x1.
        with_anchor=False (pattern via='abs'): a single value cell, 1x1.
        """
        if direction not in ('LR', 'TD'):
            raise ValueError(f'unknown direction {direction!r}')
        if not with_anchor:
            return self._place(1, 1, 'scalar', 'scalar value cell')
        if direction == 'LR':
            block = self._place(1, 2, 'scalar', 'LR scalar anchor+value pair')
        else:
            block = self._place(2, 1, 'scalar', 'TD scalar anchor+value pair')
        # The anchor must precede its value cell in the engine's scan order,
        # otherwise `cell:next` would never reach the value.
        value = value_pos_for(block.anchor, direction)
        assert value in block.cells
        assert position_before(block.anchor, value, direction, self.size, self.size)
        return block

    def place_table(self, data_rows: int, cols: int) -> Block:
        """A mini-table: 1 header row + `data_rows` data rows, `cols` wide."""
        if data_rows < 1 or cols < 1:
            raise ValueError(f'invalid table shape data_rows={data_rows} cols={cols}')
        return self._place(
            1 + data_rows, cols, 'table', f'mini-table ({data_rows} data rows x {cols} cols)'
        )
