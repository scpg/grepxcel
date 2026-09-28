"""The manifest — the single oracle for one generated case.

A "case" is one generated (pattern.csv, data.xlsx) pair. The manifest records
*exactly* what the generator wrote, plus two INDEPENDENT sets of expectations:

  * ``extract_accepted`` / ``extract_reason`` — what ``utils.validate_type()``
    does with the cell when a ``var:`` field of the case's type targets it.
  * ``profile_storage`` / ``profile_semantic`` / ``profile_flags`` — what
    ``cell_taxonomy.classify_value()`` says the cell is.

These two deliberately disagree for several (type, storage) pairs and that is
not a bug to be reconciled — see the catalog module's notes A and B. A cell
written as ``datetime.date`` is accepted by ``var: date`` yet classified
``n/datetime``; ``0``/``1`` are accepted as boolean yet classified integer.
Two code paths, two truths, both recorded, neither "fixed" to match the other.

The manifest is written next to the generated files as ``manifest.json``
(``default=str``) so a failing case can be inspected by hand.
"""
from __future__ import annotations

import datetime
import json
from dataclasses import asdict, dataclass, field


def normalise(value):
    """The single normalisation both sides of every value comparison go through.

    Mirrors ``tests/integration/test_schema_validation._JSONEncoder`` so the
    oracle compares what the CLI would actually emit as JSON: ``date`` /
    ``datetime`` / ``time`` become ISO strings, ``timedelta`` becomes ``str()``,
    everything else is untouched. Applied recursively to lists and dicts so a
    whole extract result can be normalised in one call (§10.4).
    """
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, datetime.timedelta):
        return str(value)
    if isinstance(value, dict):
        return {k: normalise(v) for k, v in value.items()}
    if isinstance(value, list):
        return [normalise(v) for v in value]
    return value


@dataclass
class CellSpec:
    """One cell the generator wrote, with its per-oracle expectations."""
    ref: str                          # 'B7'
    row: int
    col: int
    field: str | None                 # var: field name; None for anchors/headers/errors
    role: str                         # 'anchor' | 'value' | 'header' | 'data' | 'error'
    written_value: object             # exactly what openpyxl was given
    number_format: str | None

    # ── extract expectations (validate_type truth) ────────────────────────
    extract_accepted: bool | None = None   # None when the cell is not a var: target
    extract_reason: str | None = None      # substring of the warning, when rejected

    # ── profile expectations (classify_value truth) ───────────────────────
    profile_storage: str = ''
    profile_semantic: str = ''
    profile_flags: frozenset[str] = frozenset()   # MUST be present (subset check)


@dataclass
class TableSpec:
    """One mini-table definition and the instances it must produce."""
    group_key: str                    # output key, e.g. 'table_0', 't0', 'items'
    header_ref: str
    columns: list[str]                # var: field names, in order
    data_rows: list[list[object]]
    expected_instances: int           # 1 for distinct-header mode; m for same-header


@dataclass
class Manifest:
    """Everything one generated case asserts about itself."""
    case_id: str                      # '{type}-{shape}-{storage}-{fmt}-{dir}-seed{n}'
    direction: str                    # 'LR' | 'TD'
    pattern_path: str
    data_path: str
    cells: list[CellSpec] = field(default_factory=list)
    tables: list[TableSpec] = field(default_factory=list)
    expected_extract: dict = field(default_factory=dict)

    # ── hand-inspection helpers ───────────────────────────────────────────
    def to_json(self, indent: int = 2) -> str:
        """Serialise for hand inspection. ``default=str`` because cells carry
        real ``date``/``time``/``timedelta`` objects and ``frozenset`` flags —
        this is a debugging artefact, not a round-trip format."""
        return json.dumps(asdict(self), indent=indent, default=str, sort_keys=False)

    def write(self, path: str) -> str:
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write(self.to_json())
        return path

    @classmethod
    def from_json(cls, text: str) -> 'Manifest':
        """Rebuild a Manifest from :meth:`to_json` output.

        Lossy by construction: ``to_json`` stringifies values openpyxl was
        given (``datetime.date(2024, 1, 15)`` becomes ``'2024-01-15'``), so a
        round-tripped manifest is for reading, never for re-asserting. Flags
        come back as a frozenset; everything else stays as JSON gave it.
        """
        raw = json.loads(text)
        cells = [
            CellSpec(**{**c, 'profile_flags': frozenset(c.get('profile_flags') or ())})
            for c in raw.get('cells', [])
        ]
        tables = [TableSpec(**t) for t in raw.get('tables', [])]
        return cls(
            case_id=raw['case_id'],
            direction=raw['direction'],
            pattern_path=raw['pattern_path'],
            data_path=raw['data_path'],
            cells=cells,
            tables=tables,
            expected_extract=raw.get('expected_extract', {}),
        )

    @classmethod
    def read(cls, path: str) -> 'Manifest':
        with open(path, encoding='utf-8') as fh:
            return cls.from_json(fh.read())

    # ── convenience accessors used by the oracles ─────────────────────────
    @property
    def refs(self) -> frozenset[str]:
        """Every cell ref the generator wrote — the profile oracle asserts the
        profiled ref set equals exactly this."""
        return frozenset(c.ref for c in self.cells)

    def value_cells(self) -> list[CellSpec]:
        """Cells a ``var:`` field targets (scalars and table data)."""
        return [c for c in self.cells if c.field is not None and c.role in ('value', 'data')]

    def rejected_cells(self) -> list[CellSpec]:
        return [c for c in self.value_cells() if c.extract_accepted is False]

    @property
    def has_rejection(self) -> bool:
        """Drives the ``--strict`` exit-code expectation (2 vs 0)."""
        return bool(self.rejected_cells())
