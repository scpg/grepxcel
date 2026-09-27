"""The canonical per-cell type classifier — single source of truth for
"what is this cell, really."

Before this module, grepxcel had two independent, hand-rolled type
guessers that had quietly drifted apart: `utils.infer_cell_type()`
(column-wise heuristic, no number-format awareness) and
`wizard_core._infer_cell_type()` (single-cell, format-aware, used by the
wizard's live inspection UI). Both are now thin wrappers over
`classify_value()`/`classify_cell()` here, translated back to each
function's own historical vocabulary so existing callers see no change in
behaviour — verified against their existing test suites
(tests/unit/test_property_based.py, tests/unit/test_suggester.py).

New code (the `profile` subcommand, the type-based test-generation
framework) should call `classify_cell()`/`classify_value()` directly and
use the full vocabulary below, not the narrower legacy buckets.

Layers (see docs/pattern-file.md and the "Column D Decoder" design
discussion for the full taxonomy):

  Layer 1 — OOXML storage type (`t` attribute): n / s / str / b / e / inlineStr.
            Excel has no separate date/time/currency storage type — it's
            all `n` (number) plus a display format.
  Layer 2 — semantic subtype, built on Layer 1 via the cell's number
            format: string, integer, number, currency, percentage, date,
            time, datetime, duration, boolean, url, error, empty.
  Layer 3 — the Python type openpyxl hands back (`type(value).__name__`).
  Layer 4 — flags for the "beyond classic" cases: error, formula,
            text_forced_numeric, blank, merged_member, rich_value.
            rich_value (IMAGE()/linked-data-type cells) can't be detected
            from a plain cell value alone — it needs engine.py's
            richData-chain walk (`scan_image_cells()`), which only
            `profile.py` currently wires in as a post-classification
            override (semantic_type -> 'image', 'error' flag dropped in
            favour of 'rich_value'). classify_value()/classify_cell()
            themselves never see it.
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field

# ── Layer 1 ──────────────────────────────────────────────────────────────
STORAGE_NUMBER = 'n'
STORAGE_SHARED_STRING = 's'
STORAGE_INLINE_STRING = 'inlineStr'
STORAGE_FORMULA_STRING = 'str'
STORAGE_BOOLEAN = 'b'
STORAGE_ERROR = 'e'

# ── Layer 2 ──────────────────────────────────────────────────────────────
SEMANTIC_TYPES = frozenset({
    'string', 'integer', 'number', 'currency', 'percentage', 'date',
    'time', 'datetime', 'duration', 'boolean', 'url', 'error', 'empty',
    'image',
})

_BOOL_STRINGS = frozenset({'TRUE', 'FALSE', 'YES', 'NO'})
# The 9 current Excel formula errors (confirmed against ECMA-376 / current
# Excel documentation — see the "Column D Decoder" research this session).
_EXCEL_ERRORS = frozenset({
    '#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!',
    '#SPILL!', '#CALC!',
})
_DATE_FORMAT_HINTS = ('yyyy', 'yy/', '/yy', 'dd', 'd-mmm', 'd/m', 'm/d', 'mmm')
_CURRENCY_SYMBOLS = ('$', '€', '£', '¥', '₹', '₩')


@dataclass
class CellProfile:
    """The full classification of one cell, across all four layers."""
    storage_type: str
    semantic_type: str
    python_type: str
    flags: list[str] = field(default_factory=list)
    raw_repr: str = ''
    number_format: str = 'General'


def _is_text_forced_numeric(text: str) -> bool:
    """True if a *string* cell's text is itself a valid number — the
    classic 'leading zeros preserved as text' (or any Text-formatted
    numeric-looking cell) gotcha."""
    s = text.strip()
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def classify_value(value, number_format: str | None = None,
                   is_formula: bool = False) -> CellProfile:
    """Classify an already-read Python value (an openpyxl cell's `.value`)
    plus its number format string. Pure function — no cell/worksheet
    object required, so it works equally for a live cell or a bare value
    pulled from anywhere (CSV-inferred, a test fixture, etc.)."""
    fmt_raw = number_format or 'General'
    fmt = fmt_raw.lower()
    flags: list[str] = ['formula'] if is_formula else []

    if value is None:
        return CellProfile(STORAGE_NUMBER, 'empty', 'NoneType', flags + ['blank'], '', fmt_raw)

    if isinstance(value, str):
        s = value.strip()
        if s.upper() in _EXCEL_ERRORS:
            return CellProfile(STORAGE_ERROR, 'error', 'str', flags + ['error'], value, fmt_raw)
        if s.upper() in _BOOL_STRINGS:
            return CellProfile(STORAGE_SHARED_STRING, 'boolean', 'str', flags, value, fmt_raw)
        if '://' in s or s.lower().startswith('www.'):
            return CellProfile(STORAGE_SHARED_STRING, 'url', 'str', flags, value, fmt_raw)
        if _is_text_forced_numeric(s):
            flags = flags + ['text_forced_numeric']
        storage = STORAGE_FORMULA_STRING if is_formula else STORAGE_SHARED_STRING
        return CellProfile(storage, 'string', 'str', flags, value, fmt_raw)

    if isinstance(value, bool):
        return CellProfile(STORAGE_BOOLEAN, 'boolean', 'bool', flags, str(value), fmt_raw)

    if isinstance(value, datetime.timedelta):
        return CellProfile(STORAGE_NUMBER, 'duration', 'timedelta', flags, str(value), fmt_raw)

    if isinstance(value, datetime.time):
        return CellProfile(STORAGE_NUMBER, 'time', 'time', flags, value.isoformat(), fmt_raw)

    if isinstance(value, datetime.datetime):
        return CellProfile(STORAGE_NUMBER, 'datetime', 'datetime', flags, value.isoformat(), fmt_raw)

    if isinstance(value, datetime.date):
        return CellProfile(STORAGE_NUMBER, 'date', 'date', flags, value.isoformat(), fmt_raw)

    if isinstance(value, (int, float)):
        # Layer 2's whole point: Excel has no separate date/time/currency
        # storage type. It's all floats — the number FORMAT decides what
        # the value means. A caller with no format info (fmt == 'general')
        # falls through to the plain integer/number split below.
        if any(p in fmt for p in _DATE_FORMAT_HINTS):
            semantic = 'date'
        elif '%' in fmt:
            semantic = 'percentage'
        elif any(c in fmt for c in _CURRENCY_SYMBOLS):
            semantic = 'currency'
        elif isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
            semantic = 'integer'
        else:
            semantic = 'number'
        return CellProfile(STORAGE_NUMBER, semantic, type(value).__name__, flags, str(value), fmt_raw)

    # Anything else (shouldn't happen for a real openpyxl cell) — treat as
    # opaque text rather than raising.
    return CellProfile(STORAGE_SHARED_STRING, 'string', type(value).__name__, flags, str(value), fmt_raw)


def classify_cell(cell) -> CellProfile:
    """Classify a live openpyxl `Cell` (anything with `.value` and
    `.number_format`)."""
    value = cell.value
    is_formula = isinstance(value, str) and value.startswith('=')
    return classify_value(value, getattr(cell, 'number_format', None), is_formula=is_formula)
