"""The storage/format/value catalog — spec §5, §6, §7 encoded as data.

Every ``StorageCase`` below is one (type, storage) pair with a *declared*
outcome for two independent code paths:

  * ``accepted`` / ``reason``  — ``grepxcel.utils.validate_type()``
  * ``profile_storage`` / ``profile_semantic`` / ``profile_flags``
    — ``grepxcel.cell_taxonomy.classify_value()``

Both expectations describe the value **as openpyxl reads it back**, not as it
was written: the engine and the profiler both read the saved .xlsx, and
openpyxl's own round-trip re-types some values (a written ``datetime.date``
comes back a ``datetime.datetime``; a number formatted ``[h]:mm`` comes back a
``timedelta``). ``round_trip`` on each case declares that transform so the
generator can state the expected extract value without guessing, and
``tests/oracle/test_generator.py`` asserts the declared transform against a
real write/read cycle rather than trusting it.

Verified against the running code by ``tmp.local/verify_catalog.py``, which
writes every (case, format, value) triple to a real workbook, reads it back and
compares ``validate_type()`` and ``classify_value()`` to the rows below.

## Corrections made to the spec's §5 table (the code is the authority)

1. **``number`` / ``n`` / float → the spec's single row became two cases.**
   §5 says "number | n | float | accepted | n/number". That is only true for
   *non-integral* floats: ``classify_value`` checks ``value.is_integer()`` and
   calls ``0.0``, ``-1e15``, ``1e15`` **integer**, not number — and those are
   exactly the §7 boundary values for this type. Split into ``n-float``
   (non-integral only → ``n/number``) and ``n-int`` (ints *and* integral floats
   → ``n/integer``, which is also the spec's own "number | n | int" row).

2. **``boolean`` / ``s`` → the spec's single row became two cases.** §5 lists
   ``"TRUE"``/``"yes"``/``"0"`` on one row but its own "profile expects" column
   gives two different answers for them. ``classify_value`` returns
   ``s/boolean`` for the four words in ``_BOOL_STRINGS`` and
   ``s/string`` + ``text_forced_numeric`` for ``"0"``/``"1"``. One row cannot
   carry two expectations: split into ``s-boolword`` and ``s-boolnum``.

3. **``string`` accepted pool excludes ``'x' * 1001``.** ``_BOUNDARY_VALUES``
   in the existing suite carries 999/1000/1001 to straddle
   ``utils._MAX_REGEX_INPUT_LEN``. 1001 chars makes ``_safe_match`` return
   False (input-length guard), i.e. **rejected** — so it cannot sit in a row
   §5 declares unconditionally accepted. §7 asks for "a string of exactly
   ``_MAX_REGEX_INPUT_LEN`` chars", which is 1000 and is accepted; the 1001
   case belongs to a rejected row that §5 does not define, so it is out of
   scope here (``tests/unit/test_random_type_fixtures.py`` still covers it).

4. **``percentage`` pool gains ``1.5``** — §5 asks for "float in [0,1] (and one
   >1)" but ``_BOUNDARY_VALUES`` only has ``0.0``/``1.0``.

5. **``datetime`` pools are new.** §4 requires the type but the existing
   ``TYPE_CASES``/``_BOUNDARY_VALUES`` have no ``datetime`` entry.

6. **Integral floats arrive as ``int``.** Not a §5 row as such, but the fact
   that forced the ``round_trip`` transform below: openpyxl writes ``42.0`` as
   ``42`` and reads it back an ``int``. See ``StorageCase.round_trip`` for the
   coverage consequence — §5's "integer | n | float with .is_integer()" row
   cannot reach ``validate_type``'s float branch through a real .xlsx.

## Interpretation notes (where two spec sentences had to be reconciled)

* §6 says number formats are "applied only to ``n``-storage cases", yet its
  ``string`` row lists ``General``/``@`` with a note about ``@`` and
  ``text_forced_numeric`` — and ``string``'s only storage is ``s``. Formats are
  therefore iterated for ``n``-storage cases **and** for ``string``'s
  ``s``-storage case; every other ``s``/``b`` case is written with no explicit
  format. Harmless either way: ``classify_value`` ignores the number format for
  ``str`` and ``bool`` values.

* Several §5 rows pin a format that is not in their type's §6 list (``time``
  written as a ``timedelta`` with ``[h]:mm``; ``duration`` written as a
  ``datetime.time`` with ``HH:MM``). Those cases carry a ``formats`` override,
  because for a numeric cell the format is what decides openpyxl's read-back
  Python type — it is load-bearing, not decoration.

* ``PROFILE_ONLY_TYPES`` (``error``) has no §5 row and so never enters the
  storage matrix. Per §4 those cells are "checked by the lint and profile
  oracles only; never targeted by a ``var:`` field": the generator appends
  ``ERROR_CELLS_PER_CASE`` of them to every case, placed strictly last in scan
  order so a forward-scanning ``cell:next`` pattern has already run out of
  instructions before reaching them.
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass

from grepxcel.cell_taxonomy import _EXCEL_ERRORS
from grepxcel.utils import _MAX_REGEX_INPUT_LEN

# Import (never copy) the existing suite's curated values — §7.
from tests.unit.test_random_type_fixtures import _BOUNDARY_VALUES, TYPE_CASES

# ── §4 types in scope ─────────────────────────────────────────────────────
EXTRACTABLE_TYPES: list[str] = [
    'string', 'integer', 'number', 'currency', 'percentage', 'date',
    'datetime', 'time', 'duration', 'boolean', 'url',
]
PROFILE_ONLY_TYPES: list[str] = ['error']

#: The 9 Excel error codes, written as literal strings (openpyxl cannot
#: produce a real cached formula error). Sorted for determinism.
ERROR_CODES: tuple[str, ...] = tuple(sorted(_EXCEL_ERRORS))

#: How many error cells the generator appends to each case (see docstring).
ERROR_CELLS_PER_CASE = 2

#: Label texts the generator reserves; §7 forbids any pool value equal to one.
RESERVED_LABEL_PREFIXES = ('ANCHOR', 'HEAD', 'COLHEAD')

# ── §6 number formats per type (n-storage; plus string, see docstring) ────
FORMATS: dict[str, tuple[str | None, ...]] = {
    'string':     ('General', '@'),
    'integer':    ('General', '0', '#,##0'),
    'number':     ('General', '0.00', '#,##0.00'),
    'currency':   ('$#,##0.00', '€#,##0.00', '£#,##0.00'),
    'percentage': ('0%', '0.00%'),
    'date':       ('yyyy-mm-dd', 'dd/mm/yyyy', 'd-mmm-yy'),
    'datetime':   ('yyyy-mm-dd hh:mm', 'yyyy-mm-dd hh:mm:ss'),
    'time':       ('HH:MM', 'HH:MM:SS'),
    'duration':   ('[h]:mm', '[h]:mm:ss', '[mm]:ss'),
    'boolean':    ('General',),
    'url':        ('General',),
}

# ── §7 value pools ────────────────────────────────────────────────────────
# Built from the existing suite's TYPE_CASES / _BOUNDARY_VALUES (imported),
# with the additions and one removal recorded in the docstring.

def _pool(type_name: str, *extra, drop=()) -> tuple:
    """TYPE_CASES' ordinary value + that type's boundary values + extras."""
    values = [TYPE_CASES[type_name]['valid']] if type_name in TYPE_CASES else []
    values += list(_BOUNDARY_VALUES.get(type_name, ()))
    values += list(extra)
    seen, out = set(), []
    for v in values:
        if v in drop:
            continue
        key = (type(v).__name__, v)
        if key in seen:
            continue
        seen.add(key)
        out.append(v)
    return tuple(out)


#: Correction 3: 1001 chars is rejected by the ReDoS length guard, so it cannot
#: live in an "accepted" row. Exactly _MAX_REGEX_INPUT_LEN is the §7 boundary.
_TOO_LONG_STRING = 'x' * (_MAX_REGEX_INPUT_LEN + 1)

_STRING_POOL = _pool('string', drop=(_TOO_LONG_STRING,))
_INT_POOL = _pool('integer')
_INTEGRAL_FLOAT_POOL = (42.0, 0.0, -1e15, 1e15)
_NONINTEGRAL_FLOAT_POOL = (42.5, -0.01, 999_999_999.99)
_CURRENCY_POOL = _pool('currency')
_PERCENTAGE_POOL = _pool('percentage', 1.5)          # correction 4
_DATE_POOL = _pool('date')
_DATETIME_POOL = (                                    # correction 5
    datetime.datetime(2024, 1, 15, 9, 30),
    datetime.datetime(1900, 1, 1, 0, 0),
    datetime.datetime(9999, 12, 31, 23, 59, 59),
)
_TIME_POOL = _pool('time')
_DURATION_POOL = _pool('duration')
_URL_POOL = _pool('url')


@dataclass(frozen=True)
class StorageCase:
    """One (type, storage) row of §5, with both oracles' expectations."""
    key: str                      # short id, used in case_id
    storage: str                  # 'n' | 's' | 'b' — how the generator writes it
    values: tuple                 # pool; the seed picks one
    accepted: bool                # validate_type outcome
    reason: str | None            # substring that MUST appear in the warning
    profile_storage: str          # classify_value().storage_type
    profile_semantic: str         # classify_value().semantic_type
    profile_flags: frozenset = frozenset()   # flags that MUST be present
    formats: tuple | None = None  # override FORMATS[type]; None → use it

    def round_trip(self, value):
        """The value as openpyxl reads it back after a write.

        Two values change identity, both verified by a real write/read cycle
        (``tmp.local/verify_catalog.py``, and asserted in ``test_generator.py``
        so a change in openpyxl fails loudly instead of silently):

        * ``datetime.date`` → ``datetime.datetime`` at midnight. openpyxl has
          no date-only cell type (spec note A).
        * an **integral float** → ``int``. Excel stores ``42.0`` as ``42`` and
          openpyxl's reader casts a numeric string with no ``.``/``E`` back to
          ``int``. Consequence worth stating plainly: ``validate_type``'s
          ``isinstance(value, float) and value.is_integer()`` branch is
          therefore *unreachable* through any openpyxl-written .xlsx, so §5's
          "integer | n | float with .is_integer()" row cannot exercise it here
          — the value arrives as an int. That branch is only reachable by
          calling ``validate_type()`` directly; this suite does not claim to
          cover it.
        """
        if type(value) is datetime.date:
            return datetime.datetime(value.year, value.month, value.day)
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value


def formats_for(type_name: str, case: StorageCase) -> tuple:
    """The formats to iterate for one case (§6 + the two reconciliations)."""
    if case.formats is not None:
        return case.formats
    if case.storage == 'n' or type_name == 'string':
        return FORMATS[type_name]
    return (None,)


# ── §5 the storage matrix ─────────────────────────────────────────────────
# Reason substrings are copied verbatim from grepxcel/utils.py::validate_type.

STORAGE_MATRIX: dict[str, list[StorageCase]] = {
    'string': [
        StorageCase('s-text', 's', _STRING_POOL, True, None, 's', 'string'),
    ],
    'integer': [
        StorageCase('n-int', 'n', _INT_POOL, True, None, 'n', 'integer'),
        StorageCase('n-intfloat', 'n', _INTEGRAL_FLOAT_POOL, True, None, 'n', 'integer'),
        StorageCase('n-float', 'n', _NONINTEGRAL_FLOAT_POOL, False,
                    'is not a whole number', 'n', 'number'),
        StorageCase('s-text', 's', ('42', '0', '-7'), False,
                    'is not integer type', 's', 'string',
                    frozenset({'text_forced_numeric'})),
        StorageCase('b-bool', 'b', (True, False), False,
                    'boolean is not integer', 'b', 'boolean'),
    ],
    'number': [
        # Correction 1: non-integral floats only — an integral float is
        # classified 'integer', not 'number'.
        StorageCase('n-float', 'n', _NONINTEGRAL_FLOAT_POOL, True, None, 'n', 'number'),
        StorageCase('n-int', 'n', _INT_POOL + _INTEGRAL_FLOAT_POOL, True, None,
                    'n', 'integer'),
        StorageCase('s-text', 's', ('42.5', '0', '-0.01'), False,
                    'is not numeric', 's', 'string',
                    frozenset({'text_forced_numeric'})),
        StorageCase('b-bool', 'b', (True, False), False,
                    'boolean is not number', 'b', 'boolean'),
    ],
    'currency': [
        StorageCase('n-float', 'n', _CURRENCY_POOL, True, None, 'n', 'currency'),
        StorageCase('s-text', 's', ('12.50', '0.00'), False,
                    'is not numeric', 's', 'string',
                    frozenset({'text_forced_numeric'})),
        StorageCase('b-bool', 'b', (True, False), False,
                    'boolean is not currency', 'b', 'boolean'),
    ],
    'percentage': [
        StorageCase('n-float', 'n', _PERCENTAGE_POOL, True, None, 'n', 'percentage'),
        StorageCase('s-text', 's', ('0.15', '1'), False,
                    'is not numeric', 's', 'string',
                    frozenset({'text_forced_numeric'})),
        StorageCase('b-bool', 'b', (True, False), False,
                    'boolean is not percentage', 'b', 'boolean'),
    ],
    'date': [
        # Note A: accepted as var: date, yet classified n/datetime — openpyxl
        # returns midnight of that day. Two oracles, two truths.
        StorageCase('n-date', 'n', _DATE_POOL, True, None, 'n', 'datetime'),
        StorageCase('s-text', 's', ('2024-01-15', '1900-01-01'), False,
                    'is not a date', 's', 'string'),
    ],
    'datetime': [
        StorageCase('n-datetime', 'n', _DATETIME_POOL, True, None, 'n', 'datetime'),
        StorageCase('s-text', 's', ('2024-01-15 09:30', '1900-01-01 00:00'), False,
                    'is not a datetime', 's', 'string'),
    ],
    'time': [
        StorageCase('n-time', 'n', _TIME_POOL, True, None, 'n', 'time'),
        # §5: a timedelta is accepted for 'time' too. The [h]:mm format is what
        # makes openpyxl read it back as a timedelta, hence the override.
        StorageCase('n-timedelta', 'n', _DURATION_POOL, True, None, 'n', 'duration',
                    formats=('[h]:mm',)),
        StorageCase('s-text', 's', ('09:30', '00:00'), False,
                    'is not a time', 's', 'string'),
    ],
    'duration': [
        StorageCase('n-timedelta', 'n', _DURATION_POOL, True, None, 'n', 'duration'),
        StorageCase('n-time', 'n', _TIME_POOL, True, None, 'n', 'time',
                    formats=('HH:MM',)),
        StorageCase('s-text', 's', ('2:00', '0:00'), False,
                    'is not a time', 's', 'string'),
    ],
    'boolean': [
        StorageCase('b-bool', 'b', (True, False), True, None, 'b', 'boolean'),
        # Note B: 0/1 accepted as boolean, classified integer.
        StorageCase('n-01', 'n', (0, 1), True, None, 'n', 'integer'),
        StorageCase('n-two', 'n', (2, -1), False, 'is not a boolean', 'n', 'integer'),
        # Correction 2: the four bool words classify as s/boolean …
        StorageCase('s-boolword', 's', ('TRUE', 'FALSE', 'yes', 'no'), True, None,
                    's', 'boolean'),
        # … while "0"/"1" classify as s/string + text_forced_numeric.
        StorageCase('s-boolnum', 's', ('0', '1'), True, None, 's', 'string',
                    frozenset({'text_forced_numeric'})),
        StorageCase('s-text', 's', ('banana', 'maybe'), False,
                    'is not a boolean', 's', 'string'),
    ],
    'url': [
        StorageCase('s-url', 's', _URL_POOL, True, None, 's', 'url'),
        StorageCase('s-noscheme', 's', ('example.com/x', 'foo.org/a/b'), False,
                    'does not look like a URL (no scheme)', 's', 'string'),
        StorageCase('s-javascript', 's', ('javascript:alert(1)',), False,
                    'javascript: URLs are not permitted', 's', 'string'),
        StorageCase('n-int', 'n', (42,), False,
                    'does not look like a URL (no scheme)', 'n', 'integer'),
    ],
}

#: The profile expectation for an error cell (never a var: target — §4).
ERROR_CELL_PROFILE = ('e', 'error', frozenset({'error'}))


# ── §6 Phase 5: cross-format re-typing ────────────────────────────────────
# "The number format re-types the value." A numeric Excel cell has no intrinsic
# date/time-ness — the format decides, and openpyxl applies that decision when it
# reads the file. So the Python type that reaches grepxcel can differ from the one
# that was written, which silently changes which `var:` types accept the cell.
#
# Every field below was MEASURED by writing the value to a real workbook and
# reading it back (tmp.local/discover_retyping.py), never predicted. Several are
# genuinely counter-intuitive:
#
#   * `12` with `[h]:mm` is twelve DAYS, not twelve hours — Excel serial 12.
#   * `12` with `HH:MM` is not a time at all: a serial >= 1 carries a date part,
#     so it comes back `datetime(1900, 1, 12)` and a `var: time` field REJECTS it.
#   * a `timedelta` of 30 hours with `HH:MM` overflows the same way, arriving as
#     `datetime(1900, 1, 1, 6, 0)`.
#   * `General` DESTROYS a date: `date(2024, 1, 15)` written with no date format
#     comes back as the bare serial `45306`, and a `var: date` field rejects it.
#     This is the root cause already documented for fixture 24 in
#     `test_schema_validation._XFAIL_DATA_QUALITY` ("one cell in the 'dayss'
#     column lacks a date number format, so openpyxl returns the raw Excel serial
#     integer") — encoded here as a first-class expectation rather than an xfail.
#   * a string is never re-typed by a format: `'42'` with `[h]:mm` stays `'42'`.

@dataclass(frozen=True)
class RetypingCase:
    """One (value, format) pair whose read-back type differs from what was
    written, or whose semantic type is decided by the format."""
    written: object
    number_format: str
    read_back: object            # exactly what openpyxl returns
    profile_storage: str
    profile_semantic: str
    accepted_for: frozenset      # declared var: types validate_type accepts
    note: str

    @property
    def case_id(self) -> str:
        return (f'{type(self.written).__name__}-{self.written!r}'
                f'-{_slug_fmt(self.number_format)}')


def _slug_fmt(fmt: str) -> str:
    return ''.join(ch if ch.isalnum() else '_' for ch in fmt).strip('_') or 'fmt'


RETYPING_CASES: tuple[RetypingCase, ...] = (
    RetypingCase(12, '[h]:mm', datetime.timedelta(days=12), 'n', 'duration',
                 frozenset(('time', 'duration', 'string')),
                 'serial 12 under an elapsed-time format is 12 days, not 12 hours'),
    RetypingCase(0, '[h]:mm', datetime.timedelta(0), 'n', 'duration',
                 frozenset(('time', 'duration', 'string')),
                 'zero stays a timedelta, not int 0'),
    RetypingCase(1.5, '[h]:mm', datetime.timedelta(days=1, seconds=43200),
                 'n', 'duration', frozenset(('time', 'duration', 'string')),
                 'fractional serial splits into days + seconds'),
    RetypingCase(12, 'HH:MM', datetime.datetime(1900, 1, 12, 0, 0), 'n', 'datetime',
                 frozenset(('date', 'datetime', 'string')),
                 'a clock format on a serial >= 1 yields a DATETIME; var: time rejects it'),
    RetypingCase(0.5, 'HH:MM', datetime.time(12, 0), 'n', 'time',
                 frozenset(('time', 'duration', 'string')),
                 'a sub-1 serial under a clock format is a real time'),
    RetypingCase(1, 'yyyy-mm-dd', datetime.datetime(1900, 1, 1, 0, 0), 'n', 'datetime',
                 frozenset(('date', 'datetime', 'string')),
                 'serial 1 is the Excel epoch'),
    RetypingCase(45000, 'yyyy-mm-dd', datetime.datetime(2023, 3, 15, 0, 0),
                 'n', 'datetime', frozenset(('date', 'datetime', 'string')),
                 'a plain int becomes a datetime purely because of the format'),
    RetypingCase(45000.5, 'yyyy-mm-dd hh:mm', datetime.datetime(2023, 3, 15, 12, 0),
                 'n', 'datetime', frozenset(('date', 'datetime', 'string')),
                 'the fraction becomes the time of day'),
    RetypingCase(0.15, '0%', 0.15, 'n', 'percentage',
                 frozenset(('number', 'currency', 'percentage', 'string')),
                 'value unchanged; only the semantic type is format-driven'),
    RetypingCase(2, '0%', 2, 'n', 'percentage',
                 frozenset(('integer', 'number', 'currency', 'percentage', 'string')),
                 'a percentage above 100% is still accepted'),
    RetypingCase(0.15, '$#,##0.00', 0.15, 'n', 'currency',
                 frozenset(('number', 'currency', 'percentage', 'string')),
                 'same float reads as currency under a currency format'),
    RetypingCase('42', '[h]:mm', '42', 's', 'string', frozenset(('string',)),
                 'a STRING is never re-typed by a number format'),
    RetypingCase(datetime.time(9, 30), '[h]:mm', datetime.timedelta(seconds=34200),
                 'n', 'duration', frozenset(('time', 'duration', 'string')),
                 'a written time comes back a timedelta under an elapsed format'),
    RetypingCase(datetime.timedelta(seconds=7200), 'HH:MM', datetime.time(2, 0),
                 'n', 'time', frozenset(('time', 'duration', 'string')),
                 'and the reverse: a timedelta comes back a time'),
    RetypingCase(datetime.timedelta(days=1, seconds=21600), 'HH:MM',
                 datetime.datetime(1900, 1, 1, 6, 0), 'n', 'datetime',
                 frozenset(('date', 'datetime', 'string')),
                 '30h overflows a clock format into a datetime; var: duration rejects it'),
    RetypingCase(datetime.date(2024, 1, 15), 'General', 45306, 'n', 'integer',
                 frozenset(('integer', 'number', 'currency', 'percentage', 'string')),
                 'General DESTROYS a date: back as a bare serial; var: date rejects it'),
    RetypingCase(datetime.datetime(2024, 1, 15, 9, 30), 'General', 45306.39583333334,
                 'n', 'number',
                 frozenset(('number', 'currency', 'percentage', 'string')),
                 'same for a datetime, as a float serial'),
)


def iter_storage_cases(types: list[str] | None = None):
    """Yield ``(type_name, StorageCase, number_format)`` for the whole matrix."""
    for type_name in (types or EXTRACTABLE_TYPES):
        for case in STORAGE_MATRIX[type_name]:
            for fmt in formats_for(type_name, case):
                yield type_name, case, fmt


def validate_pools() -> list[str]:
    """§7 hygiene: every string pool value must be non-empty, printable and
    never equal to a reserved label. Returns a list of problems (empty = OK)."""
    problems: list[str] = []
    for type_name, case, _fmt in iter_storage_cases():
        for value in case.values:
            if not isinstance(value, str):
                continue
            if not value.strip():
                problems.append(f'{type_name}/{case.key}: empty or whitespace value')
            if any(not (c.isprintable()) for c in value):
                problems.append(f'{type_name}/{case.key}: non-printable in {value!r}')
            if value.upper().startswith(RESERVED_LABEL_PREFIXES):
                problems.append(f'{type_name}/{case.key}: reserved label {value!r}')
    return problems
