import datetime
import os

# Use the third-party `regex` engine instead of stdlib `re` for one reason:
# it supports a hard per-match `timeout=`, which stdlib `re` does not. A static
# ReDoS guard (security.check_regex_safety) can only catch known-dangerous
# *shapes*; it can never be complete. The timeout is the actual guarantee — no
# single match can run longer than _REGEX_TIMEOUT_SECONDS, so a catastrophic
# pattern that slips past the static guard is bounded instead of hanging.
# `regex` is a superset of `re`, so every pattern users already write still works.
import regex as _re

_ZERO_WIDTH = set('​‌‍﻿ ')

# Maximum length of a cell value string passed to regex matching.
# Prevents ReDoS via extremely long cell content against complex patterns.
_MAX_REGEX_INPUT_LEN = 1_000

# URL schemes accepted by validate_type for 'url' fields.
# Covers what realistically appears in Excel: web, file transfer, email/messaging, local files.
_ALLOWED_URL_SCHEMES = frozenset({
    'http', 'https',           # web
    'ftp', 'ftps',             # file transfer
    'mailto', 'tel', 'sms',   # email / messaging
    'file',                    # local files
})

# Hard wall-clock bound on a single regex match (seconds). Safe patterns on a
# <=1000-char cell finish in microseconds; this only ever fires on catastrophic
# backtracking. Override with GREPXCEL_REGEX_TIMEOUT for unusual workloads.
def _regex_timeout() -> float:
    raw = os.environ.get('GREPXCEL_REGEX_TIMEOUT', '').strip()
    if raw:
        try:
            val = float(raw)
            if val > 0:
                return val
        except ValueError:
            pass
    return 0.25


# Leading characters that spreadsheet apps (Excel, LibreOffice, Sheets) interpret
# as the start of a formula. A cell value extracted from an untrusted source file
# that begins with one of these is a CSV/formula-injection vector (CWE-1236) when
# written back into a CSV or XLSX a human will open. We neutralise by prefixing a
# single quote, the OWASP-recommended mitigation, which forces text interpretation.
_FORMULA_LEAD_CHARS = ('=', '+', '-', '@', '\t', '\r')


def neutralize_formula(value):
    """Defuse formula/CSV injection in a value bound for a CSV or XLSX cell.

    String values that begin with a formula-trigger character are prefixed with a
    single quote so spreadsheet apps treat them as literal text instead of an
    executable formula. Non-string values (numbers, dates, bools, None) cannot be
    formulas and pass through unchanged.
    """
    if isinstance(value, str) and value.startswith(_FORMULA_LEAD_CHARS):
        return "'" + value
    return value


def flatten_nested(obj: dict, prefix: str = '') -> list:
    """Flatten a nested dict into ordered ``(dotted.key, value)`` pairs.

    Order is preserved depth-first. Non-dict leaves terminate a branch::

        flatten_nested({'a': {'b': 1}, 'c': 2}) -> [('a.b', 1), ('c', 2)]
    """
    items: list = []
    for k, v in obj.items():
        key = f'{prefix}.{k}' if prefix else k
        if isinstance(v, dict):
            items.extend(flatten_nested(v, key))
        else:
            items.append((key, v))
    return items


def flatten_table_instances(result: dict) -> dict:
    """Collapse table instance wrappers produced by the ``nested`` output format.

    Transforms each table key from a list of instance envelopes::

        {"items": [{"_source": {...}, "data": [row, ...]}, ...]}

    into a flat list of data rows::

        {"items": [row, row, ...]}

    Scalar values and the ``_meta`` key are passed through unchanged.
    """
    out: dict = {}
    for key, value in result.items():
        if key == '_meta':
            out[key] = value
        elif isinstance(value, list):
            flat: list = []
            for inst in value:
                if isinstance(inst, dict) and 'data' in inst:
                    flat.extend(inst['data'])
                else:
                    flat.append(inst)
            out[key] = flat
        else:
            out[key] = value
    return out


def sanitize_for_prompt(value, max_len: int = 200) -> str:
    """Make an untrusted cell value safe to embed in an LLM prompt.

    Spreadsheet cells fed to the ``draft`` analyser come from untrusted files. A
    cell containing newlines plus fake instructions ("\\nSYSTEM: ignore all
    prior…") is a prompt-injection vector. This collapses every run of
    whitespace (including newlines and tabs) to a single space, strips remaining
    non-printable control characters, and truncates to *max_len* so a cell can
    neither break onto its own line nor bloat the prompt.
    """
    s = str(value) if value is not None else ''
    s = _re.sub(r'\s+', ' ', s).strip()
    s = ''.join(ch for ch in s if ch.isprintable())
    if len(s) > max_len:
        s = s[:max_len] + '…'
    return s


def is_empty(value, empty_aliases=None, ignore_case: bool = False) -> bool:
    """Return True if a cell value should be treated as empty.

    ``ignore_case`` makes alias matching case-insensitive, consistent with
    the ``config: | ignore.case | yes`` setting.  When True, a cell value
    ``'n/a'`` will match an alias declared as ``'N/A'``.
    """
    if value is None:
        return True
    if isinstance(value, str):
        stripped = ''.join(
            c for c in value
            if c not in _ZERO_WIDTH and (c.isprintable() or c in '\t\n\r')
        ).strip()
        if not stripped:
            return True
        if empty_aliases:
            if ignore_case:
                stripped_lower = stripped.lower()
                if any(stripped_lower == a.lower() for a in empty_aliases):
                    return True
            elif stripped in empty_aliases:
                return True
    return False


# ── Temporal text coercion ──────────────────────────────────────────────────
#
# A date can arrive as *text* rather than a real Excel date: exported from
# another system, typed with a leading apostrophe, or written by a tool that
# never applied a date number format. Before this, such a cell was simply
# rejected ("'2024-01-15' is not a date") even though the pattern had declared
# exactly what it was meant to be. When the pattern says a field is a date, and
# the text says unambiguously which date, converting it is strictly better than
# refusing it.
#
# What is deliberately NOT done here, and why: no lenient/heuristic parsing.
# `dateutil.parser.parse` was evaluated and rejected — measured, not assumed:
#
#   '01/02/2024'  ->  Jan 2, or Feb 1 with dayfirst=True. Same input, two
#                     different dates, no warning either way. Silently wrong
#                     data is worse than a refusal.
#   '09:30'       ->  datetime(<today>, 9, 30). It injects the CURRENT DATE, so
#                     the same file extracts differently tomorrow.
#   '2:00'        ->  datetime(<today>, 2, 0), not timedelta(hours=2).
#   '30:00'       ->  ParserError. It cannot represent elapsed time at all.
#
# So: unambiguous ISO 8601 by default (stdlib, no dependency), elapsed-time
# H:MM[:SS] for durations (which no library handles), and for anything else the
# user declares the format explicitly via `config: | date.format | %d/%m/%Y`.
# Declared beats guessed — it gives the same answer on every machine.

#: Field types whose text form is coerced. Mirrors validate_type's own grouping.
_DATE_TYPES = frozenset({'date', 'datetime', 'timestamp'})
_TIME_TYPES = frozenset({'time', 'duration'})
TEMPORAL_TYPES = _DATE_TYPES | _TIME_TYPES

#: Elapsed time: hours are unbounded (30:00 is thirty hours, not a clock time).
_DURATION_RE = _re.compile(r'^(\d{1,6}):([0-5]\d)(?::([0-5]\d))?$')


def parse_duration_text(text: str):
    """Parse elapsed time ``H:MM`` / ``H:MM:SS`` into a ``timedelta``.

    Hours are unbounded on purpose — ``'30:00'`` is thirty hours, which is a
    perfectly ordinary timesheet value and is exactly what a clock-time parser
    refuses. Returns None if *text* is not that shape.
    """
    match = _DURATION_RE.match(text.strip())
    if not match:
        return None
    hours, minutes, seconds = match.group(1), match.group(2), match.group(3)
    return datetime.timedelta(hours=int(hours), minutes=int(minutes),
                              seconds=int(seconds or 0))


def coerce_temporal_text(value, field_type: str, date_format: str | None = None,
                         time_format: str | None = None):
    """Convert a text cell into the temporal object its declared type implies.

    Returns the converted value, or **None** when *value* is not text, the type
    is not temporal, or the text cannot be converted unambiguously — in which
    case the caller leaves the value alone and normal validation reports it.

    A ``date`` field yields a ``datetime`` at midnight rather than a ``date``,
    deliberately: that is what openpyxl returns for a real Excel date cell (it
    has no date-only type), so text and native cells produce the same output
    shape instead of one field's JSON being ``'2024-01-15'`` and another's
    ``'2024-01-15T00:00:00'`` depending on how the sheet happened to store it.
    """
    if not isinstance(value, str) or field_type not in TEMPORAL_TYPES:
        return None
    text = value.strip()
    if not text:
        return None

    # A declared format is tried FIRST, then ISO as a fallback — not instead of
    # it. Both are deterministic, so trying both cannot introduce ambiguity, and
    # a sheet that mixes '31/12/2024' with '2024-12-31' still converts fully.
    if field_type in _DATE_TYPES:
        attempts = [_iso_datetime]
        if date_format:
            attempts.insert(0, lambda t: _by_format(t, date_format))
    elif field_type == 'time':
        # time and duration both accept a clock time and an elapsed duration (so
        # does validate_type); the declared type only decides which is tried first.
        attempts = [_iso_time, parse_duration_text]
        if time_format:
            attempts.insert(0, lambda t: _by_format(t, time_format, as_time=True))
    else:  # duration
        attempts = [parse_duration_text, _iso_time]
        if time_format:
            attempts.insert(0, lambda t: _by_format(t, time_format, as_time=True))

    for attempt in attempts:
        converted = attempt(text)
        if converted is not None:
            return converted
    return None


def _by_format(text: str, fmt: str, as_time: bool = False):
    try:
        parsed = datetime.datetime.strptime(text, fmt)
    except (ValueError, TypeError):
        return None
    return parsed.time() if as_time else parsed


def _iso_datetime(text: str):
    """ISO date or datetime, always returned as a ``datetime`` (see docstring)."""
    try:
        return datetime.datetime.fromisoformat(text)
    except ValueError:
        pass
    try:
        day = datetime.date.fromisoformat(text)
    except ValueError:
        return None
    return datetime.datetime(day.year, day.month, day.day)


def _iso_time(text: str):
    """ISO clock time, but only in a form that cannot be mistaken for a number.

    A colon is required. Python 3.11+ widened ``time.fromisoformat`` to accept
    bare-hour and compact forms, which makes it dangerous for cells of unknown
    provenance: it reads ``'12'`` as 12:00, ``'1230'`` as 12:30 and — the one
    that matters — ``'2024'`` as **20:24**, so a year, an ID or a quantity in a
    time-typed column would silently become a plausible-looking time. Refusing
    them costs nothing real: ISO basic format (``'123045'``) is vanishingly rare
    in spreadsheets, and a file that does use it can declare
    ``config: | time.format | %H%M%S``.
    """
    if ':' not in text:
        return None
    try:
        return datetime.time.fromisoformat(text)
    except ValueError:
        return None


def _safe_match(regex: str, text: str, flags: int = 0,
                max_len: int = _MAX_REGEX_INPUT_LEN) -> bool:
    """
    Run a full-match with two independent ReDoS defenses:
      1. a hard cap on input length (skip the match entirely if exceeded), and
      2. a hard per-match wall-clock timeout via the `regex` engine.

    Returns False when the text exceeds max_len, when the match times out
    (catastrophic backtracking), or when there is simply no match.
    """
    if len(text) > max_len:
        return False
    try:
        return bool(_re.fullmatch(regex, text, flags, timeout=_regex_timeout()))
    except TimeoutError:
        # Catastrophic backtracking — bounded, not hung. Treat as no-match.
        return False


def validate_type(value, field_type: str, regex: str, currency_sign: str = '€',
                  max_cell_len: int = _MAX_REGEX_INPUT_LEN,
                  ignore_case: bool = False) -> tuple:
    """
    Validate a cell value against a field type and regex.
    Returns (is_valid: bool, reason: str).

    When ignore_case is True, regex matching is case-insensitive
    (driven by the `config: | ignore.case | yes` pattern setting).
    """
    icase = _re.IGNORECASE if ignore_case else 0

    if field_type in ('string', 'text'):
        str_val = str(value) if value is not None else ''
        ok = _safe_match(regex, str_val, _re.DOTALL | icase, max_cell_len)
        return ok, ('' if ok else f'{repr(str_val)} does not match /{regex}/')

    elif field_type == 'integer':
        if isinstance(value, bool):
            return False, 'boolean is not integer'
        if isinstance(value, float):
            if not value.is_integer():
                return False, f'{value} is not a whole number'
            value = int(value)
        if not isinstance(value, int):
            return False, f'{repr(value)} is not integer type'
        ok = _safe_match(regex, str(value), icase, max_cell_len)
        return ok, ('' if ok else f'{value} does not match /{regex}/')

    elif field_type in ('currency', 'percentage', 'number', 'float', 'decimal'):
        if isinstance(value, bool):
            return False, f'boolean is not {field_type}'
        if not isinstance(value, (int, float)):
            return False, f'{repr(value)} is not numeric'
        str_val = str(value)
        ok = _safe_match(regex, str_val, icase, max_cell_len)
        return ok, ('' if ok else f'{str_val} does not match /{regex}/')

    elif field_type in ('boolean', 'bool'):
        # Accepted forms:
        #   Python bool  — Excel formula =TRUE()/=FALSE() (openpyxl returns bool)
        #   int 0 or 1   — common in CSV exports and number-formatted columns
        #   str          — "TRUE"/"FALSE", "YES"/"NO", "1"/"0" (case-insensitive, trimmed)
        # bool is a subclass of int; the numeric branches above reject it early,
        # so the isinstance(value, bool) check here is always reached for bools.
        _BOOL_STRINGS = frozenset({'TRUE', 'FALSE', 'YES', 'NO', '1', '0'})
        if isinstance(value, bool):
            str_val = str(value)          # "True" or "False"
        elif isinstance(value, int) and value in (0, 1):
            str_val = str(value)          # "0" or "1"
        elif isinstance(value, str) and value.strip().upper() in _BOOL_STRINGS:
            str_val = value.strip()
        else:
            return False, (
                f'{repr(value)} is not a boolean '
                f'(accepted: True/False, Yes/No, 1/0)'
            )
        ok = _safe_match(regex, str_val, icase, max_cell_len)
        return ok, ('' if ok else f'{repr(str_val)} does not match /{regex}/')

    elif field_type in ('time', 'duration'):
        # Clock-time cells → datetime.time; duration cells ([h]:mm) → timedelta.
        # Both are accepted for 'time' and 'duration' because Excel uses [h]:mm
        # for elapsed-time columns that users naturally mark as "time".
        ok = isinstance(value, (datetime.time, datetime.timedelta))
        return ok, ('' if ok else f'{repr(value)} is not a time')

    elif field_type in ('date', 'datetime', 'timestamp'):
        ok = isinstance(value, (datetime.date, datetime.datetime))
        return ok, ('' if ok else f'{repr(value)} is not a {field_type}')

    elif field_type == 'url':
        from urllib.parse import urlparse as _urlparse
        str_val = str(value) if value is not None else ''
        try:
            _parsed = _urlparse(str_val)
        except Exception:
            return False, f'{repr(str_val)} is not a valid URL'
        _scheme = _parsed.scheme.lower()
        if not _scheme:
            return False, f'{repr(str_val)} does not look like a URL (no scheme)'
        if _scheme == 'javascript':
            return False, 'javascript: URLs are not permitted'
        if _scheme not in _ALLOWED_URL_SCHEMES:
            return False, (
                f'URL scheme {repr(_scheme)} is not supported '
                f'(accepted: {", ".join(sorted(_ALLOWED_URL_SCHEMES))})'
            )
        ok = _safe_match(regex, str_val, _re.DOTALL | icase, max_cell_len)
        return ok, ('' if ok else f'{repr(str_val)} does not match /{regex}/')

    elif field_type == 'image':
        # Image cells carry embedded binary data, not a text value.
        # Presence is confirmed by scan_image_cells(); validate_type() always passes.
        return True, ''

    return False, f'unknown type {repr(field_type)}'


_INFER_LEGACY_MAP = {
    # Translates cell_taxonomy's richer semantic vocabulary back to this
    # function's historical 7-value one (no percentage/number/url/error
    # buckets — those simply didn't exist before cell_taxonomy.py).
    'datetime': 'datetime', 'date': 'date', 'time': 'time',
    'duration': 'time', 'boolean': 'boolean', 'integer': 'integer',
    'number': 'currency', 'percentage': 'currency', 'currency': 'currency',
    'string': 'string', 'url': 'string', 'error': 'string',
}


def infer_cell_type(values: list) -> str:
    """Return the most common grepxcel type for a list of openpyxl cell values.

    Thin wrapper over cell_taxonomy.classify_value() — the canonical
    classifier — translated back to this function's original vocabulary so
    existing callers see no behaviour change. Verified against
    tests/unit/test_property_based.py and tests/unit/test_suggester.py.
    """
    from .cell_taxonomy import classify_value
    # Fixed insertion order, matching the pre-consolidation implementation
    # exactly, so max()'s tie-breaking (first-seen wins) is unchanged.
    counts: dict[str, int] = {
        'datetime': 0, 'date': 0, 'time': 0, 'currency': 0,
        'integer': 0, 'boolean': 0, 'string': 0,
    }
    total = 0
    for v in values:
        if v is None:
            continue
        semantic = classify_value(v).semantic_type
        legacy = _INFER_LEGACY_MAP.get(semantic, 'string')
        counts[legacy] += 1
        total += 1
    if total == 0:
        return 'string'
    return max(counts, key=lambda k: counts[k])
