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
        str_val = str(value) if value is not None else ''
        has_scheme = '://' in str_val or str_val.lower().startswith('www.')
        ok = has_scheme and _safe_match(regex, str_val, _re.DOTALL | icase, max_cell_len)
        if not has_scheme:
            return False, f'{repr(str_val)} does not look like a URL (no scheme or www.)'
        return ok, ('' if ok else f'{repr(str_val)} does not match /{regex}/')

    return False, f'unknown type {repr(field_type)}'


def infer_cell_type(values: list) -> str:
    """Return the most common grepxcel type for a list of openpyxl cell values."""
    _BOOL_STRINGS = frozenset({'TRUE', 'FALSE', 'YES', 'NO'})
    counts: dict[str, int] = {
        'datetime': 0, 'date': 0, 'time': 0, 'currency': 0,
        'integer': 0, 'boolean': 0, 'string': 0,
    }
    for v in values:
        if v is None:
            continue
        if isinstance(v, datetime.datetime):
            counts['datetime'] += 1
        elif isinstance(v, datetime.date):
            counts['date'] += 1
        elif isinstance(v, (datetime.time, datetime.timedelta)):
            counts['time'] += 1
        elif isinstance(v, bool):
            counts['boolean'] += 1
        elif isinstance(v, float):
            if v.is_integer():
                counts['integer'] += 1
            else:
                counts['currency'] += 1
        elif isinstance(v, int):
            counts['integer'] += 1
        elif isinstance(v, str) and v.strip().upper() in _BOOL_STRINGS:
            counts['boolean'] += 1
        else:
            counts['string'] += 1
    total = sum(counts.values())
    if total == 0:
        return 'string'
    return max(counts, key=lambda k: counts[k])
