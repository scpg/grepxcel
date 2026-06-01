import re
import datetime

_ZERO_WIDTH = set('​‌‍﻿ ')

# Maximum length of a cell value string passed to regex matching.
# Prevents ReDoS via extremely long cell content against complex patterns.
_MAX_REGEX_INPUT_LEN = 1_000


def is_empty(value, empty_aliases=None) -> bool:
    """Return True if a cell value should be treated as empty."""
    if value is None:
        return True
    if isinstance(value, str):
        stripped = ''.join(
            c for c in value
            if c not in _ZERO_WIDTH and (c.isprintable() or c in '\t\n\r')
        ).strip()
        if not stripped:
            return True
        if empty_aliases and stripped in empty_aliases:
            return True
    return False


def _safe_match(regex: str, text: str, flags: int = 0,
                max_len: int = _MAX_REGEX_INPUT_LEN) -> bool:
    """
    Run re.fullmatch with a hard cap on input length to prevent ReDoS.
    Returns False when the text exceeds max_len rather than attempting the match.
    """
    if len(text) > max_len:
        return False
    return bool(re.fullmatch(regex, text, flags))


def validate_type(value, field_type: str, regex: str, currency_sign: str = '€',
                  max_cell_len: int = _MAX_REGEX_INPUT_LEN) -> tuple:
    """
    Validate a cell value against a field type and regex.
    Returns (is_valid: bool, reason: str).
    """
    if field_type == 'string':
        str_val = str(value) if value is not None else ''
        ok = _safe_match(regex, str_val, re.DOTALL, max_cell_len)
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
        ok = _safe_match(regex, str(value), max_len=max_cell_len)
        return ok, ('' if ok else f'{value} does not match /{regex}/')

    elif field_type in ('currency', 'percentage'):
        if isinstance(value, bool):
            return False, f'boolean is not {field_type}'
        if not isinstance(value, (int, float)):
            return False, f'{repr(value)} is not numeric'
        str_val = str(value)
        ok = _safe_match(regex, str_val, max_len=max_cell_len)
        return ok, ('' if ok else f'{str_val} does not match /{regex}/')

    elif field_type in ('date', 'datetime', 'timestamp'):
        ok = isinstance(value, (datetime.date, datetime.datetime))
        return ok, ('' if ok else f'{repr(value)} is not a {field_type}')

    return False, f'unknown type {repr(field_type)}'


def infer_cell_type(values: list) -> str:
    """Return the most common grepxcel type for a list of openpyxl cell values."""
    counts: dict[str, int] = {
        'datetime': 0, 'date': 0, 'currency': 0, 'integer': 0, 'string': 0,
    }
    for v in values:
        if v is None:
            continue
        if isinstance(v, datetime.datetime):
            counts['datetime'] += 1
        elif isinstance(v, datetime.date):
            counts['date'] += 1
        elif isinstance(v, bool):
            counts['string'] += 1
        elif isinstance(v, float):
            if v.is_integer():
                counts['integer'] += 1
            else:
                counts['currency'] += 1
        elif isinstance(v, int):
            counts['integer'] += 1
        else:
            counts['string'] += 1
    total = sum(counts.values())
    if total == 0:
        return 'string'
    return max(counts, key=lambda k: counts[k])
