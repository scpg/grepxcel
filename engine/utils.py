import re
import datetime

_ZERO_WIDTH = set('​‌‍﻿ ')

# Maximum length of a cell value string passed to regex matching.
# Prevents ReDoS via extremely long cell content against complex patterns.
_MAX_REGEX_INPUT_LEN = 10_000


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


def _safe_match(regex: str, text: str, flags: int = 0) -> bool:
    """
    Run re.fullmatch with a hard cap on input length.

    Extremely long strings combined with complex patterns can cause
    catastrophic backtracking (ReDoS).  Truncating at _MAX_REGEX_INPUT_LEN
    characters is safe here because any legitimate cell value that needs more
    than 10 000 characters to validate is almost certainly a data anomaly.
    """
    if len(text) > _MAX_REGEX_INPUT_LEN:
        return False
    return bool(re.fullmatch(regex, text, flags))


def validate_type(value, field_type: str, regex: str, currency_sign: str = '€') -> tuple:
    """
    Validate a cell value against a field type and regex.
    Returns (is_valid: bool, reason: str).
    """
    if field_type == 'string':
        str_val = str(value) if value is not None else ''
        ok = _safe_match(regex, str_val, re.DOTALL)
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
        ok = _safe_match(regex, str(value))
        return ok, ('' if ok else f'{value} does not match /{regex}/')

    elif field_type == 'currency':
        if isinstance(value, bool):
            return False, 'boolean is not currency'
        if not isinstance(value, (int, float)):
            return False, f'{repr(value)} is not numeric'
        str_val = str(value)
        ok = _safe_match(regex, str_val)
        return ok, ('' if ok else f'{str_val} does not match /{regex}/')

    elif field_type in ('date', 'datetime', 'timestamp'):
        ok = isinstance(value, (datetime.date, datetime.datetime))
        return ok, ('' if ok else f'{repr(value)} is not a {field_type}')

    return False, f'unknown type {repr(field_type)}'
