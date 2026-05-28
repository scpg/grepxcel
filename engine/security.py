"""
Security validation for input files and user-supplied regex patterns.

File checks (run before openpyxl touches any file):
  1. Extension whitelist     — only .xlsx accepted
  2. Magic bytes             — must be a real ZIP (PK signature)
  3. File size limit         — configurable, default 5 MB on disk
  4. ZIP bomb detection      — uncompressed content capped at 50 MB absolute
                               and a hard 50× expansion-ratio ceiling
  5. Macro-enabled rejected  — .xlsm / .xlsb are explicitly refused

Regex check (run at pattern-file parse time for every def: entry):
  6. ReDoS detection         — reject nested unbounded quantifiers such as
                               (a+)+, (.*)*,  ([a-z]+)+ that cause catastrophic
                               backtracking on crafted input
"""

import os
import re
import zipfile

# Python 3.11+ exposes the regex parser internals at re._parser / re._constants.
# We require >=3.11 (see pyproject) so there is a single import path: no version
# fallback to the deprecated top-level sre_parse/sre_constants modules.
#
# This is a deliberate security decision. The ReDoS detector below is built on
# this AST. A second, untested import branch in a security-critical module is a
# liability, and the legacy sre_* modules are deprecated and slated for removal.
# If a future Python ever drops re._parser, this import fails at load time and
# the whole package refuses to start — fail-closed, never silently unprotected.
import re._parser as _sre_parse
import re._constants as _sre_constants

# --- constants ----------------------------------------------------------------

_XLSX_EXTENSIONS = {'.xlsx'}                 # only pure xlsx; .xlsm/.xlsb refused
_ZIP_MAGIC       = b'PK\x03\x04'            # first 4 bytes of every ZIP file
_READ_HEADER     = 4                         # bytes to read for magic check

DEFAULT_MAX_FILE_MB         = 5              # compressed size on disk
DEFAULT_MAX_UNCOMPRESSED_MB = 50             # total uncompressed content
DEFAULT_MAX_EXPANSION_RATIO = 5              # uncompressed / compressed ceiling

# --- regex safety constants ---------------------------------------------------

_REPEAT_OPS = frozenset({_sre_constants.MAX_REPEAT, _sre_constants.MIN_REPEAT})
_MAXREPEAT  = _sre_constants.MAXREPEAT  # sentinel value meaning "unbounded"
_SUBPATTERN = _sre_constants.SUBPATTERN
_BRANCH     = _sre_constants.BRANCH
_ASSERT     = _sre_constants.ASSERT
_ASSERT_NOT = _sre_constants.ASSERT_NOT


# --- public exception ---------------------------------------------------------

class SecurityError(Exception):
    """Raised when an input file or regex fails a security check."""


# --- regex safety (ReDoS detection) ------------------------------------------

def check_regex_safety(pattern: str, field_name: str = '') -> None:
    """
    Validate a user-supplied regex for ReDoS-dangerous constructs.
    Raises SecurityError if the pattern is syntactically invalid or contains
    nested unbounded quantifiers that risk catastrophic backtracking.

    Called at pattern-file parse time for every def: entry, before the regex
    is ever applied to cell data.

    The check detects the canonical ReDoS form: a quantifier-wrapped group
    whose body itself contains an unbounded quantifier, e.g.:
        (a+)+   (.*)+ (w+)*   (x{2,})+    [where w means \\w]
    Fixed-count outer quantifiers such as (a+){3} are not flagged because they
    cannot produce exponential backtracking.
    """
    try:
        parsed = _sre_parse.parse(pattern)
    except re.error as exc:
        ctx = f' (field {field_name!r})' if field_name else ''
        raise SecurityError(
            f'Invalid regex{ctx}: {exc}  [pattern: {pattern!r}]'
        )
    if _has_nested_quantifier(parsed, in_unbounded=False):
        ctx = f' for field {field_name!r}' if field_name else ''
        raise SecurityError(
            f'Unsafe regex{ctx}: nested unbounded quantifiers risk catastrophic '
            f'backtracking (ReDoS) — {pattern!r}  '
            f'Hint: use a character class instead of a group, '
            f'e.g. [a-z]+ not ([a-z])+.'
        )


def _has_nested_quantifier(nodes, in_unbounded: bool) -> bool:
    """Walk the sre_parse AST; return True if nested unbounded quantifiers exist."""
    for op, av in nodes:
        if op in _REPEAT_OPS:
            min_count, max_count, body = av
            unbounded = (max_count == _MAXREPEAT)
            if unbounded and in_unbounded:
                return True
            if _has_nested_quantifier(body, in_unbounded or unbounded):
                return True
        elif op == _SUBPATTERN:
            # av = (group_id, add_flags, del_flags, body)  — Python 3.7+
            if _has_nested_quantifier(av[3], in_unbounded):
                return True
        elif op == _BRANCH:
            _, branches = av
            for branch in branches:
                if _has_nested_quantifier(branch, in_unbounded):
                    return True
        elif op in (_ASSERT, _ASSERT_NOT):
            _, body = av
            if _has_nested_quantifier(body, in_unbounded):
                return True
    return False


# --- XXE protection guard -----------------------------------------------------

def assert_xxe_protection() -> None:
    """
    Fail closed if openpyxl is not using defusedxml for XML parsing.

    An .xlsx is a ZIP of XML documents. Without defusedxml, openpyxl's XML
    parser is vulnerable to XXE (external entity / billion-laughs) attacks from
    a crafted workbook. openpyxl enables defusedxml automatically *only* when the
    package is importable AND the OPENPYXL_DEFUSEDXML env var is not "False".
    That protection is therefore implicit — a missing dependency or a stray env
    var would silently disable it. We assert it explicitly before every file
    load so the tool refuses to parse untrusted input without XXE protection,
    rather than parsing it unsafely.
    """
    try:
        from openpyxl.xml import DEFUSEDXML
    except Exception as exc:  # pragma: no cover - openpyxl always ships this
        raise SecurityError(
            f'Cannot verify XML (XXE) protection state in openpyxl: {exc}'
        )
    if not DEFUSEDXML:
        raise SecurityError(
            'XML parsing is NOT protected against XXE attacks. '
            'openpyxl has defusedxml disabled. Install defusedxml '
            '(pip install defusedxml) and ensure the OPENPYXL_DEFUSEDXML '
            'environment variable is not set to "False".'
        )


# --- public validation entry point --------------------------------------------

def validate_file(
    path: str,
    max_file_mb: float = DEFAULT_MAX_FILE_MB,
    max_uncompressed_mb: float = DEFAULT_MAX_UNCOMPRESSED_MB,
) -> None:
    """
    Run all security checks on *path* before it is handed to openpyxl.
    Raises SecurityError with an actionable message on any failure.
    """
    assert_xxe_protection()
    _check_exists(path)
    _check_extension(path)
    _check_magic(path)
    _check_file_size(path, max_file_mb)
    _check_zip_safety(path, max_uncompressed_mb)


# --- individual checks --------------------------------------------------------

def _check_exists(path: str) -> None:
    if not os.path.exists(path):
        raise SecurityError(f'File not found: {path!r}')
    if not os.path.isfile(path):
        raise SecurityError(f'Not a regular file: {path!r}')


def _check_extension(path: str) -> None:
    _, ext = os.path.splitext(path)
    ext = ext.lower()

    if ext in ('.xlsm', '.xlsb', '.xls'):
        raise SecurityError(
            f'{ext!r} files are not accepted. '
            f'Macro-enabled and binary Excel formats are blocked for security reasons. '
            f'Save the file as .xlsx (macro-free) before processing.'
        )
    if ext not in _XLSX_EXTENSIONS:
        raise SecurityError(
            f'Unsupported file type {ext!r}. Only .xlsx files are accepted.'
        )


def _check_magic(path: str) -> None:
    with open(path, 'rb') as fh:
        header = fh.read(_READ_HEADER)
    if header != _ZIP_MAGIC:
        raise SecurityError(
            f'{path!r} does not have a valid .xlsx signature. '
            f'The file may be corrupted or renamed from a different format.'
        )


def _check_file_size(path: str, max_mb: float) -> None:
    size_mb = os.path.getsize(path) / (1024 * 1024)
    if size_mb > max_mb:
        raise SecurityError(
            f'File is {size_mb:.1f} MB, which exceeds the {max_mb:.0f} MB limit. '
            f'Use --max-size to raise the limit if this file is legitimate.'
        )


def _check_zip_safety(path: str, max_uncompressed_mb: float) -> None:
    """
    Guard against ZIP bombs by checking both the absolute uncompressed size
    and the expansion ratio relative to the file size on disk.
    """
    try:
        with zipfile.ZipFile(path) as zf:
            total_uncompressed = sum(e.file_size for e in zf.infolist())
    except zipfile.BadZipFile:
        raise SecurityError(
            f'{path!r} is not a valid ZIP archive. '
            f'The file may be corrupted.'
        )

    uncompressed_mb = total_uncompressed / (1024 * 1024)
    if uncompressed_mb > max_uncompressed_mb:
        raise SecurityError(
            f'ZIP content expands to {uncompressed_mb:.1f} MB, '
            f'exceeding the {max_uncompressed_mb:.0f} MB uncompressed limit. '
            f'This may be a ZIP bomb or an unusually large workbook.'
        )

    compressed = os.path.getsize(path)
    if compressed > 0:
        ratio = total_uncompressed / compressed
        if ratio > DEFAULT_MAX_EXPANSION_RATIO * 10:   # hard ceiling at 50×
            raise SecurityError(
                f'ZIP expansion ratio is {ratio:.0f}:1, which is abnormally high. '
                f'This file is likely a ZIP bomb.'
            )
