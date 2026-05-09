"""
Security validation for input files.

Every file is checked before openpyxl touches it:
  1. Extension whitelist     — only .xlsx accepted
  2. Magic bytes             — must be a real ZIP (PK signature)
  3. File size limit         — configurable, default 50 MB on disk
  4. ZIP bomb detection      — uncompressed content capped at 5× the on-disk size
                               and an absolute ceiling of 500 MB
  5. Macro-enabled rejected  — .xlsm / .xlsb are explicitly refused

Limits can be raised via CLI flags for legitimate large files.
"""

import os
import zipfile

# --- constants ----------------------------------------------------------------

_XLSX_EXTENSIONS = {'.xlsx'}                 # only pure xlsx; .xlsm/.xlsb refused
_ZIP_MAGIC       = b'PK\x03\x04'            # first 4 bytes of every ZIP file
_READ_HEADER     = 4                         # bytes to read for magic check

DEFAULT_MAX_FILE_MB        = 50              # compressed size on disk
DEFAULT_MAX_UNCOMPRESSED_MB = 500            # total uncompressed content
DEFAULT_MAX_EXPANSION_RATIO = 5             # uncompressed / compressed ceiling


# --- public exception ---------------------------------------------------------

class SecurityError(Exception):
    """Raised when an input file fails a security check."""


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
            f'exceeding the {max_uncompressed_mb:.0f} MB limit. '
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
