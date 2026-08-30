import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# ---------------------------------------------------------------------------
# Fixture file discovery helpers (shared across integration tests)
# ---------------------------------------------------------------------------
# Naming convention (new, since fixture-file rename):
#   {fixture_name}_data.xlsx
#   {fixture_name}_pattern-manual.xlsx
#   {fixture_name}_pattern-from-draft.xlsx
#   {fixture_name}_data-cdr.xlsx
#
# Legacy bare names (data.xlsx, pattern-from-draft.xlsx, …) are tried as
# fallbacks so tests stay green during any in-progress rename.
#
# Priority for xlsx:  pattern-manual  >  pattern-from-draft  >  pattern (legacy)
# Priority for csv:   pattern-manual  >  pattern-from-draft
# pattern-manual.* files are user-curated — treat as read-only.

_PATTERN_XLSX_SUFFIXES = ('pattern-manual.xlsx', 'pattern-from-draft.xlsx', 'pattern.xlsx')
_PATTERN_CSV_SUFFIXES  = ('pattern-manual.csv',  'pattern-from-draft.csv')

# Keep the old constants as aliases for any code that still imports them.
_PATTERN_XLSX_NAMES = _PATTERN_XLSX_SUFFIXES
_PATTERN_CSV_NAMES  = _PATTERN_CSV_SUFFIXES


def find_data_file(folder: str) -> str:
    """Return the data xlsx path for a fixture folder.

    Prefers the prefixed name ({name}_data.xlsx).  Falls back to the legacy
    bare name (data.xlsx) so tests work during migration.
    """
    name = os.path.basename(folder)
    new  = os.path.join(folder, f'{name}_data.xlsx')
    if os.path.exists(new):
        return new
    return os.path.join(folder, 'data.xlsx')


def find_pattern_xlsx(folder: str) -> str | None:
    """Return the highest-priority pattern xlsx in a fixture folder, or None.

    For each priority level (manual > from-draft > legacy) tries the prefixed
    name first, then the bare legacy name.
    """
    name = os.path.basename(folder)
    for suffix in _PATTERN_XLSX_SUFFIXES:
        for candidate in (f'{name}_{suffix}', suffix):
            p = os.path.join(folder, candidate)
            if os.path.exists(p):
                return p
    return None


def find_pattern_csv(folder: str) -> str | None:
    """Return a real authored pattern csv in a fixture folder, or None.

    Only returns a path when an actual committed CSV file exists — not for
    synthetically-converted temp files.  Used for the Level-2 real-file test.
    """
    name = os.path.basename(folder)
    for suffix in _PATTERN_CSV_SUFFIXES:
        for candidate in (f'{name}_{suffix}', suffix):
            p = os.path.join(folder, candidate)
            if os.path.exists(p):
                return p
    return None
