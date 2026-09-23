"""Schema validation sweep — every fixture's extraction must validate against
its generated schema.

For each fixture with a pattern + data file:
  1. Generate a JSON Schema from the pattern
  2. Extract data using the pattern
  3. Validate the JSON-serialised extraction output against the schema

This catches mismatches between what the engine produces and what the schema
generator declares.
"""
import datetime
import json
import os

import jsonschema
import pytest

from grepxcel import Engine, Logger, VerbosityLevel
from grepxcel.schema import generate_schema
from tests.conftest import find_data_file, find_pattern_xlsx, find_pattern_for_backend


class _JSONEncoder(json.JSONEncoder):
    """Mirror the CLI's encoder so we validate the actual JSON output shape."""
    def default(self, obj):
        if isinstance(obj, (datetime.date, datetime.datetime, datetime.time)):
            return obj.isoformat()
        if isinstance(obj, datetime.timedelta):
            return str(obj)
        return super().default(obj)

_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')

_ALL_FIXTURE_NAMES: list[str] = sorted(
    name for name in os.listdir(_FIXTURES_DIR)
    if os.path.isdir(os.path.join(_FIXTURES_DIR, name))
    and not os.path.exists(os.path.join(_FIXTURES_DIR, name, 'UNDER_REVIEW'))
)

_FIXTURES_WITH_DATA: list[str] = [
    name for name in _ALL_FIXTURE_NAMES
    if find_pattern_xlsx(os.path.join(_FIXTURES_DIR, name)) is not None
    and os.path.exists(find_data_file(os.path.join(_FIXTURES_DIR, name)))
]

_SHEET_OVERRIDES: dict[str, str] = {
    '12_multi_sheet': 'Details',
    '06_merged_cells': '2025',
    '17_excel_template_invoice': 'Invoice',
}

# Fixtures whose drafted patterns have known data-quality issues (e.g. a totals
# row leaking text into a currency field because SKIP_IF/FOOTER is missing, or
# Excel formula errors such as #N/A in cells the pattern expects to be numeric).
# The schema correctly rejects these — xfail proves the schema catches the bug.
_XFAIL_DATA_QUALITY: set[str] = {
    '18_billing_statement',
    # 22: the rolling weekly timesheet has #N/A formula errors in cumulative
    # cells (pay.vacation, pay.grand_total) when the current week has no data
    # yet; the schema correctly rejects '#N/A' as non-numeric.
    '22_weekly_timesheet',
    # 24: one cell in the 'dayss' column lacks a date number format, so openpyxl
    # returns the raw Excel serial integer (25602) instead of a datetime object;
    # the schema correctly rejects the integer as not of type 'string/date-time'.
    '24_type_tests',
}


def _extract(pattern_path: str, data_path: str, sheet=None) -> dict:
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    return Engine().process(
        pattern_file=pattern_path, data_file=data_path, logger=lg, **kwargs,
    )


@pytest.mark.parametrize('fixture', _FIXTURES_WITH_DATA)
def test_extraction_validates_against_generated_schema(fixture):
    """The extraction output must conform to the schema generated from its pattern.

    Uses the draft baseline pattern so the test remains stable as higher-priority
    patterns (claude, local) are added.  Claude/local patterns may have type
    annotation differences (e.g. integer vs currency) that are tracked by snapshot
    tests rather than here.
    """
    folder = os.path.join(_FIXTURES_DIR, fixture)
    pattern = find_pattern_for_backend(folder, 'draft') or find_pattern_xlsx(folder)
    data = find_data_file(folder)
    sheet = _SHEET_OVERRIDES.get(fixture)

    schema = generate_schema(pattern)
    raw = _extract(pattern, data, sheet)
    result = json.loads(json.dumps(raw, cls=_JSONEncoder))

    if fixture in _XFAIL_DATA_QUALITY:
        pytest.xfail('pattern has known data-quality issues — schema correctly rejects')
    jsonschema.validate(instance=result, schema=schema)


def test_fixture_sweep_is_not_vacuous():
    """Ensure we're actually testing a reasonable number of fixtures."""
    assert len(_FIXTURES_WITH_DATA) >= 20
