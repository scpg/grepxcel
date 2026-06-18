"""
Property-based tests using Hypothesis.

Each test describes an invariant that must hold for ALL valid inputs, not just
the examples a human would think to write.  Hypothesis generates hundreds of
random cases and — if it finds a failure — shrinks the input to the minimal
example that still triggers the bug.

Targets:
  infer_cell_type           — always returns one of the valid type strings
  _type_from_number_format  — never crashes; returns None or a valid type
  is_empty                  — always returns bool; None is always True
  ExcelAnalyzer._split_into_sections  — structural invariants on the output
  PatternWriter.write       — never crashes or produces an unloadable xlsx
"""

import datetime
import os
import tempfile
from pathlib import Path

import openpyxl
from hypothesis import given, settings, strategies as st

from grepxcel.drafter import ExcelAnalyzer, PatternWriter, _type_from_number_format
from grepxcel.utils import infer_cell_type, is_empty

# ---------------------------------------------------------------------------
# Shared constants and strategies
# ---------------------------------------------------------------------------

_VALID_TYPES = frozenset({'string', 'integer', 'currency', 'percentage',
                          'date', 'datetime', 'time', 'duration'})

# Values that openpyxl returns when reading an xlsx cell.
_cell_value = st.one_of(
    st.none(),
    st.booleans(),
    st.integers(min_value=-(10**9), max_value=10**9),
    st.floats(allow_nan=False, allow_infinity=False, min_value=-1e15, max_value=1e15),
    st.text(max_size=50),
    st.dates(min_value=datetime.date(1900, 1, 1), max_value=datetime.date(2100, 12, 31)),
    st.datetimes(
        min_value=datetime.datetime(1900, 1, 1),
        max_value=datetime.datetime(2100, 12, 31),
        timezones=st.none(),          # openpyxl returns naive datetimes
    ),
)

_row  = st.lists(_cell_value, max_size=8)
_rows = st.lists(_row, max_size=30)

# Printable text safe for xlsx XML (no null bytes, surrogates, or control chars
# other than \t \n \r which openpyxl accepts).
_printable = st.text(
    alphabet=st.characters(
        blacklist_categories=('Cs', 'Cc'),
        whitelist_characters='\t\n\r',
    ),
    max_size=2000,
)

# ExcelAnalyzer instance used to call the private helper directly.
# __new__ is safe here: _split_into_sections only reads its arguments.
_analyzer = ExcelAnalyzer.__new__(ExcelAnalyzer)


# ---------------------------------------------------------------------------
# infer_cell_type
# ---------------------------------------------------------------------------

class TestInferCellTypeProperties:

    @given(st.lists(_cell_value, max_size=50))
    def test_always_returns_valid_type(self, values):
        assert infer_cell_type(values) in _VALID_TYPES

    @given(st.lists(_cell_value, max_size=50))
    def test_none_is_transparent(self, values):
        """Appending None never changes the inferred type (None is skipped)."""
        assert infer_cell_type(values) == infer_cell_type(values + [None])

    @given(st.lists(_cell_value, max_size=50))
    def test_result_is_deterministic(self, values):
        assert infer_cell_type(values) == infer_cell_type(values)

    @given(st.lists(st.integers(min_value=0, max_value=10**9), min_size=1, max_size=50))
    def test_non_negative_integers_give_integer_type(self, values):
        assert infer_cell_type(values) == 'integer'

    @given(st.lists(
        st.dates(min_value=datetime.date(1900, 1, 1), max_value=datetime.date(2100, 12, 31)),
        min_size=1, max_size=20,
    ))
    def test_dates_give_date_type(self, values):
        assert infer_cell_type(values) == 'date'

    @given(st.lists(
        st.datetimes(
            min_value=datetime.datetime(1900, 1, 1),
            max_value=datetime.datetime(2100, 12, 31),
            timezones=st.none(),
        ),
        min_size=1, max_size=20,
    ))
    def test_datetimes_give_datetime_type(self, values):
        assert infer_cell_type(values) == 'datetime'


# ---------------------------------------------------------------------------
# is_empty
# ---------------------------------------------------------------------------

class TestIsEmptyProperties:

    @given(st.one_of(
        st.none(), st.booleans(), st.integers(),
        st.text(), st.floats(allow_nan=False),
    ))
    def test_always_returns_bool(self, value):
        assert isinstance(is_empty(value), bool)

    @given(st.none())
    def test_none_is_always_empty(self, value):
        assert is_empty(value) is True

    @given(st.integers())
    def test_integers_are_never_empty(self, value):
        assert is_empty(value) is False

    @given(st.floats(allow_nan=False, allow_infinity=False))
    def test_finite_floats_are_never_empty(self, value):
        assert is_empty(value) is False

    @given(st.booleans())
    def test_booleans_are_never_empty(self, value):
        assert is_empty(value) is False


# ---------------------------------------------------------------------------
# _type_from_number_format
# ---------------------------------------------------------------------------

class TestTypeFromNumberFormatProperties:

    @given(st.text())
    def test_never_crashes(self, fmt):
        result = _type_from_number_format(fmt)
        assert result is None or result in _VALID_TYPES

    @given(st.text(alphabet='0123456789#,.%$€£¥ymdhHsM:_-"[] '))
    def test_format_like_strings_return_valid_result(self, fmt):
        result = _type_from_number_format(fmt)
        assert result is None or result in _VALID_TYPES

    @given(st.just('General'))
    def test_general_always_returns_none(self, fmt):
        assert _type_from_number_format(fmt) is None

    @given(st.text(alphabet='0%', min_size=1, max_size=10))
    def test_percentage_format_gives_percentage(self, fmt):
        if '%' in fmt:
            assert _type_from_number_format(fmt) == 'percentage'


# ---------------------------------------------------------------------------
# Cross-function invariant
# ---------------------------------------------------------------------------

@given(st.one_of(st.none(), st.text(max_size=20)))
def test_empty_value_infers_as_string(value):
    """Any value that is_empty() considers empty → infer_cell_type returns 'string'."""
    if is_empty(value):
        assert infer_cell_type([value]) == 'string'


# ---------------------------------------------------------------------------
# ExcelAnalyzer._split_into_sections
# ---------------------------------------------------------------------------

class TestSplitIntoSectionsProperties:

    @given(_rows)
    def test_no_rows_lost(self, rows):
        """Every non-empty input row appears in exactly one section."""
        sections        = _analyzer._split_into_sections(rows)
        in_sections     = sum(len(sec_rows) for _, sec_rows in sections)
        non_empty_input = sum(1 for row in rows if any(not is_empty(v) for v in row))
        assert in_sections == non_empty_input

    @given(_rows)
    def test_no_empty_sections(self, rows):
        for _, sec_rows in _analyzer._split_into_sections(rows):
            assert len(sec_rows) > 0

    @given(_rows)
    def test_start_indices_non_negative(self, rows):
        for start, _ in _analyzer._split_into_sections(rows):
            assert start >= 0

    @given(_rows)
    def test_start_indices_strictly_increasing(self, rows):
        starts = [s for s, _ in _analyzer._split_into_sections(rows)]
        assert starts == sorted(set(starts))

    @given(_rows)
    def test_every_row_in_section_has_at_least_one_non_empty_value(self, rows):
        for _, sec_rows in _analyzer._split_into_sections(rows):
            for row in sec_rows:
                assert any(not is_empty(v) for v in row)

    @given(st.lists(_row, min_size=1, max_size=30))
    def test_all_non_empty_rows_produce_at_least_one_section(self, rows):
        """If the input has any non-empty rows, there is at least one section."""
        has_content = any(any(not is_empty(v) for v in row) for row in rows)
        sections    = _analyzer._split_into_sections(rows)
        if has_content:
            assert len(sections) >= 1


# ---------------------------------------------------------------------------
# PatternWriter.write
# ---------------------------------------------------------------------------

class TestPatternWriterProperties:

    @given(_printable)
    @settings(max_examples=50)
    def test_never_crashes_on_arbitrary_text(self, text):
        fd, path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        try:
            PatternWriter().write(text, path)
            assert Path(path).exists()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    @given(_printable)
    @settings(max_examples=50)
    def test_output_is_always_loadable_by_openpyxl(self, text):
        """The written xlsx must be re-openable — not silently corrupted."""
        fd, path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        try:
            PatternWriter().write(text, path)
            wb = openpyxl.load_workbook(path)
            wb.close()
        finally:
            if os.path.exists(path):
                os.unlink(path)
