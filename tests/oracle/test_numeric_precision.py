"""Where a number in a spreadsheet stops being the number that was entered.

An Excel numeric cell is an IEEE-754 double. A double has a 53-bit mantissa, so
every integer up to 2**53 is representable and above it they are not: the
spacing becomes 2, then 4, then 8. Nothing in the file records that a value was
adjusted, and ``validate_type`` sees a perfectly ordinary integer afterwards.

This matters because long numeric identifiers are ordinary spreadsheet content —
order numbers, bank references, barcodes, EU VAT numbers, Snowflake ids. A
19-digit id stored as a number comes back off by 256, and ``var: integer``
accepts it without complaint.

The tension this creates in grepxcel's own type system is recorded in
:class:`TestTextIsExactButRejectedAsInteger`: storing such an id as *text* keeps
it exact, and that is the storage the type system is least happy with. The
guidance that follows from it — declare long-id fields ``string``, not
``integer`` — is in ``docs/pattern-file.md``.

Discovered from the owner's Excel serial reference table, whose own arithmetic
stops holding at exactly 1.17e16 for this reason
(``tmp.local/analyse_reference_table.py``).
"""
from __future__ import annotations

import openpyxl
import pytest

from grepxcel.cell_taxonomy import classify_value
from grepxcel.utils import validate_type
from tests.oracle.catalog import _INTEGRAL_FLOAT_POOL, _INT_POOL
from tests.unit.test_random_type_fixtures import write_data_xlsx

pytestmark = pytest.mark.oracle

#: The largest integer for which every smaller integer is exactly representable
#: as a double. Above this the gap between representable integers is 2.
TWO_53 = 2 ** 53


def _read_back(tmp_path, value, fmt=None):
    path = str(tmp_path / 'd.xlsx')
    write_data_xlsx(path, {(1, 1): (value, fmt)})
    return openpyxl.load_workbook(path, data_only=True).active.cell(row=1, column=1).value


class TestTheExactIntegerBoundary:
    """N-1, N, N+1 around the code constant that actually governs this."""

    def test_two_53_is_what_we_think_it_is(self):
        assert TWO_53 == 9_007_199_254_740_992

    @pytest.mark.parametrize('value', [
        999_999_999,
        10 ** 15,
        TWO_53 - 1,
        TWO_53,
    ])
    def test_at_or_below_two_53_every_integer_survives(self, tmp_path, value):
        got = _read_back(tmp_path, value)
        assert got == value and isinstance(got, int), f'{value} -> {got!r}'

    @pytest.mark.parametrize('value,expected', [
        (TWO_53 + 1, TWO_53),          # rounds DOWN to the even neighbour
        (TWO_53 + 3, TWO_53 + 4),      # rounds UP to the even neighbour
    ])
    def test_odd_integers_above_two_53_are_silently_moved(self, tmp_path,
                                                          value, expected):
        """Above 2**53 the representable integers are the even ones, so an odd
        value snaps to a neighbour — down or up, whichever is nearer. It comes
        back a plain int, so nothing downstream can tell it was changed."""
        got = _read_back(tmp_path, value)
        assert got == expected
        assert got != value, 'expected a silent adjustment'

    @pytest.mark.parametrize('value', [TWO_53 + 2, TWO_53 + 4])
    def test_even_integers_just_above_two_53_still_survive(self, tmp_path, value):
        """The failure is not 'everything above 2**53 breaks' — it is that the
        spacing is 2. Pinned so the boundary is described accurately."""
        got = _read_back(tmp_path, value)
        assert got == value

    @pytest.mark.parametrize('value,held,off_by', [
        (1234567890123456789, 1234567890123457024, 235),
    ])
    def test_a_long_id_is_off_by_much_more_than_one(self, tmp_path, value,
                                                    held, off_by):
        """At 1e18 the representable integers are 256 apart, so a 19-digit
        identifier is not merely imprecise, it is a different number.

        Note the comparison is done in EXACT integer arithmetic. Measuring the
        error as ``got - value`` while ``got`` is still a float reports 256,
        because that subtraction is itself performed in floating point — the
        very effect under test. int() first, then subtract."""
        got = _read_back(tmp_path, value)
        assert isinstance(got, float)
        assert int(got) == held
        assert int(got) - value == off_by


class TestTypeChangesAtTenToTheSixteen:
    @pytest.mark.parametrize('value', [10 ** 16, 10 ** 20])
    def test_large_integers_come_back_as_floats(self, tmp_path, value):
        """A second, separate effect: past ~1e16 openpyxl returns a float rather
        than an int, even when the value itself is unharmed."""
        got = _read_back(tmp_path, value)
        assert isinstance(got, float)
        assert got == value

    def test_a_float_that_is_whole_still_classifies_as_integer(self, tmp_path):
        """So the type change is invisible in the profile: semantic stays
        'integer' because the float is whole."""
        got = _read_back(tmp_path, 10 ** 16)
        assert classify_value(got, 'General').semantic_type == 'integer'


class TestTheCorruptedValueIsAccepted:
    """The part that makes this worth a test rather than a footnote."""

    def test_var_integer_accepts_a_value_that_was_silently_changed(self, tmp_path):
        got = _read_back(tmp_path, TWO_53 + 1)
        assert got == TWO_53          # already not what was written
        ok, reason = validate_type(got, 'integer', '.*')
        assert ok, f'unexpectedly rejected: {reason}'

    def test_var_integer_accepts_an_off_by_256_identifier(self, tmp_path):
        got = _read_back(tmp_path, 1234567890123456789)
        ok, reason = validate_type(got, 'integer', '.*')
        assert ok, f'unexpectedly rejected: {reason}'

    def test_nothing_in_the_classification_marks_it(self, tmp_path):
        """No flag, no warning, no distinguishing feature — which is exactly why
        the mitigation has to be a pattern-authoring choice, not a check."""
        prof = classify_value(_read_back(tmp_path, TWO_53 + 1), 'General')
        assert prof.semantic_type == 'integer'
        assert prof.flags == []


class TestTextIsExactButRejectedAsInteger:
    """The workaround, and the friction it hits.

    Storing a long id as text is the only way to keep it exact, and it is the
    form ``var: integer`` refuses. A pattern must declare these fields
    ``string``.
    """

    IDS = ['9007199254740993', '1234567890123456789', '0012345']

    @pytest.mark.parametrize('text', IDS)
    def test_text_keeps_every_digit(self, tmp_path, text):
        assert _read_back(tmp_path, text) == text

    @pytest.mark.parametrize('text', IDS)
    def test_var_integer_rejects_it(self, tmp_path, text):
        got = _read_back(tmp_path, text)
        ok, reason = validate_type(got, 'integer', '.*')
        assert not ok
        assert 'is not integer type' in reason

    @pytest.mark.parametrize('text', IDS)
    def test_var_string_accepts_it(self, tmp_path, text):
        got = _read_back(tmp_path, text)
        ok, reason = validate_type(got, 'string', '.*')
        assert ok, reason

    @pytest.mark.parametrize('text', IDS)
    def test_profile_flags_it_as_text_forced_numeric(self, tmp_path, text):
        """Usually a warning sign ('this number is stored as text'); for a long
        id it is the correct storage. The flag is the signal, not the verdict."""
        prof = classify_value(_read_back(tmp_path, text), 'General')
        assert prof.semantic_type == 'string'
        assert 'text_forced_numeric' in prof.flags


def test_the_oracle_catalog_stays_inside_the_exact_range():
    """Cross-reference guard. The generated matrix asserts exact round-trips for
    every integer it writes, so a pool value above 2**53 would make the suite
    fail for a reason that has nothing to do with grepxcel. The largest pool
    value today is 1e15, comfortably below."""
    for value in (*_INT_POOL, *_INTEGRAL_FLOAT_POOL):
        assert abs(value) <= TWO_53, (
            f'catalog pool value {value!r} is at or above 2**53 and is not '
            f'guaranteed to round-trip exactly through a .xlsx'
        )
