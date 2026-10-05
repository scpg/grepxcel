"""Excel's date serial has two defects at its low end. Both are pinned here.

A numeric Excel cell is a serial counted from an epoch, and the owner's reference
table for that representation maps serial 0 to 1899-12-30 with plain arithmetic.
That mapping is correct for the *representation* and wrong as a prediction of
what any Excel-compatible reader returns, because Excel treats 1900 as a leap
year — it has a 29 February that never existed. openpyxl reproduces that, so:

  * for serials 1..59 the returned date is one day AHEAD of epoch + n days;
  * serial 59 and serial 60 both return 1900-02-28 — two distinct stored values,
    one date, no way to tell them apart afterwards;
  * from serial 61 on, the naive arithmetic and the returned date agree forever.

And at the very bottom there is a two-day hole where a *date* silently stops
being a date: 1899-12-30 and 1899-12-31 (serials 0 and 1) come back as
``datetime.time(0, 0)``. 1899-12-29 survives, and so does 1900-01-01 — only those
two days fall through.

None of this is a grepxcel defect; it is the file format, and grepxcel reports it
rather than passing it on silently. These tests exist so the behaviour is
documented and cannot change unnoticed — the low end of the date range is exactly
where a boundary value in ``catalog._DATE_POOL`` (``date(1900, 1, 1)``) sits.
"""
from __future__ import annotations

import datetime

import openpyxl
import pytest

from grepxcel.cell_taxonomy import classify_value
from grepxcel.engine import Engine
from grepxcel.logger import Logger, Severity, VerbosityLevel
from grepxcel.utils import validate_type
from tests.unit.test_random_type_fixtures import write_data_xlsx, write_pattern_csv

pytestmark = pytest.mark.oracle

#: openpyxl's own epoch constant. Day 0 of the serial scheme.
TRUE_EPOCH = datetime.datetime(1899, 12, 30)

#: Below this serial, Excel's phantom 29 Feb 1900 shifts every date by one day.
LEAP_BUG_SERIAL = 60


def _read_back(tmp_path, value, fmt='yyyy-mm-dd'):
    path = str(tmp_path / 'd.xlsx')
    write_data_xlsx(path, {(1, 1): (value, fmt)})
    return openpyxl.load_workbook(path, data_only=True).active.cell(row=1, column=1).value


class TestEpochAndLeapYearBug:
    def test_openpyxl_epoch_constant_is_1899_12_30(self):
        """The premise of everything below. If openpyxl ever changes its epoch,
        every expectation in this file is void, so it is asserted, not assumed."""
        from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900
        assert CALENDAR_WINDOWS_1900 == TRUE_EPOCH

    @pytest.mark.parametrize('serial,expected', [
        (1, datetime.date(1900, 1, 1)),
        (2, datetime.date(1900, 1, 2)),
        (59, datetime.date(1900, 2, 28)),
        (60, datetime.date(1900, 2, 28)),   # the phantom 29 Feb, folded onto 28
        (61, datetime.date(1900, 3, 1)),
        (100, datetime.date(1900, 4, 9)),
        (45000, datetime.date(2023, 3, 15)),
    ])
    def test_serial_to_date(self, tmp_path, serial, expected):
        got = _read_back(tmp_path, serial)
        assert got == datetime.datetime(expected.year, expected.month, expected.day), (
            f'serial {serial} -> {got!r}, expected {expected}'
        )

    @pytest.mark.parametrize('serial', [1, 2, 30, 58, 59])
    def test_below_the_bug_the_date_is_one_day_ahead_of_naive_arithmetic(
            self, tmp_path, serial):
        """Plain 'epoch + n days' under-counts by a day here. Anyone converting
        serials by hand for early-1900 data gets the wrong date."""
        got = _read_back(tmp_path, serial)
        naive = TRUE_EPOCH + datetime.timedelta(days=serial)
        assert got == naive + datetime.timedelta(days=1)

    @pytest.mark.parametrize('serial', [60, 61, 62, 100, 45000])
    def test_from_the_bug_onward_naive_arithmetic_is_correct(self, tmp_path, serial):
        got = _read_back(tmp_path, serial)
        assert got == TRUE_EPOCH + datetime.timedelta(days=serial)

    def test_serials_59_and_60_are_indistinguishable_afterwards(self, tmp_path):
        """Two different numbers in the file, one date out. A value stored as 60
        cannot be recovered — this is data loss, not a rounding artefact."""
        fifty_nine = _read_back(tmp_path, 59)
        sixty = _read_back(tmp_path, 60)
        assert fifty_nine == sixty == datetime.datetime(1900, 2, 28)

    def test_python_has_no_29_february_1900(self):
        """The reason the fold happens at all: 1900 was not a leap year."""
        with pytest.raises(ValueError):
            datetime.date(1900, 2, 29)


class TestPre1900Hole:
    """Exactly two days where a date silently stops being a date."""

    @pytest.mark.parametrize('day', [
        datetime.date(1899, 12, 30),   # serial 0
        datetime.date(1899, 12, 31),   # serial 1
    ])
    def test_the_date_is_replaced_by_midnight(self, tmp_path, day):
        got = _read_back(tmp_path, day)
        assert isinstance(got, datetime.time) and not isinstance(got, datetime.datetime), (
            f'{day} came back as {got!r} ({type(got).__name__}); expected a time'
        )
        assert got == datetime.time(0, 0)
        # And the classifier agrees it is no longer a date.
        assert classify_value(got, 'yyyy-mm-dd').semantic_type == 'time'
        # So a var: date field refuses it rather than storing a wrong date.
        ok, reason = validate_type(got, 'date', '.*')
        assert not ok and 'is not a date' in reason

    @pytest.mark.parametrize('day', [
        datetime.date(1899, 12, 29),   # just before the hole
        datetime.date(1900, 1, 1),     # just after it — a catalog boundary value
        datetime.date(1900, 2, 28),
        datetime.date(1900, 3, 1),
        datetime.date(2024, 1, 15),
        datetime.date(9999, 12, 31),
    ])
    def test_dates_either_side_of_the_hole_round_trip(self, tmp_path, day):
        got = _read_back(tmp_path, day)
        assert got == datetime.datetime(day.year, day.month, day.day)

    def test_extract_reports_the_loss_instead_of_inventing_a_date(self, tmp_path):
        """What a user actually sees. The value is wrong (it is a time), and the
        warning now names the cause — it says the value IS a time, which is the
        clue that leads to the storage limit rather than to their pattern."""
        pattern = str(tmp_path / 'p.csv')
        data = str(tmp_path / 'd.xlsx')
        write_pattern_csv(pattern, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', 'date', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ])
        write_data_xlsx(data, {
            (1, 1): ('ANCHOR', None),
            (1, 2): (datetime.date(1899, 12, 31), 'yyyy-mm-dd'),
        })
        lg = Logger(level=VerbosityLevel.QUIET)
        result = Engine().process(pattern_file=pattern, data_file=data, logger=lg)
        issues = [r for r in lg._records
                  if r.severity in (Severity.WARNING, Severity.ERROR)]

        assert result['f'] == datetime.time(0, 0)
        assert issues, 'a lost date produced no warning at all'
        assert any('is not a date' in str(r.message) for r in issues), \
            [str(r.message) for r in issues]
        assert any('time' in str(r.message) for r in issues), (
            'the warning should name the value it actually got, so the user can '
            'see it became a time: ' + str([str(r.message) for r in issues])
        )


class TestSerialFractionIsTimeOfDay:
    """The other half of the representation: the fraction is the time of day,
    and 1.0 is one whole day. This is what makes elapsed time work."""

    @pytest.mark.parametrize('serial,expected', [
        (45000.0, datetime.datetime(2023, 3, 15, 0, 0)),
        (45000.5, datetime.datetime(2023, 3, 15, 12, 0)),
        (45000.25, datetime.datetime(2023, 3, 15, 6, 0)),
        (45000.75, datetime.datetime(2023, 3, 15, 18, 0)),
    ])
    def test_fraction_is_the_time_of_day(self, tmp_path, serial, expected):
        got = _read_back(tmp_path, serial, fmt='yyyy-mm-dd hh:mm')
        assert got == expected

    @pytest.mark.parametrize('serial,seconds', [
        (1 / 86400, 1),
        (11 / 86400, 11),
        (112 / 86400, 112),
        (1133 / 86400, 1133),
        (0.5, 43200),
        (1.0, 86400),
        (1.01394676, 87605),     # 24:20:05 — past one day, still counting hours
    ])
    def test_serial_times_86400_is_the_duration_in_seconds(self, tmp_path,
                                                           serial, seconds):
        """serial x 86400 == seconds, exactly as the reference table shows."""
        got = _read_back(tmp_path, serial, fmt='[h]:mm:ss')
        assert isinstance(got, datetime.timedelta)
        assert round(got.total_seconds()) == seconds
