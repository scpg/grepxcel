"""Text cells declared as date/time are converted, unambiguously or not at all.

Three layers:
  1. `utils.parse_duration_text` / `utils.coerce_temporal_text` — the conversion.
  2. `config: | date.format | ...` / `time.format` — parsing and validation.
  3. end to end through `Engine().process`, which is what a user sees.

The negative tests carry the weight here. The whole design rests on refusing to
guess: '01/02/2024' is both 1 Feb and 2 Jan, so without a declared format it must
stay rejected rather than silently becoming one of them.
"""
import datetime

import pytest

from grepxcel.engine import Engine
from grepxcel.logger import Logger, Severity, VerbosityLevel
from grepxcel.models import Config
from grepxcel.pattern_parser import PatternError, PatternParser
from grepxcel.utils import (TEMPORAL_TYPES, coerce_temporal_text,
                            parse_duration_text, validate_type)
from tests.unit.test_random_type_fixtures import write_data_xlsx, write_pattern_csv


# ── 1. parse_duration_text ────────────────────────────────────────────────
class TestParseDurationText:
    @pytest.mark.parametrize('text,expected', [
        ('0:00', datetime.timedelta(0)),
        ('2:00', datetime.timedelta(hours=2)),
        ('2:30', datetime.timedelta(hours=2, minutes=30)),
        # Hours are unbounded on purpose: elapsed time, not a clock.
        ('30:00', datetime.timedelta(hours=30)),
        ('999:59', datetime.timedelta(hours=999, minutes=59)),
        ('2:00:30', datetime.timedelta(hours=2, seconds=30)),
        ('  2:00  ', datetime.timedelta(hours=2)),
    ])
    def test_parses(self, text, expected):
        assert parse_duration_text(text) == expected

    @pytest.mark.parametrize('text', [
        'banana', '2', '2:', ':30', '2:60', '2:00:60', '2.00', '-2:00', '',
        '2:00:30:45',
    ])
    def test_rejects(self, text):
        assert parse_duration_text(text) is None


# ── 2. coerce_temporal_text ───────────────────────────────────────────────
class TestCoerceTemporalText:
    @pytest.mark.parametrize('field_type,text,expected', [
        ('date', '2024-01-15', datetime.datetime(2024, 1, 15)),
        ('date', '1900-01-01', datetime.datetime(1900, 1, 1)),
        ('date', '9999-12-31', datetime.datetime(9999, 12, 31)),
        ('datetime', '2024-01-15 09:30', datetime.datetime(2024, 1, 15, 9, 30)),
        ('datetime', '2024-01-15T09:30:00', datetime.datetime(2024, 1, 15, 9, 30)),
        ('timestamp', '2024-01-15', datetime.datetime(2024, 1, 15)),
        ('time', '09:30', datetime.time(9, 30)),
        ('time', '00:00', datetime.time(0, 0)),
        ('time', '23:59:59', datetime.time(23, 59, 59)),
        ('duration', '2:00', datetime.timedelta(hours=2)),
        ('duration', '30:00', datetime.timedelta(hours=30)),
    ])
    def test_iso_and_elapsed_convert(self, field_type, text, expected):
        assert coerce_temporal_text(text, field_type) == expected

    def test_a_date_field_yields_datetime_not_date(self):
        """Matches what openpyxl returns for a real Excel date cell, so a text
        date and a native date produce the same output shape."""
        got = coerce_temporal_text('2024-01-15', 'date')
        assert type(got) is datetime.datetime
        assert got == datetime.datetime(2024, 1, 15, 0, 0)

    @pytest.mark.parametrize('text', ['01/02/2024', '15/01/2024', 'Jan 15, 2024',
                                      '15-Jan-2024', '2024.01.15'])
    def test_ambiguous_or_non_iso_is_refused_without_a_declared_format(self, text):
        """The core safety property: no guessing. '01/02/2024' is 1 Feb in most
        of the world and 2 Jan in the US; picking one silently is how a tool
        returns confidently wrong data."""
        assert coerce_temporal_text(text, 'date') is None

    def test_declared_format_resolves_the_ambiguity(self):
        assert coerce_temporal_text('01/02/2024', 'date',
                                    date_format='%d/%m/%Y') == datetime.datetime(2024, 2, 1)
        assert coerce_temporal_text('01/02/2024', 'date',
                                    date_format='%m/%d/%Y') == datetime.datetime(2024, 1, 2)

    def test_declared_format_falls_back_to_iso(self):
        """A sheet may mix local and ISO text; declaring a format must not turn
        ISO cells into failures."""
        assert coerce_temporal_text('2024-01-15', 'date',
                                    date_format='%d/%m/%Y') == datetime.datetime(2024, 1, 15)

    def test_declared_time_format(self):
        assert coerce_temporal_text('09.30', 'time',
                                    time_format='%H.%M') == datetime.time(9, 30)

    @pytest.mark.parametrize('field_type', ['string', 'integer', 'number',
                                            'currency', 'boolean', 'url'])
    def test_non_temporal_types_are_untouched(self, field_type):
        assert coerce_temporal_text('2024-01-15', field_type) is None

    @pytest.mark.parametrize('value', [42, 42.5, True, None,
                                       datetime.date(2024, 1, 15)])
    def test_non_text_values_are_untouched(self, value):
        assert coerce_temporal_text(value, 'date') is None

    @pytest.mark.parametrize('text', ['', '   '])
    def test_empty_text_is_not_coerced(self, text):
        assert coerce_temporal_text(text, 'date') is None

    def test_temporal_types_matches_validate_type_grouping(self):
        """Cross-reference guard: every type coerced here must be one
        validate_type actually treats as temporal, or the conversion would
        produce a value its own validator then rejects."""
        for field_type in TEMPORAL_TYPES:
            converted = coerce_temporal_text('2024-01-15 09:30', field_type) \
                or coerce_temporal_text('09:30', field_type)
            assert converted is not None, field_type
            ok, reason = validate_type(converted, field_type, '.*')
            assert ok, f'{field_type}: coerced value rejected — {reason}'


# ── 3. config: date.format / time.format ──────────────────────────────────
def _write_pattern(tmp_path, rows):
    path = str(tmp_path / 'pattern.csv')
    write_pattern_csv(path, rows)
    return path


class TestConfigKeys:
    def test_date_and_time_format_are_parsed(self, tmp_path):
        path = _write_pattern(tmp_path, [
            ['config:', 'date.format', '%d/%m/%Y'],
            ['config:', 'time.format', '%H.%M'],
            ['var:', 'd', 'date', '.*'],
            [], ['START:'], ['cell:A1', 'd'], ['END:'],
        ])
        config, _defs, _seq = PatternParser().parse(path)
        assert config.date_format == '%d/%m/%Y'
        assert config.time_format == '%H.%M'

    def test_default_is_none(self, tmp_path):
        path = _write_pattern(tmp_path, [
            ['var:', 'd', 'date', '.*'],
            [], ['START:'], ['cell:A1', 'd'], ['END:'],
        ])
        config, _defs, _seq = PatternParser().parse(path)
        assert config.date_format is None and config.time_format is None

    def test_an_unusable_format_fails_at_parse_time(self, tmp_path):
        """Reported against the pattern, not as a wall of per-cell type
        mismatches that would point the user at their data instead."""
        path = _write_pattern(tmp_path, [
            ['config:', 'date.format', '%Q/%Z/%'],
            ['var:', 'd', 'date', '.*'],
            [], ['START:'], ['cell:A1', 'd'], ['END:'],
        ])
        with pytest.raises(PatternError, match='date.format'):
            PatternParser().parse(path)

    def test_the_key_is_recognised(self, tmp_path):
        """Not silently collected as an unknown config key."""
        path = _write_pattern(tmp_path, [
            ['config:', 'date.format', '%d/%m/%Y'],
            ['var:', 'd', 'date', '.*'],
            [], ['START:'], ['cell:A1', 'd'], ['END:'],
        ])
        config, _defs, _seq = PatternParser().parse(path)
        assert 'date.format' not in config.unknown_config_keys


# ── 4. end to end ─────────────────────────────────────────────────────────
def _run(tmp_path, rows, cells):
    pattern = _write_pattern(tmp_path, rows)
    data = str(tmp_path / 'data.xlsx')
    write_data_xlsx(data, cells)
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pattern, data_file=data, logger=lg)
    issues = [r for r in lg._records
              if r.severity in (Severity.WARNING, Severity.ERROR)]
    return result, issues


class TestEndToEnd:
    @pytest.mark.parametrize('field_type,text,expected', [
        ('date', '2024-01-15', datetime.datetime(2024, 1, 15)),
        ('datetime', '2024-01-15 09:30', datetime.datetime(2024, 1, 15, 9, 30)),
        ('time', '09:30', datetime.time(9, 30)),
        ('duration', '2:00', datetime.timedelta(hours=2)),
        ('duration', '30:00', datetime.timedelta(hours=30)),
    ])
    def test_text_cell_extracts_as_a_real_temporal_value(self, tmp_path,
                                                        field_type, text, expected):
        result, issues = _run(tmp_path, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', field_type, '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ], {(1, 1): ('ANCHOR', None), (1, 2): (text, None)})
        assert result['f'] == expected, f'got {result["f"]!r}'
        assert not issues, [r.message for r in issues]

    def test_ambiguous_text_is_still_reported_with_its_reason(self, tmp_path):
        result, issues = _run(tmp_path, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', 'date', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ], {(1, 1): ('ANCHOR', None), (1, 2): ('01/02/2024', None)})
        # Unconverted, so the raw text comes through and is reported as before.
        assert result['f'] == '01/02/2024'
        assert any('is not a date' in str(r.message) for r in issues), \
            [str(r.message) for r in issues]

    def test_declared_format_makes_it_extract(self, tmp_path):
        result, issues = _run(tmp_path, [
            ['config:', 'date.format', '%d/%m/%Y'],
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', 'date', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ], {(1, 1): ('ANCHOR', None), (1, 2): ('01/02/2024', None)})
        assert result['f'] == datetime.datetime(2024, 2, 1)
        assert not issues, [r.message for r in issues]

    def test_coercion_works_inside_a_table(self, tmp_path):
        result, issues = _run(tmp_path, [
            ['lbl:', 'h', 'string', 'COLHEAD'],
            ['var:', 'items.d', 'date', '.*'],
            [], ['START:'], ['table:*'],
            ['', 'HEADER:1', 'h'], ['', 'DATA:*', 'items.d'], ['END:'],
        ], {
            (1, 1): ('COLHEAD', None),
            (2, 1): ('2024-01-15', None),
            (3, 1): ('2024-02-20', None),
        })
        rows = result['items'][0]['data']
        assert [r['d'] for r in rows] == [datetime.datetime(2024, 1, 15),
                                          datetime.datetime(2024, 2, 20)]
        assert not issues, [r.message for r in issues]

    def test_a_label_field_is_never_rewritten(self, tmp_path):
        """lbl: fields are anchors matched as text; coercing one would break the
        match it exists to perform."""
        result, issues = _run(tmp_path, [
            ['lbl:', 'a', 'date', '2024-01-15'],
            ['var:', 'f', 'string', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ], {(1, 1): ('2024-01-15', None), (1, 2): ('value', None)})
        assert result['f'] == 'value'
        assert not issues, [r.message for r in issues]

    def test_string_field_holding_a_date_stays_text(self, tmp_path):
        """Only the declared type decides. A string field keeps its string."""
        result, _ = _run(tmp_path, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', 'string', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ], {(1, 1): ('ANCHOR', None), (1, 2): ('2024-01-15', None)})
        assert result['f'] == '2024-01-15'

    def test_native_excel_date_is_unaffected(self, tmp_path):
        """Regression guard: the coercion path must not touch real date cells."""
        result, issues = _run(tmp_path, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', 'date', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ], {(1, 1): ('ANCHOR', None),
            (1, 2): (datetime.date(2024, 1, 15), 'yyyy-mm-dd')})
        assert result['f'] == datetime.datetime(2024, 1, 15)
        assert not issues

    def test_coercion_is_recorded_for_debug_output(self, tmp_path):
        """A computed conversion that reaches no output path is invisible: the
        extracted value would differ from the cell with nothing to explain it."""
        pattern = _write_pattern(tmp_path, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'f', 'date', '.*'],
            [], ['START:'], ['cell:next', 'a'], ['cell:next', 'f'], ['END:'],
        ])
        data = str(tmp_path / 'data.xlsx')
        write_data_xlsx(data, {(1, 1): ('ANCHOR', None), (1, 2): ('2024-01-15', None)})
        lg = Logger(level=VerbosityLevel.DEBUG)
        Engine().process(pattern_file=pattern, data_file=data, logger=lg)
        assert any('Coerced text' in str(r.message) for r in lg._records), \
            [str(r.message) for r in lg._records]
