"""Unit tests for grepxcel.cell_taxonomy — the canonical cell classifier."""
import datetime

from grepxcel.cell_taxonomy import CellProfile, classify_cell, classify_value


class _FakeCell:
    """Minimal stand-in for an openpyxl Cell — just .value/.number_format."""
    def __init__(self, value, number_format='General'):
        self.value = value
        self.number_format = number_format


# ── Layer 1/2/3: plain values, no format info ──────────────────────────────

class TestClassifyValueBasic:
    def test_none_is_empty(self):
        p = classify_value(None)
        assert p.semantic_type == 'empty'
        assert 'blank' in p.flags
        assert p.python_type == 'NoneType'

    def test_plain_string(self):
        p = classify_value('hello')
        assert p.semantic_type == 'string'
        assert p.storage_type == 's'
        assert p.python_type == 'str'

    def test_bool_true(self):
        p = classify_value(True)
        assert p.semantic_type == 'boolean'
        assert p.storage_type == 'b'
        assert p.python_type == 'bool'

    def test_bool_string(self):
        for s in ('TRUE', 'false', 'Yes', 'no'):
            assert classify_value(s).semantic_type == 'boolean'

    def test_integer(self):
        p = classify_value(42)
        assert p.semantic_type == 'integer'
        assert p.python_type == 'int'

    def test_float_whole_number_is_integer(self):
        p = classify_value(42.0)
        assert p.semantic_type == 'integer'

    def test_float_non_whole_is_number_with_no_format(self):
        p = classify_value(42.5)
        assert p.semantic_type == 'number'

    def test_url_string(self):
        assert classify_value('https://example.com').semantic_type == 'url'
        assert classify_value('www.example.com').semantic_type == 'url'

    def test_error_string(self):
        p = classify_value('#DIV/0!')
        assert p.semantic_type == 'error'
        assert p.storage_type == 'e'
        assert 'error' in p.flags

    def test_all_nine_error_codes(self):
        codes = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!',
                 '#VALUE!', '#SPILL!', '#CALC!']
        for code in codes:
            assert classify_value(code).semantic_type == 'error', code


# ── Layer 2: number-format-driven semantic subtype ──────────────────────────

class TestNumberFormatDriven:
    def test_date_format(self):
        p = classify_value(45601.0, 'yyyy-mm-dd')
        assert p.semantic_type == 'date'

    def test_percentage_format(self):
        p = classify_value(0.15, '0.00%')
        assert p.semantic_type == 'percentage'

    def test_currency_format_dollar(self):
        p = classify_value(19.99, '$#,##0.00')
        assert p.semantic_type == 'currency'

    def test_currency_format_euro(self):
        p = classify_value(9.99, '€#,##0.00')
        assert p.semantic_type == 'currency'

    def test_plain_number_no_special_format(self):
        p = classify_value(42.5, 'General')
        assert p.semantic_type == 'number'

    def test_real_date_object_ignores_format_string(self):
        # An actual datetime.date always classifies as 'date' regardless
        # of number_format — the Python type itself is unambiguous.
        p = classify_value(datetime.date(2024, 1, 1), 'General')
        assert p.semantic_type == 'date'


# ── Layer 3: date/time/datetime/duration python objects ────────────────────

class TestDatetimeFamily:
    def test_date(self):
        p = classify_value(datetime.date(2024, 1, 15))
        assert p.semantic_type == 'date'
        assert p.python_type == 'date'
        assert p.raw_repr == '2024-01-15'

    def test_datetime(self):
        p = classify_value(datetime.datetime(2024, 1, 15, 9, 30))
        assert p.semantic_type == 'datetime'
        assert p.python_type == 'datetime'

    def test_time(self):
        p = classify_value(datetime.time(9, 30))
        assert p.semantic_type == 'time'
        assert p.python_type == 'time'

    def test_timedelta_is_duration(self):
        p = classify_value(datetime.timedelta(hours=2, minutes=30))
        assert p.semantic_type == 'duration'
        assert p.python_type == 'timedelta'

    def test_date_and_time_are_distinct_semantics(self):
        # Unlike wizard_core's legacy collapsing of datetime->date and
        # timedelta->time, the canonical classifier keeps all four distinct.
        d = classify_value(datetime.date(2024, 1, 1)).semantic_type
        dt = classify_value(datetime.datetime(2024, 1, 1)).semantic_type
        t = classify_value(datetime.time(1, 0)).semantic_type
        td = classify_value(datetime.timedelta(hours=1)).semantic_type
        assert {d, dt, t, td} == {'date', 'datetime', 'time', 'duration'}


# ── Layer 4: flags ───────────────────────────────────────────────────────────

class TestFlags:
    def test_text_forced_numeric(self):
        p = classify_value('00123')
        assert p.semantic_type == 'string'
        assert 'text_forced_numeric' in p.flags

    def test_text_forced_numeric_float(self):
        p = classify_value('3.14')
        assert 'text_forced_numeric' in p.flags

    def test_plain_text_has_no_numeric_flag(self):
        p = classify_value('hello world')
        assert 'text_forced_numeric' not in p.flags

    def test_formula_flag(self):
        p = classify_value('some cached result', is_formula=True)
        assert 'formula' in p.flags

    def test_formula_string_storage_type(self):
        p = classify_value('cached', is_formula=True)
        assert p.storage_type == 'str'

    def test_blank_flag(self):
        p = classify_value(None)
        assert 'blank' in p.flags

    def test_error_and_formula_can_combine(self):
        # A formula that errored: cached value is the error text.
        p = classify_value('#DIV/0!', is_formula=True)
        assert 'error' in p.flags
        assert 'formula' in p.flags


# ── classify_cell: live openpyxl-shaped object ──────────────────────────────

class TestClassifyCell:
    def test_reads_value_and_format(self):
        cell = _FakeCell(19.99, '$#,##0.00')
        p = classify_cell(cell)
        assert p.semantic_type == 'currency'

    def test_formula_detection_from_leading_equals(self):
        cell = _FakeCell('=A1+B1')
        p = classify_cell(cell)
        assert 'formula' in p.flags

    def test_default_number_format_when_missing(self):
        class _Bare:
            value = 42
        p = classify_cell(_Bare())
        assert p.number_format == 'General'
        assert p.semantic_type == 'integer'


# ── CellProfile shape ────────────────────────────────────────────────────────

class TestCellProfileShape:
    def test_is_a_dataclass_with_expected_fields(self):
        p = classify_value(1)
        assert isinstance(p, CellProfile)
        assert hasattr(p, 'storage_type')
        assert hasattr(p, 'semantic_type')
        assert hasattr(p, 'python_type')
        assert hasattr(p, 'flags')
        assert hasattr(p, 'raw_repr')
        assert hasattr(p, 'number_format')

    def test_flags_default_to_empty_list_not_shared(self):
        # Dataclass mutable-default footgun check: two independent calls
        # must not share the same underlying list object.
        p1 = classify_value(1)
        p2 = classify_value(2)
        p1.flags.append('x')
        assert p2.flags == []
