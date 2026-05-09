import pytest
from engine.logger import (
    Logger, LogRecord, EngineError, VerbosityLevel,
    Severity, Category, col_letter, cell_ref,
)


# ─── cell reference helpers ──────────────────────────────────────────────────

def test_col_letter_single():
    assert col_letter(1) == 'A'
    assert col_letter(26) == 'Z'

def test_col_letter_double():
    assert col_letter(27) == 'AA'
    assert col_letter(52) == 'AZ'
    assert col_letter(53) == 'BA'

def test_cell_ref_no_sheet():
    assert cell_ref(1, 1) == 'A1'
    assert cell_ref(10, 2) == 'B10'

def test_cell_ref_with_sheet():
    assert cell_ref(10, 2, 'Sheet1') == 'Sheet1!B10'


# ─── LogRecord ───────────────────────────────────────────────────────────────

def test_log_record_required_fields():
    r = LogRecord(severity=Severity.INFO, category=Category.ENGINE, message='hello')
    assert r.severity == Severity.INFO
    assert r.category == Category.ENGINE
    assert r.message == 'hello'

def test_log_record_optional_fields_default_empty():
    r = LogRecord(Severity.INFO, Category.ENGINE, 'msg')
    assert r.location == ''
    assert r.field == ''
    assert r.field_type == ''
    assert r.expected == ''
    assert r.found == ''
    assert r.hint == ''

def test_log_record_timestamp_set():
    r = LogRecord(Severity.INFO, Category.ENGINE, 'msg')
    assert r.timestamp  # non-empty ISO string
    assert 'T' in r.timestamp

def test_log_record_to_dict_keys():
    r = LogRecord(Severity.WARNING, Category.VALIDATION, 'bad value',
                  location='A1', field='qty', found='x')
    d = r.to_dict()
    assert set(d.keys()) == {'severity', 'category', 'message', 'location',
                              'field', 'field_type', 'expected', 'found',
                              'hint', 'timestamp'}
    assert d['location'] == 'A1'
    assert d['field'] == 'qty'


# ─── Logger: basic record collection ─────────────────────────────────────────

def test_logger_starts_empty():
    lg = Logger(level=VerbosityLevel.QUIET)
    assert lg.records() == []

def test_logger_has_errors_false():
    lg = Logger(level=VerbosityLevel.QUIET)
    assert not lg.has_errors()

def test_logger_issues_empty():
    lg = Logger(level=VerbosityLevel.QUIET)
    assert lg.issues() == []


# ─── warn_validation / warn_empty_field / commit_warnings ────────────────────

def test_warn_validation_not_stored_until_commit(capsys):
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'qty', 'integer', r'[0-9]+', 'bad')
    assert isinstance(rec, LogRecord)
    assert rec.severity == Severity.WARNING
    assert len(lg.records()) == 0  # not stored yet

def test_warn_validation_stored_after_commit():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'qty', 'integer', r'[0-9]+', 'bad')
    lg.commit_warnings([rec])
    assert len(lg.records()) == 1
    assert lg.has_errors() is False
    assert len(lg.issues()) == 1

def test_warn_empty_field_not_stored_until_commit():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_empty_field(2, 3, 'amount', 'currency')
    assert isinstance(rec, LogRecord)
    assert len(lg.records()) == 0

def test_warn_empty_field_stored_after_commit():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_empty_field(2, 3, 'amount', 'currency')
    lg.commit_warnings([rec])
    issues = lg.issues()
    assert len(issues) == 1
    assert issues[0].field == 'amount'
    assert 'empty' in issues[0].message.lower()

def test_warn_undefined_field_stored_after_commit():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_undefined_field(1, 1, 'ghost.field')
    lg.commit_warnings([rec])
    assert lg.issues()[0].field == 'ghost.field'

def test_commit_multiple_warnings():
    lg = Logger(level=VerbosityLevel.QUIET)
    recs = [
        lg.warn_validation(1, 1, 'a', 'string', '.*', 'x'),
        lg.warn_empty_field(2, 2, 'b', 'integer'),
    ]
    lg.commit_warnings(recs)
    assert len(lg.records()) == 2


# ─── issues() filtering ──────────────────────────────────────────────────────

def test_issues_only_warning_and_error():
    lg = Logger(level=VerbosityLevel.QUIET)
    lg._emit(Severity.INFO, Category.ENGINE, 'info msg')
    rec = lg.warn_validation(1, 1, 'f', 'string', '.*', 'v')
    lg.commit_warnings([rec])
    issues = lg.issues()
    assert len(issues) == 1
    assert issues[0].severity == Severity.WARNING


# ─── fatal raises EngineError and stores ERROR record ────────────────────────

def test_fatal_raises_engine_error():
    lg = Logger(level=VerbosityLevel.QUIET)
    with pytest.raises(EngineError) as exc_info:
        lg.fatal('something broke')
    assert 'something broke' in str(exc_info.value)

def test_fatal_stores_error_record():
    lg = Logger(level=VerbosityLevel.QUIET)
    with pytest.raises(EngineError):
        lg.fatal('boom', location='B5', expected='number', found='text')
    assert lg.has_errors()
    errors = [r for r in lg.records() if r.severity == Severity.ERROR]
    assert len(errors) == 1
    assert errors[0].location == 'B5'


# ─── hint logic ──────────────────────────────────────────────────────────────

def test_hint_newline_in_value():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'label', 'string', 'Plant St. Loc.', 'Plant\nSt. Loc.')
    assert 'newline' in rec.hint.lower()

def test_hint_float_precision():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'price', 'currency', r'[0-9]+\.[0-9]{2}',
                             1.3800000000000001)
    assert 'precision' in rec.hint.lower() or 'float' in rec.hint.lower()

def test_hint_integer_mismatch():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'code', 'integer', r'[1-9][0-9]{3}', 5)
    assert rec.hint  # should have some hint text

def test_hint_empty_for_unrecognised():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'name', 'string', r'[A-Z]+', 'abc')
    assert rec.hint == ''


# ─── to_dict (JSON-serialisable) ─────────────────────────────────────────────

def test_to_dict_is_list_of_dicts():
    lg = Logger(level=VerbosityLevel.QUIET)
    rec = lg.warn_validation(1, 1, 'f', 'string', '.*', 'v')
    lg.commit_warnings([rec])
    result = lg.to_dict()
    assert isinstance(result, list)
    assert isinstance(result[0], dict)
