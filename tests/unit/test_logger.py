import json

import pytest
from grepxcel.logger import (
    Logger, LogRecord, EngineError, VerbosityLevel,
    Severity, Category, col_letter, cell_ref, LOG_SCHEMA_VERSION,
)


# ─── structured (NDJSON) logging + redaction + correlation ───────────────────

def _json_log_lines(tmp_path, **kw):
    """Write one validation warning to a json log and return the parsed lines."""
    log_path = tmp_path / 'run.jsonl'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='json', source='data.xlsx', **kw)
    rec = lg.warn_validation(10, 2, 'client.name', 'string', r'.+', 'Alice Wonderland')
    lg.commit_warnings([rec])
    lg.close()
    return [json.loads(ln) for ln in
            log_path.read_text(encoding='utf-8').splitlines() if ln.strip()]


def test_json_log_emits_ndjson(tmp_path):
    lines = _json_log_lines(tmp_path)
    assert len(lines) == 1
    rec = lines[0]
    assert rec['severity'] == 'WARNING'
    assert rec['field'] == 'client.name'


def test_json_log_has_correlation_and_schema(tmp_path):
    rec = _json_log_lines(tmp_path)[0]
    assert rec['source'] == 'data.xlsx'
    assert rec['schema_version'] == LOG_SCHEMA_VERSION
    assert rec['level'] == 'WARNING'
    assert rec['run_id']            # present + non-empty


def test_json_log_redacts_cell_value_by_default(tmp_path):
    rec = _json_log_lines(tmp_path)[0]
    assert 'Alice Wonderland' not in json.dumps(rec)
    assert rec['found'] == '<redacted>'


def test_json_log_raw_includes_value(tmp_path):
    rec = _json_log_lines(tmp_path, redact=False)[0]
    assert 'Alice Wonderland' in rec['found']


def test_run_id_is_stable_within_a_run(tmp_path):
    lines = []
    log_path = tmp_path / 'r.jsonl'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='json', source='d.xlsx')
    lg.commit_warnings([lg.warn_validation(1, 1, 'a', 'string', '.+', 'x')])
    lg.commit_warnings([lg.warn_validation(2, 1, 'b', 'string', '.+', 'y')])
    lg.close()
    ids = {json.loads(ln)['run_id']
           for ln in log_path.read_text().splitlines() if ln.strip()}
    assert len(ids) == 1            # same run_id across the run


def test_text_log_unchanged_includes_values(tmp_path):
    """Text mode still mirrors the console (values shown) — for human use."""
    log_path = tmp_path / 'run.log'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path), source='d.xlsx')
    lg.commit_warnings([lg.warn_validation(1, 1, 'a', 'string', '.+', 'SECRET')])
    lg.close()
    assert 'SECRET' in log_path.read_text(encoding='utf-8')


# ─── log file is appended, not truncated ─────────────────────────────────────

def test_log_file_appends_not_truncates(tmp_path):
    """--log must not silently truncate an existing file (it's documented as
    'append')."""
    log_path = tmp_path / 'run.log'
    log_path.write_text('PRE-EXISTING CONTENT\n', encoding='utf-8')
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path))
    lg.close()
    assert 'PRE-EXISTING CONTENT' in log_path.read_text(encoding='utf-8')


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


# ─── summary issues recap ─────────────────────────────────────────────────────

def test_summary_lists_each_issue_cell(capsys):
    lg = Logger(level=VerbosityLevel.NORMAL, sheet_name='Sheet1')
    rec = lg.warn_validation(10, 2, 'po.number', 'string', r'PO-\d+', 'xyz')
    lg.commit_warnings([rec])
    lg.summary({'cells': {'po.number': 'xyz'}, 'tables': []})
    out = capsys.readouterr().err
    assert 'ISSUES' in out
    assert 'Sheet1!B10' in out      # the failing cell
    assert 'po.number' in out       # the field
    assert "'xyz'" in out           # what was found


def test_summary_no_issues_section_when_clean(capsys):
    lg = Logger(level=VerbosityLevel.NORMAL)
    lg.summary({'cells': {'a': 1}, 'tables': []})
    out = capsys.readouterr().err
    assert 'ISSUES' not in out


def test_issue_line_marks_error_vs_warning():
    lg = Logger(level=VerbosityLevel.QUIET)
    warn = lg.warn_validation(1, 1, 'f', 'string', '.*', 'v')
    line = lg._issue_line(warn)
    assert line.startswith('⚠')


def test_summary_scope_excludes_prior_sheet_issues(capsys):
    lg = Logger(level=VerbosityLevel.NORMAL, sheet_name='Sheet1')
    # Sheet 1: one warning
    lg.begin_summary_scope()
    lg.commit_warnings([lg.warn_validation(1, 1, 'a', 'string', r'x', 'bad')])
    lg.summary({'cells': {'a': 'bad'}, 'tables': []})
    capsys.readouterr()  # discard sheet-1 output

    # Sheet 2: clean — must NOT inherit sheet 1's warning
    lg.begin_summary_scope()
    lg.summary({'cells': {'b': 1}, 'tables': []})
    out = capsys.readouterr().err
    assert 'Warnings          : 0' in out
    assert 'ISSUES' not in out
    assert "'bad'" not in out


# ─── per-field extraction trace (VERBOSE) ─────────────────────────────────────

def test_cell_trace_shows_pass_mark(capsys):
    lg = Logger(level=VerbosityLevel.VERBOSE, sheet_name='Sheet1')
    lg.cell_processed(1, 2, 'po.number', 'PO-2026', ok=True, regex=r'PO-\d+')
    out = capsys.readouterr().err
    assert 'po.number' in out
    assert 'Sheet1!B1' in out
    assert '←' in out
    assert '✓' in out


def test_cell_trace_shows_fail_mark_and_regex(capsys):
    lg = Logger(level=VerbosityLevel.VERBOSE, sheet_name='Sheet1')
    lg.cell_processed(1, 1, 'code', 'bad', ok=False, regex=r'[A-Z]{3}')
    out = capsys.readouterr().err
    assert '✗' in out
    assert 'does not match /[A-Z]{3}/' in out


def test_cell_trace_no_mark_when_ok_none(capsys):
    # An empty optional field (ok=None) → traced without a ✓/✗ mark.
    lg = Logger(level=VerbosityLevel.VERBOSE)
    lg.cell_processed(1, 1, 'x', None, ok=None)
    out = capsys.readouterr().err
    assert '✓' not in out and '✗' not in out


def test_trace_suppressed_below_verbose(capsys):
    lg = Logger(level=VerbosityLevel.NORMAL)
    lg.cell_processed(1, 1, 'x', 'v', ok=True)
    out = capsys.readouterr().err
    assert out == ''            # NORMAL must not print the per-field trace


def test_commit_traces_emits_only_at_verbose(capsys):
    line = '  [FIELD] f ← A1 = 1 ✓'
    quiet = Logger(level=VerbosityLevel.QUIET)
    quiet.commit_traces([line])
    assert capsys.readouterr().err == ''

    verbose = Logger(level=VerbosityLevel.VERBOSE)
    verbose.commit_traces([line])
    assert line in capsys.readouterr().err


def test_trace_field_returns_line_without_emitting(capsys):
    lg = Logger(level=VerbosityLevel.VERBOSE, sheet_name='S')
    line = lg.trace_field(4, 1, 'row.item', 'Laptop', ok=True)
    assert capsys.readouterr().err == ''     # building a trace does not print
    assert 'row.item' in line and 'S!A4' in line and '✓' in line
