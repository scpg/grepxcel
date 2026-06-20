import json

import pytest
from grepxcel.logger import (
    Logger, LogRecord, EngineError, VerbosityLevel,
    Severity, Category, col_letter, cell_ref, LOG_SCHEMA_VERSION,
    _SAFE_LOG_KEYS, _SAFE_SUMMARY_KEYS,
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
    assert rec['level'] == 'WARNING'
    assert rec['field'] == 'client.name'
    assert 'found' not in rec
    assert 'hint' not in rec
    assert 'message' not in rec
    assert 'expected' not in rec


def test_json_log_has_correlation_and_schema(tmp_path):
    rec = _json_log_lines(tmp_path)[0]
    assert rec['source'] == 'data.xlsx'
    assert rec['schema_version'] == LOG_SCHEMA_VERSION
    assert rec['level'] == 'WARNING'
    assert rec['run_id']            # present + non-empty


def test_json_log_contains_no_cell_values(tmp_path):
    """Structured JSON logs never contain extracted cell values — by construction
    (allow-list), not redaction.  A non-reversible fingerprint is included instead."""
    rec = _json_log_lines(tmp_path)[0]
    raw = json.dumps(rec)
    assert 'Alice Wonderland' not in raw
    assert 'found' not in rec
    assert 'hint' not in rec
    assert 'expected' not in rec
    assert 'message' not in rec
    assert rec['value_len'] == len('Alice Wonderland')
    assert isinstance(rec['value_sha8'], str) and len(rec['value_sha8']) == 8
    assert set(rec.keys()).issubset(set(_SAFE_LOG_KEYS))


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
                              'hint', 'event', 'value_len', 'value_sha8',
                              'timestamp'}
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


# ─── #34 sentinel: prove no Excel cell data appears in structured logs ────────

def test_sentinel_no_cell_data_in_json_log(tmp_path):
    """End-to-end: extract fixture 01, write NDJSON log, assert none of the
    extracted cell values appear anywhere in the log bytes."""
    import grepxcel

    fixture = 'tests/fixtures/01_simple_invoice'
    pattern = f'{fixture}/pattern-from-draft.xlsx'
    data = f'{fixture}/data.xlsx'
    log_path = tmp_path / 'sentinel.jsonl'

    lg = Logger(level=VerbosityLevel.VERBOSE, log_file=str(log_path),
                log_format='json', source='data.xlsx')
    result = grepxcel.extract(pattern, data, logger=lg)
    lg.summary(result)
    lg.close()

    log_bytes = log_path.read_text(encoding='utf-8')

    # Collect every scalar value the extraction produced.
    sentinels = []
    def _collect(obj):
        if isinstance(obj, dict):
            for v in obj.values():
                _collect(v)
        elif isinstance(obj, list):
            for v in obj:
                _collect(v)
        elif obj is not None:
            sentinels.append(str(obj))

    _collect(result)
    assert sentinels, 'fixture must produce at least some values'

    for val in sentinels:
        if len(val) < 3:
            continue  # skip trivially short values (e.g. single digits)
        assert val not in log_bytes, (
            f'Extracted cell value leaked into structured log: {val!r}'
        )


def test_json_record_keys_pinned_to_allow_list(tmp_path):
    """Every key in every NDJSON record must belong to the appropriate allow-list:
    _SAFE_LOG_KEYS for per-cell records, _SAFE_SUMMARY_KEYS for the summary event."""
    log_path = tmp_path / 'pin.jsonl'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='json', source='x.xlsx')
    lg.commit_warnings([lg.warn_validation(1, 1, 'f', 'string', '.+', 'secret')])
    lg.commit_warnings([lg.warn_empty_field(2, 2, 'g', 'integer')])
    lg.commit_warnings([lg.warn_undefined_field(3, 3, 'h')])
    lg.summary({'cells': {'f': 'secret', 'g': ''}, 'tables': []})
    lg.close()

    for line in log_path.read_text(encoding='utf-8').splitlines():
        if not line.strip():
            continue
        rec = json.loads(line)
        if rec.get('event') == 'summary':
            allowed = set(_SAFE_SUMMARY_KEYS)
        else:
            allowed = set(_SAFE_LOG_KEYS)
        extra = set(rec.keys()) - allowed
        assert not extra, f'Key(s) {extra} not in allow-list for event={rec.get("event")}'


# ─── #35 extraction statistics — safe summary event ──────────────────────────

def test_summary_event_emitted_in_json_log(tmp_path):
    """summary() must emit a 'summary' event with extraction statistics."""
    log_path = tmp_path / 'stats.jsonl'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='json', source='data.xlsx')
    lg.commit_warnings([lg.warn_validation(1, 1, 'name', 'string', '.+', 'x')])
    lg.summary({'cells': {'name': 'x', 'email': 'y', 'phone': ''}, 'tables': []})
    lg.close()

    records = [json.loads(ln) for ln in
               log_path.read_text(encoding='utf-8').splitlines() if ln.strip()]
    summaries = [r for r in records if r.get('event') == 'summary']
    assert len(summaries) == 1
    s = summaries[0]
    assert s['scalars_defined'] == 3
    assert s['scalars_populated'] == 2
    assert s['scalars_empty'] == 1
    assert s['empty_field_names'] == ['phone']
    assert s['tables_defined'] == 0
    assert s['table_instances'] == 0
    assert s['warnings'] == 1
    assert s['errors'] == 0
    assert s['issues_by_event'] == {'value_mismatch': 1}
    assert isinstance(s['duration_ms'], int)
    assert s['run_id']
    assert s['source'] == 'data.xlsx'


def test_summary_event_not_emitted_in_text_mode(tmp_path):
    """In text log mode there is no summary JSON event."""
    log_path = tmp_path / 'text.log'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='text', source='d.xlsx')
    lg.summary({'cells': {'a': 1}, 'tables': []})
    lg.close()
    content = log_path.read_text(encoding='utf-8')
    assert '"event"' not in content
    assert 'summary' not in content.lower() or 'EXTRACTION SUMMARY' in content


def test_summary_event_contains_no_cell_values(tmp_path):
    """The summary event must not contain any extracted cell values."""
    log_path = tmp_path / 'sentinel.jsonl'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='json', source='d.xlsx')
    lg.summary({'cells': {'name': 'SUPERSECRET', 'code': 42}, 'tables': []})
    lg.close()

    records = [json.loads(ln) for ln in
               log_path.read_text(encoding='utf-8').splitlines() if ln.strip()]
    summaries = [r for r in records if r.get('event') == 'summary']
    assert len(summaries) == 1
    raw = json.dumps(summaries[0])
    assert 'SUPERSECRET' not in raw
    assert '42' not in raw or raw.count('42') == raw.count('"42"')  # not as a value


def test_summary_with_tables(tmp_path):
    """Summary event correctly counts table groups and instances."""
    log_path = tmp_path / 'tbl.jsonl'
    lg = Logger(level=VerbosityLevel.NORMAL, log_file=str(log_path),
                log_format='json', source='d.xlsx')
    tables = [
        {'table_index': 0, 'data': [{'a': 1}]},
        {'table_index': 0, 'data': [{'a': 2}]},
        {'table_index': 1, 'data': [{'b': 3}]},
    ]
    lg.summary({'cells': {}, 'tables': tables})
    lg.close()

    records = [json.loads(ln) for ln in
               log_path.read_text(encoding='utf-8').splitlines() if ln.strip()]
    s = [r for r in records if r.get('event') == 'summary'][0]
    assert s['tables_defined'] == 2
    assert s['table_instances'] == 3


# ─── #36 build_meta — opt-in _meta block ─────────────────────────────────────

def test_build_meta_returns_safe_dict(tmp_path):
    """build_meta() returns a dict with run_id, source, stats, and safe issues."""
    lg = Logger(level=VerbosityLevel.NORMAL, source='data.xlsx')
    lg.commit_warnings([lg.warn_validation(1, 1, 'name', 'string', '.+', 'SECRET')])
    lg.summary({'cells': {'name': 'SECRET', 'code': ''}, 'tables': []})

    meta = lg.build_meta()
    assert meta['run_id']
    assert meta['source'] == 'data.xlsx'
    assert meta['stats']['scalars_defined'] == 2
    assert meta['stats']['scalars_empty'] == 1
    assert meta['stats']['empty_field_names'] == ['code']
    assert meta['stats']['warnings'] == 1
    assert len(meta['issues']) == 1
    raw = json.dumps(meta)
    assert 'SECRET' not in raw


def test_build_meta_issues_use_allow_list_keys():
    """Issues in _meta must use the same allow-list keys as JSON log records."""
    lg = Logger(level=VerbosityLevel.QUIET, source='x.xlsx')
    lg.commit_warnings([lg.warn_validation(1, 1, 'f', 'string', '.+', 'val')])
    lg.summary({'cells': {'f': 'val'}, 'tables': []})

    meta = lg.build_meta()
    for issue in meta['issues']:
        extra = set(issue.keys()) - set(_SAFE_LOG_KEYS)
        assert not extra, f'Issue key(s) {extra} not in _SAFE_LOG_KEYS'
