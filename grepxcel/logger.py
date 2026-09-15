"""
Structured logging for the Excel pattern engine.

Records are always stored in memory as structured LogRecord objects,
making them available for console output, file writing, and future
web frontend rendering.

Verbosity levels:
  0  QUIET   — no console output during processing
  1  NORMAL  — summary + ISSUES recap (one line per problem cell) (default)
  2  VERBOSE — + detailed per-cell warnings (Found/Expected/→) + per-field
                trace (🟢/🔴 field ← cell = value), tables matched
  3  DEBUG   — + every anchor attempted and why it was accepted or rejected
"""

from __future__ import annotations
import re as _re_stdlib
from dataclasses import dataclass, field as dc_field
from datetime import datetime, timezone
from enum import IntEnum
from typing import Optional, NoReturn

from .utils import _safe_match, _MAX_REGEX_INPUT_LEN
import hashlib
import json
import os
import sys
import uuid

import structlog

# Bumped only when the structured-log JSON shape changes incompatibly.
LOG_SCHEMA_VERSION = 1

# Structured (JSON) log records are built from an ALLOW-LIST of safe keys only —
# never free-form message/hint/expected/found — so no extracted Excel cell value
# (PII or otherwise) can ever reach a log that might be shipped to a SIEM. A
# non-reversible value fingerprint (length + short sha) aids diagnostics instead.
_SAFE_LOG_KEYS = ('ts', 'level', 'category', 'event', 'cell', 'field',
                  'field_type', 'value_len', 'value_sha8',
                  'run_id', 'source', 'schema_version')

_SAFE_SUMMARY_KEYS = ('ts', 'level', 'category', 'event',
                      'run_id', 'source', 'schema_version',
                      'scalars_defined', 'scalars_populated', 'scalars_empty',
                      'empty_field_names',
                      'tables_defined', 'table_instances',
                      'warnings', 'errors', 'issues_by_event',
                      'duration_ms')


def _value_fingerprint(value) -> "tuple[int, str]":
    """Return (length, sha8) for a value — non-reversible, never the value."""
    s = str(value)
    return len(s), hashlib.sha256(s.encode('utf-8')).hexdigest()[:8]


# ---------------------------------------------------------------------------
# structlog processors — reusable pipeline stages for JSON output
# ---------------------------------------------------------------------------

def allow_list_filter(logger, method_name, event_dict):
    """Strip any keys not in the appropriate allow-list.

    Summary events get _SAFE_SUMMARY_KEYS; all others get _SAFE_LOG_KEYS.
    This is the structured-log safety boundary: no extracted cell value can
    pass through because the allow-lists contain only safe coordinate/metadata
    keys.
    """
    if event_dict.get('event') == 'summary':
        allowed = set(_SAFE_SUMMARY_KEYS)
    else:
        allowed = set(_SAFE_LOG_KEYS)
    return {k: v for k, v in event_dict.items() if k in allowed}


def _make_json_chain():
    """Build the structlog processor chain for NDJSON output."""
    return [
        allow_list_filter,
        structlog.processors.JSONRenderer(default=str),
    ]


_JSON_CHAIN = _make_json_chain()


def render_json(event_dict: dict) -> str:
    """Run the JSON processor chain on an event dict, return an NDJSON line."""
    d = event_dict
    for proc in _JSON_CHAIN:
        d = proc(None, None, d)
    return d


from .color import MARK_FAIL, MARK_OK, MARK_WARN, colorize_marks, paint, should_color


# ---------------------------------------------------------------------------
# Severity and verbosity enums
# ---------------------------------------------------------------------------

class Severity(str):
    DEBUG = 'DEBUG'
    INFO = 'INFO'
    WARNING = 'WARNING'
    ERROR = 'ERROR'


class Category(str):
    ENGINE = 'ENGINE'           # internal engine activity
    EXTRACTION = 'EXTRACTION'   # cells and tables successfully extracted
    VALIDATION = 'VALIDATION'   # data does not match pattern expectations
    STRUCTURAL = 'STRUCTURAL'   # fatal pattern/structure mismatch


class VerbosityLevel(IntEnum):
    QUIET = 0    # no console output; records still collected
    NORMAL = 1   # summary + all validation issues (default)
    VERBOSE = 2  # + step-by-step engine activity
    DEBUG = 3    # + every anchor probe and rejection reason


# ---------------------------------------------------------------------------
# Structured log record — JSON-serialisable, suitable for web rendering
# ---------------------------------------------------------------------------

@dataclass
class LogRecord:
    severity: str           # Severity.*
    category: str           # Category.*
    message: str            # short human-readable description
    location: str = ''      # Excel cell ref, e.g. 'Sheet1!B10'
    field: str = ''         # field name from def:, e.g. 'header.item.qty.label'
    field_type: str = ''    # type from def:, e.g. 'string', 'integer'
    expected: str = ''      # what the pattern required
    found: str = ''         # what was actually in the cell
    hint: str = ''          # actionable suggestion for fixing
    event: str = ''         # stable machine code for structured logs (allow-list)
    value_len: "int | None" = None  # length of the offending value (no value itself)
    value_sha8: str = ''    # first 8 hex of sha256(value) — non-reversible fingerprint
    timestamp: str = dc_field(default_factory=lambda: datetime.now(timezone.utc).isoformat())

    def to_dict(self) -> dict:
        # Exclude private attrs (e.g. the _formatted console-render cache, which
        # bakes in the raw cell value) — to_dict() is the structured data view.
        return {k: v for k, v in self.__dict__.items() if not k.startswith('_')}


# ---------------------------------------------------------------------------
# Fatal engine error — raised to stop processing immediately
# ---------------------------------------------------------------------------

class EngineError(Exception):
    """Raised on fatal structural errors to stop processing immediately."""
    def __init__(self, record: LogRecord):
        super().__init__(record.message)
        self.record = record


# ---------------------------------------------------------------------------
# Cell reference helpers
# ---------------------------------------------------------------------------

def col_letter(n: int) -> str:
    """Convert 1-based column number to Excel letter(s): 1→A, 26→Z, 27→AA."""
    letters = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def cell_ref(row: int, col: int, sheet: str = '') -> str:
    """Return Excel-style cell reference, e.g. 'Sheet1!B10' or 'B10'."""
    ref = f'{col_letter(col)}{row}'
    return f'{sheet}!{ref}' if sheet else ref


# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------

class Logger:
    """
    Collects structured log records and routes output to console and/or file.

    All records are stored in memory regardless of verbosity level,
    so callers (e.g. a web frontend) can always inspect the full log.
    """

    def __init__(
        self,
        level: VerbosityLevel = VerbosityLevel.NORMAL,
        log_file: Optional[str] = None,
        sheet_name: str = 'Sheet1',
        log_format: str = 'text',
        source: Optional[str] = None,
        run_id: Optional[str] = None,
    ):
        self.level = level
        self.sheet_name = sheet_name
        self._records: list[LogRecord] = []
        # Index into _records marking where the current sheet's records begin, so
        # summary()/the ISSUES recap report only *this* sheet's warnings in
        # --all-sheets mode (records accumulate across sheets for programmatic use).
        self._summary_start = 0
        # Structured-logging config. 'json' emits NDJSON (one record per line) to
        # the log file so any SIEM/cloud agent can ingest it. Records are built
        # from an allow-list (see _record_json) so they NEVER contain extracted
        # cell values — there is intentionally no opt-out for that.
        self._log_format = log_format
        self._source = source
        self._run_id = run_id or uuid.uuid4().hex[:12]
        self._start_time = datetime.now(timezone.utc)
        self._bound_context = {
            'run_id': self._run_id,
            'source': self._source,
            'schema_version': LOG_SCHEMA_VERSION,
        }
        self._last_stats: dict | None = None
        self._file = None
        if log_file:
            # Append (never truncate) — the --log file is documented as appended,
            # and a user pointing it at an existing file should not lose its data.
            self._file = open(log_file, 'a', encoding='utf-8')
        # Colour the console copy only when stderr is an interactive terminal.
        self._color = should_color(sys.stderr)

    def _record_event_dict(self, rec: LogRecord) -> dict:
        """Build the raw event dict from a LogRecord.

        Contains all safe coordinate/metadata keys plus the bound context
        (run_id, source, schema_version). The structlog processor chain
        (allow_list_filter → JSONRenderer) enforces that only allow-listed
        keys reach the output — free-form fields (message/hint/expected/found)
        are never included, so no extracted cell value can reach the log.
        """
        d = {
            'ts': rec.timestamp,
            'level': rec.severity,
            'category': rec.category,
            'event': rec.event or rec.category.lower(),
            'cell': rec.location,
            'field': rec.field,
            'field_type': rec.field_type,
            **self._bound_context,
        }
        if rec.value_len is not None:
            d['value_len'] = rec.value_len
            d['value_sha8'] = rec.value_sha8
        return d

    def _store(self, rec: LogRecord) -> None:
        """Record a LogRecord in memory and, in JSON mode, append it to the log
        file as one NDJSON line via the structlog processor chain."""
        self._records.append(rec)
        if self._file and self._log_format == 'json':
            line = render_json(self._record_event_dict(rec))
            self._file.write(line + '\n')
            self._file.flush()

    def close(self):
        if self._file:
            self._file.close()
            self._file = None

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()

    # --- Public record access (for web frontend / programmatic use) ---------

    def records(self) -> list[LogRecord]:
        """Return all collected log records."""
        return list(self._records)

    def to_dict(self) -> list[dict]:
        """Return all records as list of dicts (JSON-serialisable)."""
        return [r.to_dict() for r in self._records]

    def issues(self) -> list[LogRecord]:
        """Return only WARNING and ERROR records."""
        return [r for r in self._records if r.severity in (Severity.WARNING, Severity.ERROR)]

    def has_errors(self) -> bool:
        return any(r.severity == Severity.ERROR for r in self._records)

    def has_warnings(self) -> bool:
        return any(r.severity == Severity.WARNING for r in self._records)

    @property
    def last_stats(self) -> dict | None:
        return self._last_stats

    # --- Engine lifecycle ---------------------------------------------------

    def engine_start(self, pattern_file: str, data_file: str):
        pf = paint(os.path.basename(pattern_file), 'cyan', self._color)
        df = paint(os.path.basename(data_file), 'cyan', self._color)
        self._emit(
            Severity.INFO, Category.ENGINE,
            f'Pattern: {pf} | Data: {df}',
            min_level=VerbosityLevel.NORMAL,
            prefix='ENGINE START',
        )

    def sheet_info(self, sheet_name: str, max_row: int, max_col: int, direction: str):
        sname = paint(sheet_name, 'bold', self._color)
        self._emit(
            Severity.INFO, Category.ENGINE,
            f'Sheet: {sname}  |  {max_row} rows × {max_col} cols  |  '
            f'Scan direction: {direction}',
            min_level=VerbosityLevel.NORMAL,
        )

    def config_verbose(self, config) -> None:
        """Print the active pattern config at VERBOSE level (shown with extract -v).

        Uses the same 8-key layout as validate-pattern -v so config is readable
        in both commands.  Only emitted when verbosity >= VERBOSE.
        """
        def _key(k):   return paint(k, 'dim', self._color)
        def _faint(v): return paint(str(v), 'dim', self._color)
        aliases_val = (', '.join(config.empty_aliases)
                       if config.empty_aliases else _faint('(none)'))
        lines = [
            f'   {_faint("config:")}',
            f'     {_key("pattern.version")} {config.pattern_version}',
            f'     {_key("read.direction")}  {config.read_direction}',
            f'     {_key("currency.sign")}   {config.currency_sign}',
            f'     {_key("ignore.case.labels")}  {config.ignore_case_labels}',
            f'     {_key("ignore.case.values")}  {config.ignore_case_values}',
            f'     {_key("trim.ws.labels")}      {config.trim_whitespace_labels}',
            f'     {_key("trim.ws.values")}      {config.trim_whitespace_values}',
            f'     {_key("lbl.match")}       {config.lbl_match}',
            f'     {_key("var.match")}       {config.var_match}',
            f'     {_key("empty.aliases")}   {aliases_val}',
        ]
        self._write(VerbosityLevel.VERBOSE, '\n'.join(lines))

    def begin_summary_scope(self) -> None:
        """Mark the start of a new sheet's records.

        Called per sheet so the summary + ISSUES recap reflect only the current
        sheet, not warnings accumulated from earlier sheets in --all-sheets mode.
        """
        self._summary_start = len(self._records)

    def summary(self, result: dict):
        cells = result.get('cells', {})
        cells_count = len(cells)
        tables = result.get('tables', [])

        by_group: dict[int, int] = {}
        for t in tables:
            idx = t['table_index']
            by_group[idx] = by_group.get(idx, 0) + 1

        scoped = self._records[self._summary_start:]
        warnings = [r for r in scoped if r.severity == Severity.WARNING]
        errors = [r for r in scoped if r.severity == Severity.ERROR]

        div     = paint('─' * 62, 'dim', self._color)
        header  = paint('EXTRACTION SUMMARY', 'bold', self._color)
        n_warn  = len(warnings)
        n_err   = len(errors)
        w_count = paint(str(n_warn), 'yellow', self._color) if n_warn else str(n_warn)
        e_count = paint(str(n_err),  'red',    self._color) if n_err  else str(n_err)
        lines = [
            '',
            div,
            header,
            f'  Cells extracted   : {cells_count}',
            f'  Mini-tables found : {sum(by_group.values())}',
        ]
        for idx, count in sorted(by_group.items()):
            lines.append(f'    Table group {idx}   : {count} instance(s)')
        lines += [
            f'  Warnings          : {w_count}',
            f'  Errors            : {e_count}',
            div,
        ]
        issues = warnings + errors
        if issues:
            lines.append(paint('ISSUES (cell — reason):', 'bold', self._color))
            for rec in issues:
                lines.append('  ' + self.issue_line(rec))
                if getattr(rec, 'hint', ''):
                    lines.append('       ' + paint(f'→ {rec.hint}', 'dim', self._color))
            lines.append(div)
        self._write(VerbosityLevel.NORMAL, '\n'.join(lines))

        self._last_stats = self.build_stats(result)
        self._emit_summary_event(result)

    def build_stats(self, result: dict) -> dict:
        """Compute safe extraction statistics from the result dict.

        Returns a dict suitable for the _meta block or the JSON log summary
        event.  Contains counts and field names only — never cell values.
        """
        cells = result.get('cells', {})
        tables = result.get('tables', [])
        by_group: dict[int, int] = {}
        for t in tables:
            idx = t['table_index']
            by_group[idx] = by_group.get(idx, 0) + 1

        scoped = self._records[self._summary_start:]
        warnings = [r for r in scoped if r.severity == Severity.WARNING]
        errors = [r for r in scoped if r.severity == Severity.ERROR]

        empty_fields = [k for k, v in cells.items() if v is None or v == '']
        issues_by_event: dict[str, int] = {}
        for rec in warnings + errors:
            ev = rec.event or 'unknown'
            issues_by_event[ev] = issues_by_event.get(ev, 0) + 1

        elapsed = (datetime.now(timezone.utc) - self._start_time)
        duration_ms = int(elapsed.total_seconds() * 1000)

        return {
            'scalars_defined': len(cells),
            'scalars_populated': len(cells) - len(empty_fields),
            'scalars_empty': len(empty_fields),
            'empty_field_names': empty_fields,
            'tables_defined': len(by_group),
            'table_instances': sum(by_group.values()),
            'warnings': len(warnings),
            'errors': len(errors),
            'issues_by_event': issues_by_event,
            'duration_ms': duration_ms,
        }

    def build_meta(self) -> dict:
        """Build the _meta block for opt-in JSON output (--meta).

        Must be called after summary() so that _last_stats is populated.
        """
        scoped = self._records[self._summary_start:]
        safe_issues = []
        for rec in scoped:
            if rec.severity in (Severity.WARNING, Severity.ERROR):
                safe_issues.append(
                    allow_list_filter(None, None, self._record_event_dict(rec))
                )
        return {
            'run_id': self._run_id,
            'source': self._source,
            'schema_version': LOG_SCHEMA_VERSION,
            'stats': self._last_stats or {},
            'issues': safe_issues,
        }

    def _emit_summary_event(self, result: dict) -> None:
        """Write a single 'summary' event to the JSON log with safe statistics."""
        if not (self._file and self._log_format == 'json'):
            return

        event = {
            'ts': datetime.now(timezone.utc).isoformat(),
            'level': Severity.INFO,
            'category': Category.ENGINE,
            'event': 'summary',
            **self._bound_context,
            **self.build_stats(result),
        }
        line = render_json(event)
        self._file.write(line + '\n')
        self._file.flush()

    def issue_line(self, rec: LogRecord) -> str:
        """One concise line summarising a single problem cell for the recap.

        Used both in the NORMAL-level summary block and by --quiet mode, which
        suppresses the summary header/stats but still surfaces each issue.
        """
        mark = MARK_FAIL if rec.severity == Severity.ERROR else MARK_WARN
        where = rec.location or '(no cell)'
        field = f' [{rec.field}]' if rec.field else ''
        if rec.found and rec.expected:
            detail = f'found {rec.found}, expected {rec.expected}'
        elif rec.found:
            detail = f'found {rec.found}'
        else:
            detail = rec.message
        return f'{mark}  {where}{field}  —  {detail}'

    # Keep the private alias so any external callers (tests, scripts) still work.
    _issue_line = issue_line

    # --- Step-by-step (VERBOSE) --------------------------------------------

    def cell_processed(self, row: int, col: int, field: str, value,
                       ok: Optional[bool] = None, regex: str = ''):
        """Trace a scalar cell extraction at VERBOSE: 'field ← B1 = value ✓/✗'.

        ok=None  → no validation mark (e.g. an empty optional field);
        ok=True  → ✓ ; ok=False → ✗ plus the regex it failed.
        """
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.INFO, Category.EXTRACTION,
                        f'{field} = {repr(value)}',
                        location=location, field=field)
        self._store(rec)
        self._write(VerbosityLevel.VERBOSE,
                    self._trace_line(field, location, value, ok, regex, kind='CELL'))

    def _trace_line(self, field: str, location: str, value,
                    ok: Optional[bool] = None, regex: str = '',
                    kind: str = 'CELL') -> str:
        """Render one per-field extraction-trace line.

        Layout: 🟢/🔴 [KIND] field ← Sheet!ref = value   (reason if any)
        The status mark is leftmost so the eye can scan the left edge for pass/fail.
        """
        bracket = paint(f'[{kind}]', 'dim', self._color)
        fname   = paint(f'{field:<20}', 'cyan', self._color)
        loc     = paint(f'{location:<10}', 'dim', self._color)
        if ok is True:
            mark = f'{MARK_OK} '
        elif ok is False:
            mark = f'{MARK_FAIL} '
        else:
            mark = '   '          # 3 spaces: emoji width (2) + separator (1)
        line = f'  {mark}{bracket} {fname} ← {loc} = {repr(value)}'
        if ok is False and regex:
            line += paint(f'   (does not match /{regex}/)', 'dim', self._color)
        return line

    def trace_field(self, row: int, col: int, field: str, value,
                    ok: Optional[bool] = None, regex: str = '') -> str:
        """Build a per-field trace line for a table cell. Returned (not emitted)
        so the caller can commit it only when the mini-table actually matches —
        mirroring how local_warnings are collected and committed."""
        return self._trace_line(field, cell_ref(row, col, self.sheet_name),
                                value, ok, regex, kind='FIELD')

    def commit_traces(self, traces: list) -> None:
        """Emit per-field trace lines collected during a committed mini-table match."""
        for line in traces:
            self._write(VerbosityLevel.VERBOSE, line)

    def cell_ignored(self, row: int, col: int, value):
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.INFO, Category.ENGINE,
                        f'IGNORE: {repr(value)}', location=location)
        self._store(rec)
        bracket = paint('[CELL]', 'dim', self._color)
        loc     = paint(f'{location:<12}', 'dim', self._color)
        self._write(VerbosityLevel.VERBOSE,
                    f'  {bracket}  {loc} IGNORE  →  {repr(value)}')

    def direction_changed(self, direction: str):
        rec = LogRecord(Severity.INFO, Category.ENGINE,
                        f'Scan direction changed to {direction}')
        self._store(rec)
        bracket = paint('[DIR]', 'dim', self._color)
        self._write(VerbosityLevel.VERBOSE,
                    f'  {bracket}   scan direction → {direction}')

    def table_group_start(self, table_index: int, direction: str):
        rec = LogRecord(Severity.INFO, Category.ENGINE,
                        f'Table group {table_index}: scanning (direction: {direction})')
        self._store(rec)
        label = paint(f'[TABLE {table_index}]', 'cyan', self._color)
        self._write(VerbosityLevel.VERBOSE,
                    f'\n  {label}  Scanning for mini-tables  '
                    f'(direction: {direction})')

    def table_group_done(self, table_index: int, count: int):
        msg = (f'Table group {table_index}: {count} instance(s) found'
               if count else f'Table group {table_index}: no instances found')
        rec = LogRecord(Severity.INFO, Category.ENGINE, msg)
        self._store(rec)
        label = paint(f'[TABLE {table_index}]', 'cyan', self._color)
        text = (f'  {label}  {count} instance(s) extracted'
                if count else f'  {label}  No instances found')
        self._write(VerbosityLevel.VERBOSE, text)

    def mini_table_matched(self, table_index: int, instance: int,
                           anchor_row: int, anchor_col: int,
                           end_row: int, end_col: int):
        anchor = cell_ref(anchor_row, anchor_col, self.sheet_name)
        span = (f'{col_letter(anchor_col)}{anchor_row}:'
                f'{col_letter(end_col)}{end_row}')
        rec = LogRecord(Severity.INFO, Category.EXTRACTION,
                        f'Mini-table matched at {anchor}, span {span}',
                        location=anchor)
        self._store(rec)
        label = paint('[MATCH]', 'cyan', self._color)
        self._write(VerbosityLevel.VERBOSE,
                    f'    {label}  instance {instance}  anchor {anchor}  '
                    f'span {span}')

    # --- Debug probes (DEBUG) -----------------------------------------------

    def anchor_probe(self, row: int, col: int, value):
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Probing anchor {location}: {repr(value)}',
                        location=location)
        self._store(rec)
        label = paint('[PROBE]', 'dim', self._color)
        self._write(VerbosityLevel.DEBUG,
                    f'    {label}  {location}: {repr(value)}')

    def anchor_rejected(self, row: int, col: int, reason: str):
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Anchor rejected: {reason}', location=location)
        self._store(rec)
        label = paint('[REJECT]', 'dim', self._color)
        self._write(VerbosityLevel.DEBUG,
                    f'    {label} {location}: {reason}')

    def data_row(self, row: int, col_count: int):
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Data row {row}: {col_count} column(s)')
        self._store(rec)
        label = paint('[DATA]', 'dim', self._color)
        self._write(VerbosityLevel.DEBUG,
                    f'      {label}  row {row}: {col_count} column(s)')

    def data_row_skipped(self, row: int):
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Data row {row}: skipped (matches SKIP_IF)')
        self._store(rec)
        label = paint('[SKIP]', 'dim', self._color)
        self._write(VerbosityLevel.DEBUG,
                    f'      {label}  row {row}: matches SKIP_IF — skipped')

    def warn_data_min_not_reached(self, min_rows: int, found: int) -> LogRecord:
        """Warn when a DATA:{n,m} section has fewer physical rows than declared minimum."""
        hint = (
            f'The pattern declares DATA:{{{min_rows},…}} but only {found} physical '
            f'row(s) were found before the footer or the max limit. '
            f'Check that the data file has at least {min_rows} row(s) in this section, '
            f'or lower the minimum bound in the pattern.'
        )
        rec = LogRecord(
            severity=Severity.WARNING,
            category=Category.VALIDATION,
            message=f'DATA minimum not reached: expected ≥{min_rows} rows, found {found}',
            expected=f'≥{min_rows} physical data rows',
            found=str(found),
            hint=hint,
        )
        lines = [
            f'\n  {MARK_WARN}  [DATA min not reached]',
            f'     Found:    {found} physical row(s)',
            f'     Expected: ≥{min_rows} row(s)',
            f'     → {hint}',
        ]
        rec._formatted = '\n'.join(lines)
        return rec

    def footer_detected(self, row: int, col: int, value):
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Footer detected at {location}: {repr(value)}',
                        location=location)
        self._store(rec)
        label = paint('[FOOTER]', 'dim', self._color)
        self._write(VerbosityLevel.DEBUG,
                    f'    {label} {location}: {repr(value)} — ending DATA section')

    # --- Validation warnings (NORMAL) ---------------------------------------

    def warn_validation(self, row: int, col: int, field: str, field_type: str,
                        regex: str, value) -> LogRecord:
        """
        Log a validation warning: value was extracted but does not match its pattern.
        Returns the LogRecord so callers can collect and commit it later.
        """
        location = cell_ref(row, col, self.sheet_name)
        found_repr = repr(value)
        hint = self._hint_validation(field_type, regex, value)

        rec = LogRecord(
            severity=Severity.WARNING,
            category=Category.VALIDATION,
            message='Value does not match the expected pattern',
            location=location,
            field=field,
            field_type=field_type,
            expected=f'matches /{regex}/',
            found=found_repr,
            hint=hint,
            event='value_mismatch',
        )
        if value is not None:
            rec.value_len, rec.value_sha8 = _value_fingerprint(value)
        lines = [
            f'\n  {MARK_WARN}  {location}  [{field} / {field_type}]',
            f'     Found:    {found_repr}',
            f'     Expected: matches /{regex}/',
        ]
        if hint:
            lines.append(f'     → {hint}')
        rec._formatted = '\n'.join(lines)
        return rec

    def warn_empty_field(self, row: int, col: int, field: str, field_type: str) -> LogRecord:
        """
        Log a warning for an empty cell where a value was expected.
        Returns the LogRecord so callers can collect and commit it later.
        """
        location = cell_ref(row, col, self.sheet_name)
        hint = (
            f'Cell {location} is empty but the pattern defines it as '
            f'a required {field_type} field. '
            f'Verify that the data file contains a value here, or check whether '
            f'the pattern column count is wider than the actual table.'
        )
        rec = LogRecord(
            severity=Severity.WARNING,
            category=Category.VALIDATION,
            message='Required field is empty',
            location=location,
            field=field,
            field_type=field_type,
            expected=f'non-empty {field_type} value',
            found='empty cell',
            hint=hint,
            event='empty_required',
        )
        lines = [
            f'\n  {MARK_WARN}  {location}  [{field} / {field_type}]',
            f'     Found:    empty cell',
            f'     Expected: non-empty {field_type} value',
            f'     → {hint}',
        ]
        rec._formatted = '\n'.join(lines)
        return rec

    def commit_warnings(self, records: list[LogRecord]):
        """
        Commit a batch of warning records collected during a successful match.
        Records are always stored; the detailed per-cell block (Found/Expected/→)
        only prints at VERBOSE (-v).  The summary ISSUES recap (one line each)
        still prints at NORMAL so no information is lost.
        """
        for rec in records:
            self._store(rec)
            self._write(VerbosityLevel.VERBOSE,
                        getattr(rec, '_formatted', rec.message))

    # --- Fatal errors (always shown) ----------------------------------------

    def warn_undefined_field(self, row: int, col: int, field: str) -> LogRecord:
        """
        Log a warning for a field referenced in the pattern but missing from the def: section.
        Returns the LogRecord so callers can collect and commit it later.
        """
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(
            severity=Severity.WARNING,
            category=Category.STRUCTURAL,
            message=f'Field {field!r} is not defined in the pattern def: section',
            location=location,
            field=field,
            hint=f'Add a def: row to the pattern file for field {field!r}.',
            event='undefined_field',
        )
        lines = [
            f'\n  {MARK_WARN}  {location}  [{field} / undefined]',
            f'     Field {field!r} is not defined in the pattern def: section.',
            f'     → Add a def: row to the pattern file for this field name.',
        ]
        rec._formatted = '\n'.join(lines)
        return rec

    def warn_uncached_formulas(self) -> None:
        """Emit a one-time warning when the data file contains uncached formula cells."""
        msg = (
            'Sheet contains formula cells whose cached values are missing. '
            'Open the file in Excel or LibreOffice, save it, and re-run grepxcel '
            'to ensure formula results are available.'
        )
        self._write(VerbosityLevel.NORMAL, f'\n  {MARK_WARN}  {msg}')

    def warn_assert(self, message: str, hint: str = '') -> LogRecord:
        """Record a failed assert: cross-field rule as a WARNING."""
        rec = LogRecord(
            severity=Severity.WARNING,
            category=Category.VALIDATION,
            message=message,
            hint=hint,
        )
        self._store(rec)
        lines = [f'\n  {MARK_WARN}  [assert] {message}']
        if hint:
            lines.append(f'     → {hint}')
        rec._formatted = '\n'.join(lines)
        self._write(VerbosityLevel.NORMAL, rec._formatted)
        return rec

    def fatal(self, message: str, location: str = '',
              expected: str = '', found: str = '') -> NoReturn:
        """Log a fatal error and raise EngineError to stop processing."""
        rec = LogRecord(
            severity=Severity.ERROR,
            category=Category.STRUCTURAL,
            message=message,
            location=location,
            expected=expected,
            found=found,
        )
        self._store(rec)

        label = paint('FATAL ERROR', 'bold', self._color)
        lines = [f'\n  {MARK_FAIL}  {label}: {message}']
        if location:
            lines.append(f'     Location: {location}')
        if expected:
            lines.append(f'     Expected: {expected}')
        if found:
            lines.append(f'     Found:    {found}')
        lines.append('     Processing stopped.')
        self._write(VerbosityLevel.QUIET, '\n'.join(lines))

        raise EngineError(rec)

    # --- Internal helpers ---------------------------------------------------

    def _emit(self, severity: str, category: str, message: str,
              min_level: VerbosityLevel = VerbosityLevel.NORMAL,
              prefix: str = ''):
        rec = LogRecord(severity=severity, category=category, message=message)
        self._store(rec)
        if prefix:
            tag = paint(f'[{prefix}]', 'bold', self._color)
            text = f'{tag} {message}'
        else:
            text = message
        self._write(min_level, text)

    def _write(self, min_level: VerbosityLevel, text: str):
        if self.level >= min_level:
            print(colorize_marks(text, self._color), file=sys.stderr)
        # In text mode the file mirrors the console. In json mode the file is
        # NDJSON written per-record by _store(), so skip the human text here.
        if self._file and self._log_format == 'text':
            self._file.write(text + '\n')
            self._file.flush()

    def _hint_validation(self, field_type: str, regex: str, value) -> str:
        # Leading/trailing whitespace — most actionable when trimming fixes the mismatch.
        if isinstance(value, str) and value != value.strip():
            trimmed = value.strip()
            if trimmed and _safe_match(regex, trimmed, _re_stdlib.DOTALL, _MAX_REGEX_INPUT_LEN):
                return (
                    f"The cell value has leading or trailing whitespace. "
                    f"The trimmed value {trimmed!r} DOES match /{regex}/. "
                    f"Add 'trim-whitespace' to this field in column A "
                    f"(e.g. 'var:trim-whitespace') or set "
                    f"'config: | trim.whitespace | yes' to strip whitespace globally."
                )
            elif trimmed:
                return (
                    f"The cell value {value!r} has leading or trailing whitespace — "
                    f"trimmed to {trimmed!r}, which still does not match /{regex}/. "
                    f"The mismatch is not caused solely by whitespace."
                )

        if isinstance(value, str):
            if '\n' in value or '\r' in value:
                if '.*' not in regex and r'[\s\S]' not in regex and r'\n' not in regex:
                    if ' ' in regex:
                        suggested = regex.replace(' ', r'[\s\S]*')
                    else:
                        suggested = regex + r' (add [\s\S]* where the newline occurs)'
                    return (
                        'The cell value contains a newline character (created by pressing '
                        'Alt+Enter in Excel). The pattern regex does not allow newlines. '
                        'Update the pattern definition to use a newline-aware pattern, '
                        f'e.g.: {suggested!r}'
                    )

        if field_type == 'integer':
            try:
                v = int(float(value))
                return (
                    f'The integer {v} does not satisfy /{regex}/. '
                    f'Check the required digit count and leading-digit rules '
                    f'in the def: section of the pattern file.'
                )
            except (ValueError, TypeError):
                pass

        if field_type == 'currency':
            try:
                f_val = float(value)
                s = str(f_val)
                if '.' in s and len(s.split('.')[1]) > 2:
                    return (
                        f'The value {f_val} has a floating-point precision tail '
                        f'(e.g. 1.3800000000000001 instead of 1.38). '
                        f'This is an Excel/Python float artefact. '
                        f'Consider rounding the value after extraction, or '
                        f'widening the regex to accept more decimal places.'
                    )
            except (ValueError, TypeError):
                pass

        return ''
