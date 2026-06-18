"""
Structured logging for the Excel pattern engine.

Records are always stored in memory as structured LogRecord objects,
making them available for console output, file writing, and future
web frontend rendering.

Verbosity levels:
  0  QUIET   — no console output during processing
  1  NORMAL  — summary + ISSUES recap (one line per problem cell) (default)
  2  VERBOSE — + detailed per-cell warnings (Found/Expected/→) + per-field
                trace (field ← cell = value ✓/✗), tables matched
  3  DEBUG   — + every anchor attempted and why it was accepted or rejected
"""

from __future__ import annotations
from dataclasses import dataclass, field as dc_field
from datetime import datetime, timezone
from enum import IntEnum
from typing import Optional, NoReturn
import os
import sys

from .color import colorize_marks, should_color


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
    timestamp: str = dc_field(default_factory=lambda: datetime.now(timezone.utc).isoformat())

    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items()}


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
    ):
        self.level = level
        self.sheet_name = sheet_name
        self._records: list[LogRecord] = []
        # Index into _records marking where the current sheet's records begin, so
        # summary()/the ISSUES recap report only *this* sheet's warnings in
        # --all-sheets mode (records accumulate across sheets for programmatic use).
        self._summary_start = 0
        self._file = None
        if log_file:
            self._file = open(log_file, 'w', encoding='utf-8')
        # Colour the console copy only when stderr is an interactive terminal.
        self._color = should_color(sys.stderr)

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

    # --- Engine lifecycle ---------------------------------------------------

    def engine_start(self, pattern_file: str, data_file: str):
        self._emit(
            Severity.INFO, Category.ENGINE,
            f'Pattern: {os.path.basename(pattern_file)} | '
            f'Data: {os.path.basename(data_file)}',
            min_level=VerbosityLevel.NORMAL,
            prefix='ENGINE START',
        )

    def sheet_info(self, sheet_name: str, max_row: int, max_col: int, direction: str):
        self._emit(
            Severity.INFO, Category.ENGINE,
            f'Sheet: {sheet_name}  |  {max_row} rows × {max_col} cols  |  '
            f'Scan direction: {direction}',
            min_level=VerbosityLevel.NORMAL,
        )

    def begin_summary_scope(self) -> None:
        """Mark the start of a new sheet's records.

        Called per sheet so the summary + ISSUES recap reflect only the current
        sheet, not warnings accumulated from earlier sheets in --all-sheets mode.
        """
        self._summary_start = len(self._records)

    def summary(self, result: dict):
        cells_count = len(result.get('cells', {}))
        tables = result.get('tables', [])

        by_group: dict[int, int] = {}
        for t in tables:
            idx = t['table_index']
            by_group[idx] = by_group.get(idx, 0) + 1

        scoped = self._records[self._summary_start:]
        warnings = [r for r in scoped if r.severity == Severity.WARNING]
        errors = [r for r in scoped if r.severity == Severity.ERROR]

        lines = [
            '',
            '─' * 62,
            'EXTRACTION SUMMARY',
            f'  Cells extracted   : {cells_count}',
            f'  Mini-tables found : {sum(by_group.values())}',
        ]
        for idx, count in sorted(by_group.items()):
            lines.append(f'    Table group {idx}   : {count} instance(s)')
        lines += [
            f'  Warnings          : {len(warnings)}',
            f'  Errors            : {len(errors)}',
            '─' * 62,
        ]
        # Consolidated recap of every problem cell, so failures are never lost
        # in the scrollback of a long run — one clear line per issue.
        issues = warnings + errors
        if issues:
            lines.append('ISSUES (cell — reason):')
            for rec in issues:
                lines.append('  ' + self._issue_line(rec))
            lines.append('─' * 62)
        self._write(VerbosityLevel.NORMAL, '\n'.join(lines))

    def _issue_line(self, rec: LogRecord) -> str:
        """One concise line summarising a single problem cell for the recap."""
        mark = '✗' if rec.severity == Severity.ERROR else '⚠'
        where = rec.location or '(no cell)'
        field = f' [{rec.field}]' if rec.field else ''
        if rec.found and rec.expected:
            detail = f'found {rec.found}, expected {rec.expected}'
        elif rec.found:
            detail = f'found {rec.found}'
        else:
            detail = rec.message
        return f'{mark}  {where}{field}  —  {detail}'

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
        self._records.append(rec)
        self._write(VerbosityLevel.VERBOSE,
                    self._trace_line(field, location, value, ok, regex, kind='CELL'))

    def _trace_line(self, field: str, location: str, value,
                    ok: Optional[bool] = None, regex: str = '',
                    kind: str = 'CELL') -> str:
        """Render one per-field extraction-trace line."""
        mark = '' if ok is None else (' ✓' if ok else ' ✗')
        line = f'  [{kind}] {field:<20} ← {location:<10} = {repr(value)}{mark}'
        if ok is False and regex:
            line += f'   (does not match /{regex}/)'
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
        self._records.append(rec)
        self._write(VerbosityLevel.VERBOSE,
                    f'  [CELL]  {location:<12} IGNORE  →  {repr(value)}')

    def direction_changed(self, direction: str):
        rec = LogRecord(Severity.INFO, Category.ENGINE,
                        f'Scan direction changed to {direction}')
        self._records.append(rec)
        self._write(VerbosityLevel.VERBOSE,
                    f'  [DIR]   scan direction → {direction}')

    def table_group_start(self, table_index: int, direction: str):
        rec = LogRecord(Severity.INFO, Category.ENGINE,
                        f'Table group {table_index}: scanning (direction: {direction})')
        self._records.append(rec)
        self._write(VerbosityLevel.VERBOSE,
                    f'\n  [TABLE {table_index}]  Scanning for mini-tables  '
                    f'(direction: {direction})')

    def table_group_done(self, table_index: int, count: int):
        msg = (f'Table group {table_index}: {count} instance(s) found'
               if count else f'Table group {table_index}: no instances found')
        rec = LogRecord(Severity.INFO, Category.ENGINE, msg)
        self._records.append(rec)
        text = (f'  [TABLE {table_index}]  {count} instance(s) extracted'
                if count else f'  [TABLE {table_index}]  No instances found')
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
        self._records.append(rec)
        self._write(VerbosityLevel.VERBOSE,
                    f'    [MATCH]  instance {instance}  anchor {anchor}  '
                    f'span {span}')

    # --- Debug probes (DEBUG) -----------------------------------------------

    def anchor_probe(self, row: int, col: int, value):
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Probing anchor {location}: {repr(value)}',
                        location=location)
        self._records.append(rec)
        self._write(VerbosityLevel.DEBUG,
                    f'    [PROBE]  {location}: {repr(value)}')

    def anchor_rejected(self, row: int, col: int, reason: str):
        location = cell_ref(row, col, self.sheet_name)
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Anchor rejected: {reason}', location=location)
        self._records.append(rec)
        self._write(VerbosityLevel.DEBUG,
                    f'    [REJECT] {location}: {reason}')

    def data_row(self, row: int, col_count: int):
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Data row {row}: {col_count} column(s)')
        self._records.append(rec)
        self._write(VerbosityLevel.DEBUG,
                    f'      [DATA]  row {row}: {col_count} column(s)')

    def data_row_skipped(self, row: int):
        rec = LogRecord(Severity.DEBUG, Category.ENGINE,
                        f'Data row {row}: skipped (matches SKIP_IF)')
        self._records.append(rec)
        self._write(VerbosityLevel.DEBUG,
                    f'      [SKIP]  row {row}: matches SKIP_IF — skipped')

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
            f'\n  ⚠  [DATA min not reached]',
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
        self._records.append(rec)
        self._write(VerbosityLevel.DEBUG,
                    f'    [FOOTER] {location}: {repr(value)} — ending DATA section')

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
        )
        lines = [
            f'\n  ⚠  {location}  [{field} / {field_type}]',
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
        )
        lines = [
            f'\n  ⚠  {location}  [{field} / {field_type}]',
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
            self._records.append(rec)
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
        )
        lines = [
            f'\n  ⚠  {location}  [{field} / undefined]',
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
        self._write(VerbosityLevel.NORMAL, f'\n  ⚠  {msg}')

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
        self._records.append(rec)

        lines = [f'\n  ✗  FATAL ERROR: {message}']
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
        self._records.append(rec)
        text = f'[{prefix}] {message}' if prefix else message
        self._write(min_level, text)

    def _write(self, min_level: VerbosityLevel, text: str):
        if self.level >= min_level:
            print(colorize_marks(text, self._color), file=sys.stderr)
        if self._file:
            self._file.write(text + '\n')   # file log is always plain text
            self._file.flush()

    def _hint_validation(self, field_type: str, regex: str, value) -> str:
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
