"""Write extraction results as a structured Excel report for human review.

Layout:
  - Scalar fields: one ``var:`` row per field (col A=``var:``, B=seq#, C=name, D=value)
  - Tables: one ``table:`` header row per instance (``HEADER``), followed by ``DATA``
    rows, an optional ``FOOTER`` row (values right-aligned in data columns, ``label``
    key excluded), and a blank row between instances.
  - Column names are fully qualified: ``table_key.field`` (e.g. ``txn.date``).
"""

from __future__ import annotations

import os
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .utils import flatten_nested, neutralize_formula


# ── color palette ────────────────────────────────────────────────────────────

_FILL_SCALAR = PatternFill('solid', fgColor='D6E4F0')
_FILL_TABLE_HEADER = PatternFill('solid', fgColor='4472C4')
_FILL_TABLE_ROW_EVEN = PatternFill('solid', fgColor='D9E2F3')
_FILL_TABLE_ROW_ODD = PatternFill('solid', fgColor='FFFFFF')
_FILL_TABLE_FOOTER = PatternFill('solid', fgColor='E2EFDA')
_FILL_NONE = PatternFill(fill_type=None)

_FONT_LABEL = Font(bold=True, size=10)
_FONT_TABLE_HEADER = Font(bold=True, color='FFFFFF', size=10)
_FONT_VALUE = Font(size=10)
_FONT_FOOTER = Font(italic=True, size=10)

_ALIGN_LEFT = Alignment(horizontal='left', vertical='top', wrap_text=True)


def _write_row(ws, row: int, total_width: int, values: list) -> None:
    for col_idx in range(1, total_width + 1):
        val = values[col_idx - 1] if col_idx <= len(values) else None
        ws.cell(row=row, column=col_idx, value=val)


def _apply_row_style(ws, row: int, ncols: int, fill, font) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = _ALIGN_LEFT


def _auto_width(ws, min_width: int = 10, max_width: int = 50) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        width = min_width
        for cell in col_cells:
            if cell.value is not None:
                width = max(width, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[col_letter].width = width


def nested_to_xlsx(result: dict, output_path: str, logger=None) -> None:
    """Write a nested extraction result as a structured Excel workbook.

    Format uses prefix columns (A=type, B=index, C=role) so every row is
    machine-readable and visually scannable:
      var:   rows  — scalar field name + value
      table: rows  — HEADER (column names), DATA (values), FOOTER (right-aligned)

    Args:
        result:      Nested extraction result dict.
        output_path: Path to write the .xlsx file.
        logger:      Optional Logger instance; when provided, an "Audit" sheet
                     is added showing per-field provenance and status.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Extraction'

    # ── separate scalars and tables ───────────────────────────────────────
    scalars: list[tuple[str, Any]] = []
    tables: dict[str, list] = {}

    for key, value in result.items():
        if key == '_meta':
            continue
        if isinstance(value, list):
            tables[key] = value
        elif isinstance(value, dict):
            scalars.extend(flatten_nested(value, key))
        else:
            scalars.append((key, value))

    # ── collect qualified column names per table ──────────────────────────
    # e.g. table key 'txn', field 'date' → 'txn.date'
    table_cols: dict[str, list[str]] = {}
    for tbl_key, instances in tables.items():
        cols: list[str] = []
        seen: set[str] = set()
        for inst in instances:
            for data_row in inst.get('data', []):
                for field in data_row:
                    qname = f'{tbl_key}.{field}'
                    if qname not in seen:
                        seen.add(qname)
                        cols.append(qname)
        table_cols[tbl_key] = cols

    # ── compute sheet width ───────────────────────────────────────────────
    # 3 prefix columns + max data columns across all tables (min 1)
    max_data_cols = max((len(cols) for cols in table_cols.values()), default=1)
    total_width = 3 + max_data_cols

    row = 1

    # ── scalar rows ───────────────────────────────────────────────────────
    for seq_num, (field_path, value) in enumerate(scalars, 1):
        _write_row(ws, row, total_width, ['var:', seq_num, field_path, _safe_value(value)])
        _apply_row_style(ws, row, total_width, _FILL_SCALAR, _FONT_VALUE)
        ws.cell(row=row, column=1).font = _FONT_LABEL   # 'var:' bold
        ws.cell(row=row, column=3).font = _FONT_LABEL   # field name bold
        row += 1

    # ── blank separator between scalar and table sections ────────────────────
    if scalars and tables:
        row += 1

    # ── table rows ────────────────────────────────────────────────────────
    for tbl_key, instances in tables.items():
        data_cols = table_cols[tbl_key]   # ['tbl.col1', 'tbl.col2', …]

        for inst_idx, inst in enumerate(instances):
            inst_num = inst_idx + 1

            # HEADER row
            _write_row(ws, row, total_width, ['table:', inst_num, 'HEADER'] + data_cols)
            _apply_row_style(ws, row, total_width, _FILL_TABLE_HEADER, _FONT_TABLE_HEADER)
            row += 1

            # DATA rows
            for data_idx, data_row in enumerate(inst.get('data', [])):
                fields = [
                    _safe_value(data_row.get(col.split('.', 1)[-1]))
                    for col in data_cols
                ]
                _write_row(ws, row, total_width, [None, None, 'DATA'] + fields)
                fill = _FILL_TABLE_ROW_EVEN if data_idx % 2 == 0 else _FILL_TABLE_ROW_ODD
                _apply_row_style(ws, row, total_width, fill, _FONT_VALUE)
                ws.cell(row=row, column=3).font = _FONT_LABEL  # 'DATA' label bold
                row += 1

            # FOOTER row — skip 'label', right-align remaining values
            if inst.get('footer'):
                footer = inst['footer']
                footer_vals = [_safe_value(v) for k, v in footer.items() if k != 'label']
                n_empty = len(data_cols) - len(footer_vals)
                footer_data = [None] * n_empty + footer_vals
                _write_row(ws, row, total_width, [None, None, 'FOOTER'] + footer_data)
                _apply_row_style(ws, row, total_width, _FILL_TABLE_FOOTER, _FONT_FOOTER)
                ws.cell(row=row, column=3).font = _FONT_LABEL  # 'FOOTER' label bold
                row += 1

            # Blank row between instances (not after the last one)
            if inst_idx < len(instances) - 1:
                row += 1

    _auto_width(ws)

    if logger is not None:
        _add_audit_sheet(wb, logger)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)


def _safe_value(v):
    """Convert a value for Excel output; neutralize formula injection (CWE-1236)."""
    if v is None:
        return None
    if isinstance(v, (dict, list)):
        v = str(v)
    if isinstance(v, str):
        return neutralize_formula(v)
    return v


# ── Audit sheet ───────────────────────────────────────────────────────────────

# Audit sheet colour palette
_FILL_AUDIT_CLEAN   = PatternFill('solid', fgColor='E2EFDA')   # light green — clean
_FILL_AUDIT_WARN    = PatternFill('solid', fgColor='FFF2CC')   # amber      — warning
_FILL_AUDIT_MISSING = PatternFill('solid', fgColor='FCE4D6')   # light red  — missing/null
_FILL_AUDIT_HEADER  = PatternFill('solid', fgColor='2F4F8F')   # navy blue  — header row

_FONT_AUDIT_HEADER = Font(bold=True, color='FFFFFF', size=10)
_FONT_AUDIT_BODY   = Font(size=10)

_AUDIT_COLS = ['Field', 'Source Cell', 'Value', 'Status', 'Details']
_AUDIT_COL_WIDTHS = [35, 18, 40, 12, 60]

# Category strings from logger
_CAT_EXTRACTION = 'EXTRACTION'
_CAT_VALIDATION = 'VALIDATION'
_SEV_WARNING    = 'WARNING'
_SEV_ERROR      = 'ERROR'


def _add_audit_sheet(wb: openpyxl.Workbook, logger) -> None:
    """Add a second 'Audit' worksheet to *wb* showing field-level provenance.

    Reads structured ``LogRecord`` objects from ``logger._records`` to build a
    table with one row per extracted scalar field, colour-coded by outcome:

    - 🟢 Green  — field extracted cleanly (no validation issues)
    - 🟡 Amber  — field has a validation warning (type mismatch, pattern miss)
    - 🔴 Red    — field is missing or null (extraction warning)

    Table columns:
      Field | Source Cell | Value | Status | Details

    The audit sheet is read-only documentation — it does not affect extraction.
    """
    ws = wb.create_sheet('Audit')

    # ── Header row ────────────────────────────────────────────────────────────
    for col_idx, (header, width) in enumerate(
        zip(_AUDIT_COLS, _AUDIT_COL_WIDTHS), start=1
    ):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _FONT_AUDIT_HEADER
        cell.fill = _FILL_AUDIT_HEADER
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20

    # ── Collect data from logger records ──────────────────────────────────────
    # index: field_name → list of records about that field
    field_extraction: dict[str, list] = {}   # EXTRACTION INFO records
    field_warnings: dict[str, list] = {}     # VALIDATION WARNING/ERROR records

    for rec in logger._records:
        fname = rec.field
        if not fname:
            continue
        if rec.category == _CAT_EXTRACTION and rec.severity == 'INFO':
            field_extraction.setdefault(fname, []).append(rec)
        elif rec.category == _CAT_VALIDATION:
            field_warnings.setdefault(fname, []).append(rec)

    # Also collect assert: failures (no field name set)
    assert_failures: list = [
        r for r in logger._records
        if r.category == _CAT_VALIDATION and r.severity == _SEV_WARNING
        and not r.field and 'Assertion failed' in r.message
    ]

    # Determine all fields to display (union of extracted + warned)
    all_fields: list[str] = sorted(
        set(field_extraction.keys()) | set(field_warnings.keys())
    )

    # ── Data rows ─────────────────────────────────────────────────────────────
    row = 2
    for fname in all_fields:
        extractions = field_extraction.get(fname, [])
        warnings = field_warnings.get(fname, [])

        # Source cell: from the most recent EXTRACTION record for this field
        source_cell = extractions[-1].location if extractions else ''

        # Value: parse from the last extraction record's message ('field = repr')
        if extractions:
            msg = extractions[-1].message
            # message format: "field = repr(value)" — extract repr part
            eq_idx = msg.find(' = ')
            raw_val = msg[eq_idx + 3:] if eq_idx >= 0 else msg
            # Unquote simple string repr
            if (raw_val.startswith("'") and raw_val.endswith("'")) or \
               (raw_val.startswith('"') and raw_val.endswith('"')):
                raw_val = raw_val[1:-1]
            value_display = neutralize_formula(raw_val)[:200]
        else:
            value_display = None

        # Status and fill
        if warnings:
            # Check if it's a "missing/empty" warning
            is_missing = any(
                'missing' in r.message.lower() or 'empty' in r.message.lower()
                for r in warnings
            )
            if is_missing or not extractions:
                status = 'Missing'
                fill = _FILL_AUDIT_MISSING
            else:
                status = 'Warning'
                fill = _FILL_AUDIT_WARN
        else:
            status = 'Clean'
            fill = _FILL_AUDIT_CLEAN

        # Details: concatenate warning messages
        details = '; '.join(
            neutralize_formula(r.message) for r in warnings[:3]
        ) if warnings else ''

        values = [fname, source_cell, value_display, status, details]
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.fill = fill
            cell.font = _FONT_AUDIT_BODY
            cell.alignment = Alignment(
                horizontal='left', vertical='center',
                wrap_text=(col_idx == len(_AUDIT_COLS)),
            )
        row += 1

    # ── Assert failures section ───────────────────────────────────────────────
    if assert_failures:
        row += 1  # blank separator
        # Mini-header for assert section
        h_cell = ws.cell(row=row, column=1, value='Assert Rules')
        h_cell.font = _FONT_AUDIT_HEADER
        h_cell.fill = _FILL_AUDIT_HEADER
        ws.merge_cells(
            start_row=row, start_column=1,
            end_row=row, end_column=len(_AUDIT_COLS)
        )
        row += 1

        for rec in assert_failures:
            hint = neutralize_formula(rec.hint) if rec.hint else ''
            values = ['', '', '', 'Failed', neutralize_formula(rec.message)]
            for col_idx, val in enumerate(values, start=1):
                cell = ws.cell(row=row, column=col_idx, value=val)
                cell.fill = _FILL_AUDIT_MISSING
                cell.font = _FONT_AUDIT_BODY
                cell.alignment = Alignment(horizontal='left', vertical='center')
            row += 1

    # ── Legend ────────────────────────────────────────────────────────────────
    row += 1
    legend_items = [
        ('Clean',   _FILL_AUDIT_CLEAN,   'Field extracted with no issues'),
        ('Warning', _FILL_AUDIT_WARN,    'Field extracted but validation warning raised'),
        ('Missing', _FILL_AUDIT_MISSING, 'Field not extracted or value is null'),
    ]
    legend_header = ws.cell(row=row, column=1, value='Legend')
    legend_header.font = Font(bold=True, size=9)
    legend_header.alignment = Alignment(horizontal='left')
    row += 1
    for label, fill, desc in legend_items:
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.fill = fill
        label_cell.font = Font(size=9)
        ws.cell(row=row, column=2, value=desc).font = Font(size=9, italic=True)
        row += 1
