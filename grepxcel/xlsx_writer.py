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


def nested_to_xlsx(result: dict, output_path: str) -> None:
    """Write a nested extraction result as a structured Excel workbook.

    Format uses prefix columns (A=type, B=index, C=role) so every row is
    machine-readable and visually scannable:
      var:   rows  — scalar field name + value
      table: rows  — HEADER (column names), DATA (values), FOOTER (right-aligned)
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
