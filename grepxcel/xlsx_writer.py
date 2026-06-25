"""Write extraction results as a colored Excel report for human review.

Layout (top-down):
  1. Scalar fields — one row per field: path | extracted value
  2. Blank separator
  3. Each table key — header row (bold), data rows (alternating), footer if present
"""

from __future__ import annotations

import os
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .utils import flatten_nested, neutralize_formula


# ── color palette ────────────────────────────────────────────────────────────

_FILL_SECTION_HEADER = PatternFill('solid', fgColor='2F5496')
_FILL_SCALAR_ROW = PatternFill('solid', fgColor='D6E4F0')
_FILL_TABLE_HEADER = PatternFill('solid', fgColor='4472C4')
_FILL_TABLE_ROW_EVEN = PatternFill('solid', fgColor='D9E2F3')
_FILL_TABLE_ROW_ODD = PatternFill('solid', fgColor='FFFFFF')
_FILL_TABLE_FOOTER = PatternFill('solid', fgColor='E2EFDA')
_FILL_SOURCE = PatternFill('solid', fgColor='F2F2F2')

_FONT_SECTION_HEADER = Font(bold=True, color='FFFFFF', size=11)
_FONT_TABLE_HEADER = Font(bold=True, color='FFFFFF', size=10)
_FONT_FIELD_PATH = Font(bold=True, size=10)
_FONT_VALUE = Font(size=10)
_FONT_FOOTER = Font(italic=True, size=10)
_FONT_SOURCE = Font(color='808080', size=9)

_ALIGN_LEFT = Alignment(horizontal='left', vertical='top', wrap_text=True)


def _apply_row_style(ws, row: int, ncols: int, fill, font):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = _ALIGN_LEFT


def _auto_width(ws, min_width: int = 10, max_width: int = 50):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        width = min_width
        for cell in col_cells:
            if cell.value is not None:
                width = max(width, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[col_letter].width = width


def nested_to_xlsx(result: dict, output_path: str) -> None:
    """Write a nested extraction result as a colored Excel workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Extraction'

    row = 1
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

    # ── section: scalars ─────────────────────────────────────────────────
    if scalars:
        ws.cell(row=row, column=1, value='Scalar Fields')
        ws.cell(row=row, column=2, value='Value')
        _apply_row_style(ws, row, 2, _FILL_SECTION_HEADER, _FONT_SECTION_HEADER)
        row += 1

        for field_path, value in scalars:
            ws.cell(row=row, column=1, value=field_path)
            ws.cell(row=row, column=1).font = _FONT_FIELD_PATH
            ws.cell(row=row, column=1).fill = _FILL_SCALAR_ROW
            ws.cell(row=row, column=1).alignment = _ALIGN_LEFT

            ws.cell(row=row, column=2, value=_safe_value(value))
            ws.cell(row=row, column=2).font = _FONT_VALUE
            ws.cell(row=row, column=2).fill = _FILL_SCALAR_ROW
            ws.cell(row=row, column=2).alignment = _ALIGN_LEFT
            row += 1

        row += 1  # blank separator

    # ── section: tables ──────────────────────────────────────────────────
    for table_key, instances in tables.items():
        all_data_cols: list[str] = []
        seen: set[str] = set()
        for inst in instances:
            for data_row in inst.get('data', []):
                for k in data_row:
                    if k not in seen:
                        seen.add(k)
                        all_data_cols.append(k)

        header_cols: list[str] = []
        hseen: set[str] = set()
        for inst in instances:
            if inst.get('header'):
                for k, v in flatten_nested(inst['header']):
                    if k not in hseen:
                        hseen.add(k)
                        header_cols.append(k)

        footer_cols: list[str] = []
        fseen: set[str] = set()
        for inst in instances:
            if inst.get('footer'):
                for k, v in flatten_nested(inst['footer']):
                    if k not in fseen:
                        fseen.add(k)
                        footer_cols.append(k)

        ncols = max(len(all_data_cols), 2)

        # table title
        ws.cell(row=row, column=1, value=f'Table: {table_key}')
        _apply_row_style(ws, row, ncols, _FILL_SECTION_HEADER, _FONT_SECTION_HEADER)
        row += 1

        for inst_idx, inst in enumerate(instances):
            # source reference
            source = inst.get('_source')
            if source:
                src_text = f"[{source.get('sheet', '?')} {source.get('ref', '')}]"
                ws.cell(row=row, column=1, value=src_text)
                _apply_row_style(ws, row, ncols, _FILL_SOURCE, _FONT_SOURCE)
                row += 1

            # instance header (from pattern HEADER: rows)
            if inst.get('header'):
                flat_h = flatten_nested(inst['header'])
                for col_idx, (k, v) in enumerate(flat_h, 1):
                    ws.cell(row=row, column=col_idx, value=f'{k}: {_safe_value(v)}')
                _apply_row_style(ws, row, max(len(flat_h), ncols),
                                 _FILL_TABLE_FOOTER, _FONT_FOOTER)
                row += 1

            # column headers
            if all_data_cols:
                for col_idx, col_name in enumerate(all_data_cols, 1):
                    ws.cell(row=row, column=col_idx, value=col_name)
                _apply_row_style(ws, row, len(all_data_cols),
                                 _FILL_TABLE_HEADER, _FONT_TABLE_HEADER)
                row += 1

            # data rows
            for data_idx, data_row in enumerate(inst.get('data', [])):
                fill = _FILL_TABLE_ROW_EVEN if data_idx % 2 == 0 else _FILL_TABLE_ROW_ODD
                for col_idx, col_name in enumerate(all_data_cols, 1):
                    ws.cell(row=row, column=col_idx,
                            value=_safe_value(data_row.get(col_name)))
                _apply_row_style(ws, row, len(all_data_cols), fill, _FONT_VALUE)
                row += 1

            # footer
            if inst.get('footer'):
                flat_f = flatten_nested(inst['footer'])
                for col_idx, (k, v) in enumerate(flat_f, 1):
                    ws.cell(row=row, column=col_idx, value=f'{k}: {_safe_value(v)}')
                _apply_row_style(ws, row, max(len(flat_f), ncols),
                                 _FILL_TABLE_FOOTER, _FONT_FOOTER)
                row += 1

            # separator between instances
            if inst_idx < len(instances) - 1:
                row += 1

        row += 1  # blank row before next table

    _auto_width(ws)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)


def _safe_value(v):
    """Convert a value to something Excel can display.

    Strings that would be parsed as formulas are neutralised so an extracted
    value from an untrusted file is never written as an executable formula
    (CWE-1236).
    """
    if v is None:
        return ''
    if isinstance(v, (dict, list)):
        v = str(v)
    return neutralize_formula(v)
