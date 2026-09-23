"""
grepxcel web wizard — FastAPI backend.

Launched by ``grepxcel web-wizard data.xlsx``.  Opens a browser window at
http://localhost:<port> where the user can classify cells visually instead of
using the terminal TUI.  Single-user local tool; state lives in a module-level
dict for the lifetime of the server process.

Optional dependency group: ``pip install "grepxcel[web]"``
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
import threading
import time
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

# ── Reuse existing pure functions ─────────────────────────────────────────────
from grepxcel.wizard_core import (
    WizardState,
    _build_state_from_choices,
    _build_cell_order,
    _choices_to_csv,
    _col_a_extra_from_parts,
    _col_a_extra_to_parts,
    _preload_from_pattern,
    _slugify,
    _infer_cell_type,
)

try:
    from fastapi import FastAPI, HTTPException, Request
    from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse
    from fastapi.staticfiles import StaticFiles
    from jinja2 import Environment, FileSystemLoader
    import uvicorn
    _WEB_OK = True
except ImportError:
    _WEB_OK = False

# ── Module-level session (single-user local tool) ─────────────────────────────

_STATE: dict[str, Any] = {}   # single session dict


# ── Session log ───────────────────────────────────────────────────────────────

class _SessionLog:
    """Writes a structured session log in the same format as the TUI wizard."""

    def __init__(self, xlsx_path: str, sheet_name: str) -> None:
        self._fh = None
        self._path: str | None = None

        try:
            stem     = Path(xlsx_path).stem
            ts_file  = datetime.now().strftime('%Y-%m-%d-%H%M%S')
            abs_path = str(Path(xlsx_path).resolve())

            # SHA-256 + size
            sha256 = hashlib.sha256()
            file_size = 0
            with open(xlsx_path, 'rb') as fh:
                for chunk in iter(lambda: fh.read(65536), b''):
                    sha256.update(chunk)
                    file_size += len(chunk)
            sha8 = sha256.hexdigest()[:8]

            session_dir = Path(os.getcwd()) / 'logs' / 'wizard' / stem / f'{ts_file}_{sha8}'
            session_dir.mkdir(parents=True, exist_ok=True)
            log_path = session_dir / 'session.log'
            self._fh   = log_path.open('w', encoding='utf-8')
            self._path = str(log_path)

            sep = '═' * 51
            v = _get_version()
            self._fh.write('grepxcel web-wizard session\n')
            self._fh.write(f'{sep}\n')
            self._fh.write(f'Data file : {abs_path}\n')
            self._fh.write(f'SHA-256   : {sha256.hexdigest()}\n')
            self._fh.write(f'File size : {file_size} bytes\n')
            self._fh.write(f'Sheet     : {sheet_name}\n')
            self._fh.write(f'Interface : web (FastAPI)\n')
            self._fh.write(f'Version   : grepxcel {v}\n')
            self._fh.write(f'Started   : {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            self._fh.write(f'{sep}\n')
            self._fh.write('# Columns: timestamp   EVENT_TYPE    detail\n')
            self._fh.write('# CLASSIFY      ref:ACTION name=… — cell classified\n')
            self._fh.write('# CONFIG        direction=… — global config changed\n')
            self._fh.write('# PRELOAD       pattern=… cells_loaded=N — pattern file pre-filled session\n')
            self._fh.write('# PRELOAD_FIELD ref:role name=… — one field loaded from pattern\n')
            self._fh.write('# PRELOAD_WARN  message — warning from preload pass\n')
            self._fh.write('# SCHEMA_SUMMARY lbl=N var=N total=N mods=… — state at save\n')
            self._fh.write('# UNDO          ref — last classification reversed\n')
            self._fh.write('# SAVE          path — pattern file written to disk\n')
            self._fh.write(f'{sep}\n\n')
            self._fh.flush()
        except OSError:
            pass

    def write(self, event_type: str, detail: str) -> None:
        if not self._fh:
            return
        now = datetime.now()
        ts  = now.strftime('%H:%M:%S.') + f'{now.microsecond // 1000:03d}'
        try:
            self._fh.write(f'{ts}  {event_type:<12}  {detail}\n')
            self._fh.flush()
        except OSError:
            pass

    def close(self, choices: dict, notes: dict) -> None:
        if not self._fh:
            return
        try:
            ts = datetime.now().strftime('%H:%M:%S')
            self._fh.write('#\n')
            self._fh.write(f'# Session ended: {ts}\n')
            counts: dict = {}
            for info in choices.values():
                ch = info.get('choice', '?')
                counts[ch] = counts.get(ch, 0) + 1
            self._fh.write(f'# Classifications: {counts}\n')
            if notes:
                self._fh.write('#\n# Cell notes:\n')
                for ref, note in sorted(notes.items()):
                    self._fh.write(f'#   {ref}: {note}\n')
            self._fh.close()
            self._fh = None
        except OSError:
            pass

    @property
    def path(self) -> str | None:
        return self._path


def _cell_ref(row: int, col: int) -> str:
    return f'{get_column_letter(col)}{row}'


def _ui_role_to_meta(role: str) -> str:
    """Map UI role V/L/I/E → internal meta role var/label/ignore.

    V = Variable (extract value)
    L = Label (text anchor to match)
    I = Ignore (skip column)
    E = Empty (expect blank cell; same meta role as I but orig_field='EMPTY')
    C = legacy alias for L (was mistakenly included; retained for backward compat)
    """
    r = (role or 'V').upper()
    if r == 'V':            return 'var'
    if r in ('L', 'C'):     return 'label'
    return 'ignore'   # I, E, or anything else


def _web_row_configs_to_meta(anchor_ref: str, end_ref: str, name: str, mult: str,
                              row_configs: list,
                              ws: 'openpyxl.worksheet.worksheet.Worksheet') -> dict:
    """Convert the web modal row_configs list to a standard T-anchor meta dict.

    This meta dict is understood by ``_choices_to_csv`` in wizard_tui.py.
    row_configs is a list of dicts with keys:
      sheet_row   int             — the actual worksheet row number
      row_type    str             — 'header', 'data', 'data_inherited',
                                    'footer', 'skip', 'ignore'
      row_n       int (optional)  — for header/footer, the :N suffix (default 1)
      cols        list (optional) — per-column config (not present for data_inherited)
    """
    ar, ac = _parse_ref(anchor_ref)
    er, ec = _parse_ref(end_ref)
    start_row, end_row = min(ar, er), max(ar, er)
    start_col, end_col = min(ac, ec), max(ac, ec)
    table_name = name.strip()

    def _build_col(i: int, sheet_row: int, col_cfg: dict,
                   role_default: str = 'V', hf_mode: bool = False) -> dict:
        """Build a column descriptor dict for one cell in a table row.

        hf_mode=True (header/footer rows): label columns.
          - If a field name was given (raw_name, e.g. ``header_item_lbl``): use it as
            the HEADER/FOOTER column identifier AND emit a global lbl: definition.
          - If no field name: fall back to the cell text as an implicit literal identifier
            (the engine matches it directly without a global lbl: def; no_global_lbl=True).
        """
        c = start_col + i
        role = _ui_role_to_meta(col_cfg.get('role', role_default))
        raw_name = (col_cfg.get('name') or '').strip()
        cell_val = ws.cell(row=sheet_row, column=c).value
        cell_str = str(cell_val).strip() if cell_val is not None else ''

        if hf_mode and role == 'label':
            lmatch_text = (col_cfg.get('lmatch') or '').strip() or cell_str
            # Prefer explicit field name; fall back to cell text as implicit identifier.
            fn = raw_name or lmatch_text or 'IGNORE'
        elif (table_name and raw_name and raw_name.upper() != 'IGNORE'
                and not raw_name.startswith(table_name + '.')
                and '.' not in raw_name):   # already namespace-qualified → don't add prefix
            fn = f'{table_name}.{raw_name}'
        else:
            fn = raw_name or 'IGNORE'

        # Propagate modifiers → col_a_extra (e.g. 'nullable', 'not-null', 'trim-whitespace')
        modifiers_raw = col_cfg.get('modifiers') or 'none'
        col_a_extra   = _col_a_extra_from_parts('', modifiers_raw)

        # For ignore-role columns, preserve the EMPTY vs IGNORE distinction so
        # round-trip CSV generation is lossless.
        # E role always maps to EMPTY; I role uses the stored orig_field (default IGNORE).
        if role == 'ignore':
            ui_role = (col_cfg.get('role') or '').upper()
            if ui_role == 'E':
                orig_field_ignore = 'EMPTY'
            else:
                orig_field_ignore = col_cfg.get('orig_field', 'IGNORE')
        else:
            orig_field_ignore = ''

        return {
            'ref':          _cell_ref(sheet_row, c),
            'role':         role,
            'var_name':     fn if role == 'var'   else 'IGNORE',
            'var_type':     col_cfg.get('ftype', 'string'),
            'var_match':    col_cfg.get('match', '.*'),
            'col_a_extra':  col_a_extra,
            'lbl_name':     fn if role == 'label' else 'IGNORE',
            'lbl_type':     'string',
            'lbl_match':    (col_cfg.get('lmatch') or '').strip() or cell_str,
            'lbl_mode':     col_cfg.get('lbl_mode', ''),
            'cell_value':   cell_str,
            'orig_field':   orig_field_ignore,
            # Suppress global lbl: def only when the column identifier IS the raw cell
            # text (no explicit field name given).  Named fields (e.g. header_item_lbl)
            # require a global def so the engine can locate and match them.
            'no_global_lbl': hf_mode and role == 'label' and not raw_name,
        }

    header_rows: list = []
    data_vars: list | None = None
    footer_rows: list = []
    skip_configs: list = []
    row_types: dict = {}

    for rc in row_configs:
        sheet_row = int(rc.get('sheet_row', 0))
        rtype     = (rc.get('row_type') or '').lower()
        cols      = rc.get('cols') or []

        if rtype.startswith('header'):
            row_types[sheet_row] = 'H'
            header_rows.append({'row': sheet_row, 'cols': [
                _build_col(i, sheet_row, c, role_default='L', hf_mode=True) for i, c in enumerate(cols)
            ]})

        elif rtype == 'data':
            row_types[sheet_row] = 'D'
            if data_vars is None:           # first data row defines the schema
                data_vars = [
                    _build_col(i, sheet_row, c) for i, c in enumerate(cols)
                ]

        elif rtype == 'data_inherited':
            row_types[sheet_row] = 'D'     # schema inherited from first data row

        elif rtype.startswith('footer'):
            row_types[sheet_row] = 'F'
            footer_rows.append({'row': sheet_row, 'cols': [
                _build_col(i, sheet_row, c, role_default='L', hf_mode=True) for i, c in enumerate(cols)
            ]})

        elif rtype == 'skip':
            row_types[sheet_row] = 'S'
            s_cols = []
            for i, col in enumerate(cols):
                cond   = col.get('condition', 'IGNORE').upper()
                lmatch = (col.get('lmatch') or '').strip()
                if cond == 'LABEL' and lmatch:
                    # Auto-generate an lbl_name scoped to the table
                    lbl_name = f'{table_name}.skip_{sheet_row}_{i}' if table_name else f'skip_{sheet_row}_{i}'
                else:
                    lbl_name = 'IGNORE'
                s_cols.append({'condition': cond, 'lmatch': lmatch, 'lbl_name': lbl_name})
            skip_configs.append({'cols': s_cols})

    return {
        'choice':           'T',
        'name':             table_name,
        'mult':             mult,
        'start_ref':        anchor_ref,
        'end_ref':          end_ref,
        'start_row':        start_row,
        'end_row':          end_row,
        'start_col':        start_col,
        'end_col':          end_col,
        'header_rows':      header_rows,
        'data_vars':        data_vars if data_vars is not None else [],
        'footer_rows':      footer_rows,
        '_web_skip_configs': skip_configs,
        '_web_row_configs': row_configs,   # preserved for round-trip edit
        'row_types':        row_types,
    }


def _parse_ref(ref: str) -> tuple[int, int]:
    """'B3' → (row=3, col=2)."""
    col_str = ''.join(c for c in ref if c.isalpha()).upper()
    row_str = ''.join(c for c in ref if c.isdigit())
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return int(row_str), col


def _cell_display(value: Any, max_len: int = 28) -> str:
    if value is None:
        return ''
    if isinstance(value, datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0 and value.microsecond == 0:
            s = value.strftime('%Y-%m-%d')
        else:
            s = value.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(value, timedelta):
        total_secs = int(value.total_seconds())
        h = total_secs // 3600
        m = (total_secs % 3600) // 60
        sec = total_secs % 60
        s = f'{h}:{m:02d}:{sec:02d}'
    else:
        s = str(value)
    if len(s) > max_len:
        return s[:max_len] + '…'
    return s


def _to_json_safe(obj: Any) -> Any:
    """Recursively convert non-JSON-serialisable types for API responses.

    Specifically converts ``datetime`` objects to ISO-8601 strings so that the
    extraction result can be returned directly as JSON without a custom encoder.
    """
    if isinstance(obj, dict):
        return {k: _to_json_safe(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_to_json_safe(v) for v in obj]
    if isinstance(obj, datetime):
        return obj.isoformat()
    if obj is None or isinstance(obj, (bool, int, float, str)):
        return obj
    return str(obj)


def _cell_type_display(cell) -> str:
    """Return a display type string for a cell, using number_format for richer inference."""
    if cell.value is None:
        # Check image presence separately via has_image flag in _build_sheet_data
        return 'empty'
    return _infer_cell_type(cell)


def _build_merge_info(ws) -> tuple[dict, set]:
    """Return (merge_topleft, merge_skip) from the worksheet's merged cell ranges.

    merge_topleft  dict[ref → (colspan, rowspan)] for anchor cells
    merge_skip     set[ref] of non-anchor cells that must not be rendered
    """
    merge_topleft: dict[str, tuple[int, int]] = {}
    merge_skip: set[str] = set()
    for m in ws.merged_cells.ranges:
        tl = _cell_ref(m.min_row, m.min_col)
        colspan = m.max_col - m.min_col + 1
        rowspan = m.max_row - m.min_row + 1
        merge_topleft[tl] = (colspan, rowspan)
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                ref = _cell_ref(r, c)
                if ref != tl:
                    merge_skip.add(ref)
    return merge_topleft, merge_skip


def _build_sheet_data() -> dict:
    """Serialize the active worksheet into a JSON-friendly structure."""
    ws: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']
    choices: dict = _STATE['choices']
    notes: dict = _STATE['notes']
    # image_cells is {sheet_name: {cell_ref: {count, mimes, suspicious}}}
    _all_image_cells: dict = _STATE.get('image_cells', {})
    image_cells: dict = _all_image_cells.get(ws.title, {})
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    # Cap to reasonable display size
    display_rows = min(max_row, _STATE.get('max_rows', 150))
    display_cols = min(max_col, _STATE.get('max_cols', 40))

    merge_topleft, merge_skip = _build_merge_info(ws)

    rows = []
    for r in range(1, display_rows + 1):
        row = []
        for c in range(1, display_cols + 1):
            ref = _cell_ref(r, c)
            cell = ws.cell(row=r, column=c)
            choice_info = choices.get(ref, {})
            colspan, rowspan = merge_topleft.get(ref, (1, 1))
            is_anchor = ref in merge_topleft
            is_skip   = ref in merge_skip
            img_info    = image_cells.get(ref)
            has_img     = img_info is not None
            img_suspicious = has_img and img_info.get('suspicious', False)
            cell_type = 'image' if has_img else _cell_type_display(cell)
            row.append({
                'ref':       ref,
                'row':       r,
                'col':       c,
                'col_letter': get_column_letter(c),
                'value':     _cell_display(cell.value),
                'raw':       str(cell.value) if cell.value is not None else '',
                'type':      cell_type,
                'has_image':       has_img,
                'image_suspicious': img_suspicious,
                'image_count': img_info['count'] if img_info else 0,
                'choice':      choice_info.get('choice', ''),
                'name':        choice_info.get('name', ''),
                'anchor':      choice_info.get('anchor', ''),  # set for T-HEAD / T-DATA
                'table_role':    choice_info.get('table_role', ''),    # L/V/I within the table row
                'row_class':     choice_info.get('row_class', ''),     # header/data/footer
                'is_table_end':  choice_info.get('is_table_end', False),  # True on the bottom-right cell
                'note':        notes.get(ref, ''),
                'empty':     cell.value is None and not is_anchor,
                'colspan':   colspan,
                'rowspan':   rowspan,
                'merged':    is_anchor,   # True = top-left of a merged range
                'skip':      is_skip,     # True = inside a merge, must not render
            })
        rows.append(row)

    col_letters = [get_column_letter(c) for c in range(1, display_cols + 1)]
    wb = _STATE.get('wb')
    all_sheets = wb.sheetnames if wb else [ws.title]

    return {
        'rows':        rows,
        'max_row':     display_rows,
        'max_col':     display_cols,
        'col_letters': col_letters,
        'total_rows':  max_row,
        'total_cols':  max_col,
        'choices':     choices,
        'notes':       notes,
        'sheet_name':  ws.title,
        'all_sheets':  all_sheets,
    }


def _build_stats() -> dict:
    ws       = _STATE.get('ws')
    choices  = _STATE.get('choices', {})

    # Ignore any choice stored on a non-anchor merged cell (ghost classifications)
    _, merge_skip = _build_merge_info(ws) if ws else ({}, set())

    counts: dict[str, int] = {}
    for ref, info in choices.items():
        if ref in merge_skip:
            continue          # don't count ghost cells from pre-existing patterns
        ch = info.get('choice', '')
        counts[ch] = counts.get(ch, 0) + 1

    # "Total" = cells the user needs to make a decision about: non-empty cells
    # OR cells that already have a classification (e.g. empty T-DATA cells in a
    # table range are structurally classified even though their value is None).
    non_empty_refs: set[str] = set()
    if ws:
        for r in range(1, (ws.max_row or 1) + 1):
            for c in range(1, (ws.max_column or 1) + 1):
                ref = _cell_ref(r, c)
                if ref not in merge_skip and ws.cell(row=r, column=c).value is not None:
                    non_empty_refs.add(ref)
    classified_refs = set(choices) - merge_skip
    # Union so that classified empty cells also count in total → classified ≤ total
    total      = len(non_empty_refs | classified_refs)
    classified = len(classified_refs)
    return {
        'L':      counts.get('L', 0),
        'V':      counts.get('V', 0),
        'T':      counts.get('T', 0),
        'T_HEAD': counts.get('T-HEAD', 0),   # table column-header cells (from preload)
        'T_DATA': counts.get('T-DATA', 0),   # table data-row cells (from preload)
        'I':      counts.get('I', 0),
        'total':      total,
        'classified': classified,
        'unclassified': max(0, len(non_empty_refs) - classified),
    }


def _push_undo(ref: str) -> None:
    stack: list = _STATE.setdefault('undo_stack', [])
    snapshot = {
        'ref':     ref,
        'choices': json.loads(json.dumps(_STATE.get('choices', {}))),
        'notes':   json.loads(json.dumps(_STATE.get('notes', {}))),
    }
    stack.append(snapshot)
    if len(stack) > 50:
        stack.pop(0)


def create_app(
    xlsx_path: str,
    pattern_path: str | None = None,
    max_rows: int = 150,
    max_cols: int = 40,
    sheet: str | None = None,
    max_file_mb: float = 5,
    max_uncompressed_mb: float = 50,
) -> 'FastAPI':
    """Create and return the FastAPI application for the web wizard."""
    if not _WEB_OK:
        raise ImportError(
            'Web wizard requires FastAPI and uvicorn. '
            'Install with: pip install "grepxcel[web]"'
        )

    # ── Security checks (same guards as the extract command) ─────────────────
    from .security import validate_file, validate_pattern_file
    from .pattern_check import check_pattern
    from .color import MARK_FAIL, MARK_WARN, colorize_marks, paint, should_color
    validate_file(xlsx_path, max_file_mb=max_file_mb, max_uncompressed_mb=max_uncompressed_mb)
    _pv_errors: list[str] = []
    _pv_warnings: list[str] = []
    if pattern_path and Path(pattern_path).exists():
        validate_pattern_file(pattern_path)
        _pv = check_pattern(pattern_path)
        _color = should_color(sys.stderr)
        if _pv.errors:
            print(colorize_marks(
                f'{MARK_FAIL}  Pattern has errors — loaded for editing but cannot extract:',
                _color), file=sys.stderr)
            for e in _pv.errors:
                print(colorize_marks(f'   {MARK_FAIL}  {e}', _color), file=sys.stderr)
            _pv_errors = _pv.errors
        if _pv.warnings:
            if not _pv.errors:
                print(colorize_marks(f'{MARK_WARN}  Pattern warnings:', _color), file=sys.stderr)
            for w in _pv.warnings:
                print(colorize_marks(f'   {MARK_WARN}  {w}', _color), file=sys.stderr)
            _pv_warnings = _pv.warnings

    # ── Load workbook ─────────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # Select initial sheet (--sheet flag or active)
    if sheet is None:
        ws = wb.active
    elif isinstance(sheet, int) or (isinstance(sheet, str) and sheet.lstrip('-').isdigit()):
        idx = int(sheet)
        ws = wb.worksheets[max(0, min(idx, len(wb.worksheets) - 1))]
    elif sheet in wb.sheetnames:
        ws = wb[sheet]
    else:
        ws = wb.active  # unknown name → fall back to active

    # ── Image presence scan (ZIP, no Pillow needed) ───────────────────────────
    try:
        from .engine import scan_image_cells as _scan_img
        _image_cells = _scan_img(xlsx_path)
    except Exception:
        _image_cells = {}

    # ── Image extraction for in-wizard preview ────────────────────────────────
    _extracted_images: dict[str, str] = {}  # {cell_ref: abs_file_path}
    _img_tmp_dir: str | None = None
    try:
        from .engine import extract_images as _ext_img
        import tempfile
        _img_tmp_dir = tempfile.mkdtemp(prefix='grepxcel_wizard_img_')
        _extracted_images, _ = _ext_img(xlsx_path, _img_tmp_dir,
                                         stem=Path(xlsx_path).stem)
    except Exception:
        pass

    # ── Pre-populate from existing pattern ────────────────────────────────────
    state = WizardState(sheet_name=ws.title)
    choices: dict[str, dict] = {}
    notes: dict[str, str] = {}

    # Warnings from preload surfaced to the UI log panel.
    preload_warnings: list[str] = [
        f'[ERROR] {e}' for e in _pv_errors
    ] + [
        f'[WARN] {w}' for w in _pv_warnings
    ]

    if pattern_path and Path(pattern_path).exists():
        try:
            loaded_choices, preload_cfg, _warnings = _preload_from_pattern(ws, pattern_path)
            preload_warnings.extend(_warnings)
            choices.update(loaded_choices)
            if preload_cfg.get('direction'):
                state.direction       = preload_cfg['direction']
            if preload_cfg.get('lbl_match'):
                state.lbl_match       = preload_cfg['lbl_match']
            if preload_cfg.get('var_match'):
                state.var_match       = preload_cfg['var_match']
            # Preload: handle both old single-key and new split-key patterns.
            _ic = preload_cfg.get('ignore_case')
            if _ic is not None:
                state.ignore_case_labels = _ic
                state.ignore_case_values = _ic
            if preload_cfg.get('ignore_case_labels') is not None:
                state.ignore_case_labels = preload_cfg['ignore_case_labels']
            if preload_cfg.get('ignore_case_values') is not None:
                state.ignore_case_values = preload_cfg['ignore_case_values']
            _tw = preload_cfg.get('trim_whitespace')
            if _tw is not None:
                state.trim_whitespace_values = _tw
            if preload_cfg.get('trim_whitespace_labels') is not None:
                state.trim_whitespace_labels = preload_cfg['trim_whitespace_labels']
            if preload_cfg.get('trim_whitespace_values') is not None:
                state.trim_whitespace_values = preload_cfg['trim_whitespace_values']
            if preload_cfg.get('currency_sign'):
                state.currency_sign   = preload_cfg['currency_sign']
            if preload_cfg.get('empty_aliases'):
                state.empty_aliases   = preload_cfg['empty_aliases']
        except Exception as exc:  # noqa: BLE001
            preload_warnings.append(f'Could not load pattern: {exc}')

    # ── Session log ───────────────────────────────────────────────────────────
    session_log = _SessionLog(xlsx_path, ws.title)

    # ── Module-level session ──────────────────────────────────────────────────
    _STATE.clear()
    _STATE.update({
        'xlsx_path':        xlsx_path,
        'pattern_path':     pattern_path,
        'wb':               wb,
        'ws':               ws,
        'state':            state,
        'choices':          choices,
        'choices_by_sheet': {ws.title: choices},  # in-memory per-sheet choices
        'notes':            notes,
        'undo_stack':       [],
        'max_rows':         max_rows,
        'max_cols':         max_cols,
        'log':              session_log,
        'preload_warnings': preload_warnings,
        'image_cells':      _image_cells,
        'extracted_images': _extracted_images,
        'img_tmp_dir':      _img_tmp_dir,
    })

    # Log initial config
    st0 = state
    session_log.write('CONFIG',
        f'direction={st0.direction} '
        f'ic_labels={st0.ignore_case_labels} ic_values={st0.ignore_case_values} '
        f'trim_labels={st0.trim_whitespace_labels} trim_values={st0.trim_whitespace_values} '
        f'currency_sign={st0.currency_sign!r} '
        f'lbl_match={st0.lbl_match!r} var_match={st0.var_match!r} '
        f'aliases={st0.empty_aliases}'
    )
    if pattern_path:
        session_log.write('PRELOAD',
            f'pattern={pattern_path} cells_loaded={len(choices)}'
        )
        for w in preload_warnings:
            session_log.write('PRELOAD_WARN', w)
        for ref, info in sorted(choices.items()):
            ch = info.get('choice', '?')
            if ch == 'L':
                session_log.write('PRELOAD_FIELD',
                    f'{ref}:lbl name={info.get("name","")!r} '
                    f'ltype={info.get("ltype","")!r} '
                    f'lmatch={info.get("lmatch","")!r} '
                    f'lbl_mode={info.get("lbl_mode","")!r}'
                )
            elif ch == 'V':
                session_log.write('PRELOAD_FIELD',
                    f'{ref}:var name={info.get("name","")!r} '
                    f'ftype={info.get("ftype","")!r} '
                    f'match={info.get("match","")!r} '
                    f'col_a_extra={info.get("col_a_extra","")!r}'
                )

    # ── Jinja2 env ────────────────────────────────────────────────────────────
    tpl_dir = Path(__file__).parent / 'templates'
    jinja_env = Environment(loader=FileSystemLoader(str(tpl_dir)),
                            autoescape=True)

    # ── FastAPI app ───────────────────────────────────────────────────────────
    app = FastAPI(title='grepxcel Web Wizard', docs_url=None, redoc_url=None)
    static_dir = Path(__file__).parent / 'static'
    app.mount('/static', StaticFiles(directory=str(static_dir)), name='static')

    # ─────────────────────────────── HTML PAGE ────────────────────────────────

    @app.get('/', response_class=HTMLResponse)
    async def index():
        tpl = jinja_env.get_template('wizard.html')
        html = tpl.render(
            xlsx_path=xlsx_path,
            sheet_name=ws.title,
            gx_version=_get_version(),
        )
        return HTMLResponse(html)

    # ─────────────────────────────── API ─────────────────────────────────────

    @app.get('/api/sheet')
    async def api_sheet():
        return JSONResponse(_build_sheet_data())

    @app.get('/api/sheets')
    async def api_sheets():
        """Return the list of all sheet names and which is currently active."""
        wb_ = _STATE['wb']
        return JSONResponse({
            'sheets':  wb_.sheetnames,
            'active':  _STATE['ws'].title,
        })

    @app.post('/api/switch-sheet')
    async def api_switch_sheet(request: Request):
        """Switch to a different sheet, preserving current choices in memory."""
        body = await request.json()
        target = body.get('sheet', '')
        wb_ = _STATE['wb']

        if target not in wb_.sheetnames:
            # Try numeric index
            if str(target).lstrip('-').isdigit():
                idx = int(target)
                if 0 <= idx < len(wb_.worksheets):
                    target = wb_.worksheets[idx].title
            if target not in wb_.sheetnames:
                raise HTTPException(404, f'Sheet {target!r} not found')

        current_sheet = _STATE['ws'].title
        if target == current_sheet:
            return JSONResponse({'ok': True, 'sheet': current_sheet, 'changed': False})

        # Save current choices to the per-sheet store
        _STATE['choices_by_sheet'][current_sheet] = _STATE['choices']

        # Switch worksheet
        new_ws = wb_[target]
        _STATE['ws'] = new_ws
        _STATE['state'].sheet_name = target

        # Restore or create choices for the new sheet
        new_choices = _STATE['choices_by_sheet'].get(target, {})
        _STATE['choices'] = new_choices
        _STATE['choices_by_sheet'][target] = new_choices
        _STATE['notes'] = {}
        _STATE['undo_stack'] = []

        _STATE['log'].write('SWITCH_SHEET', f'from={current_sheet!r} to={target!r}')
        return JSONResponse({'ok': True, 'sheet': target, 'changed': True})

    @app.get('/api/list-patterns')
    async def api_list_patterns():
        """List pattern files in the same directory as the data xlsx."""
        data_dir = Path(_STATE['xlsx_path']).parent
        patterns = []
        for ext in ('*.xlsx', '*.csv'):
            for p in sorted(data_dir.glob(ext), key=lambda f: f.stat().st_mtime, reverse=True):
                name_lower = p.name.lower()
                if 'pattern' in name_lower:
                    patterns.append({
                        'name':  p.name,
                        'path':  str(p),
                        'size':  p.stat().st_size,
                    })
        return JSONResponse({'patterns': patterns})

    @app.get('/api/state')
    async def api_state():
        st: WizardState = _STATE['state']
        return JSONResponse({
            'config': {
                'direction':       st.direction,
                'template':               st.template,
                'ignore_case_labels':     st.ignore_case_labels,
                'ignore_case_values':     st.ignore_case_values,
                'trim_whitespace_labels': st.trim_whitespace_labels,
                'trim_whitespace_values': st.trim_whitespace_values,
                'currency_sign':   st.currency_sign,
                'lbl_match':       st.lbl_match,
                'var_match':       st.var_match,
                'empty_aliases':   st.empty_aliases,
            },
            'choices': _STATE['choices'],
            'notes':   _STATE['notes'],
            'stats':   _build_stats(),
        })

    @app.get('/api/logs')
    async def api_logs():
        """Return session log entries and preload warnings for the UI log panel."""
        log_obj: _SessionLog = _STATE.get('log')
        log_path = log_obj._path if log_obj else None
        recent_lines: list[str] = []
        if log_path:
            try:
                with open(log_path, encoding='utf-8') as fh:
                    recent_lines = fh.readlines()[-200:]
            except OSError:
                pass
        return JSONResponse({
            'preload_warnings': _STATE.get('preload_warnings', []),
            'log_path':         log_path,
            'log_lines':        [l.rstrip('\n') for l in recent_lines],
        })

    _VALID_DIRECTIONS   = frozenset({'LR', 'TD'})
    _VALID_MATCH_MODES  = frozenset({'', 'literal', 'glob', 'regexp'})

    @app.post('/api/config')
    async def api_config(request: Request):
        body = await request.json()
        st: WizardState = _STATE['state']
        # Validate enum fields at API entry time so invalid values are rejected
        # immediately rather than surfacing as opaque PatternError at extraction.
        direction = body.get('direction', st.direction)
        if direction not in _VALID_DIRECTIONS:
            raise HTTPException(400, f"Invalid direction {direction!r}; must be 'LR' or 'TD'")
        lbl_match = body.get('lbl_match', st.lbl_match)
        if lbl_match not in _VALID_MATCH_MODES:
            raise HTTPException(400,
                f"Invalid lbl_match {lbl_match!r}; must be one of {sorted(_VALID_MATCH_MODES)}")
        var_match = body.get('var_match', st.var_match)
        if var_match not in _VALID_MATCH_MODES:
            raise HTTPException(400,
                f"Invalid var_match {var_match!r}; must be one of {sorted(_VALID_MATCH_MODES)}")
        st.direction              = direction
        st.template               = body.get('template',               st.template)
        st.ignore_case_labels     = body.get('ignore_case_labels',     st.ignore_case_labels)
        st.ignore_case_values     = body.get('ignore_case_values',     st.ignore_case_values)
        st.trim_whitespace_labels = body.get('trim_whitespace_labels', st.trim_whitespace_labels)
        st.trim_whitespace_values = body.get('trim_whitespace_values', st.trim_whitespace_values)
        st.currency_sign          = body.get('currency_sign',          st.currency_sign)
        st.lbl_match              = lbl_match
        st.var_match              = var_match
        st.empty_aliases          = body.get('empty_aliases',          st.empty_aliases)
        _STATE['log'].write('CONFIG',
            f'direction={st.direction} '
            f'ic_labels={st.ignore_case_labels} ic_values={st.ignore_case_values} '
            f'trim_labels={st.trim_whitespace_labels} trim_values={st.trim_whitespace_values} '
            f'currency_sign={st.currency_sign!r} '
            f'lbl_match={st.lbl_match!r} var_match={st.var_match!r} '
            f'aliases={st.empty_aliases}'
        )
        return JSONResponse({'ok': True, 'config': body})

    @app.post('/api/classify')
    async def api_classify(request: Request):
        body   = await request.json()
        ref    = body.get('ref', '').upper()
        action = body.get('action', '')  # L, V, C, I, T, CLEAR
        fields = body.get('fields', {})
        note   = fields.get('notes', '').strip()

        if not ref or action not in ('L', 'V', 'I', 'T', 'CLEAR'):
            raise HTTPException(400, f'Invalid ref={ref!r} or action={action!r}')
        # Block classifying a non-anchor merged cell (ghost cell)
        _, merge_skip = _build_merge_info(_STATE['ws'])
        if ref in merge_skip:
            raise HTTPException(400, f'{ref} is inside a merged cell — classify the top-left anchor instead')

        _push_undo(ref)
        choices = _STATE['choices']
        notes   = _STATE['notes']

        if action == 'CLEAR':
            removed_meta = choices.pop(ref, {})
            notes.pop(ref, None)
            removed_choice = removed_meta.get('choice', '')
            # Determine cascade anchor (T anchor clears its members; member clears its anchor+siblings)
            if removed_choice == 'T':
                cascade_anchor = ref
            elif removed_choice in ('T-HEAD', 'T-DATA'):
                cascade_anchor = removed_meta.get('anchor', '')
                if cascade_anchor:
                    choices.pop(cascade_anchor, None)
                    notes.pop(cascade_anchor, None)
            else:
                cascade_anchor = ''
            cleared = [ref]
            if cascade_anchor:
                stale = [r for r, m in list(choices.items())
                         if m.get('choice') in ('T-HEAD', 'T-DATA') and m.get('anchor') == cascade_anchor]
                for r in stale:
                    choices.pop(r, None)
                    notes.pop(r, None)
                    cleared.append(r)
                if cascade_anchor != ref:
                    cleared.append(cascade_anchor)
            _STATE['log'].write('CLASSIFY', f'{ref}:CLEAR ({len(cleared)} cells)')
            return JSONResponse({'ok': True, 'ref': ref, 'action': 'CLEAR', 'cleared': cleared})

        ws: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']
        try:
            row, col = _parse_ref(ref)
        except (ValueError, IndexError):
            raise HTTPException(400, f'Cannot parse cell reference {ref!r}')
        try:
            cell_value = ws.cell(row=row, column=col).value
        except Exception as exc:
            raise HTTPException(400, f'Cannot read cell {ref}: {exc}')

        if action == 'L':
            lbl_mode_raw = fields.get('match_mode', '')
            # '(default)' is the UI sentinel for "use global default" — normalize to ''
            lbl_mode = '' if lbl_mode_raw == '(default)' else lbl_mode_raw
            choices[ref] = {
                'choice':   'L',
                'name':     fields.get('name', _slugify(str(cell_value or '')) + '_label'),
                'ltype':    fields.get('type', 'string'),
                'lmatch':   fields.get('match', str(cell_value or '')),
                'lbl_mode': lbl_mode,
            }
        elif action == 'V':
            var_mode_raw  = fields.get('match_mode', '(default)')
            modifiers_raw = fields.get('modifiers', 'none')
            col_a_extra   = _col_a_extra_from_parts(var_mode_raw, modifiers_raw)
            match_pattern = fields.get('match', '.*')
            # ReDoS guard: validate the user-supplied regex before storing it.
            # Patterns are saved to disk and later passed to the engine's regex
            # engine against potentially large cell values.
            from .security import check_regex_safety, SecurityError as _SecErr
            try:
                check_regex_safety(match_pattern, field_name=fields.get('name', ref))
            except _SecErr as exc:
                raise HTTPException(400, str(exc))
            choices[ref] = {
                'choice':      'V',
                'name':        fields.get('name', _slugify(str(cell_value or ''))),
                'ftype':       fields.get('type', _infer_cell_type(ws.cell(row=row, column=col))),
                'match':       match_pattern,
                'col_a_extra': col_a_extra,
            }
        elif action == 'I':
            choices[ref] = {'choice': 'I'}
        elif action == 'T':
            table_name  = (fields.get('name') or _slugify(str(cell_value or '')) + '_table').strip()
            mult        = fields.get('mult', '*')
            end_ref_raw = (fields.get('end_ref') or '').strip().upper()
            row_configs = fields.get('row_configs', [])

            # Clear any previous T-HEAD/T-DATA cells that belonged to this anchor
            stale = [r for r, m in choices.items()
                     if m.get('choice') in ('T-HEAD', 'T-DATA')
                     and m.get('anchor') == ref]
            for r in stale:
                choices.pop(r, None)

            ws_: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']

            if row_configs and end_ref_raw:
                # Full table modal config — convert to the standard meta format
                meta_dict = _web_row_configs_to_meta(
                    ref, end_ref_raw, table_name, mult, row_configs, ws_)
                choices[ref] = meta_dict

                # Mark T-HEAD cells (all cells in header rows, except the anchor itself)
                _anchor_col_role = (meta_dict.get('header_rows') or [{}])[0]
                _anchor_col_role = (_anchor_col_role.get('cols') or [{}])[0].get('role', 'ignore') if _anchor_col_role else 'ignore'
                choices[ref]['table_role'] = 'L' if _anchor_col_role == 'label' else ('V' if _anchor_col_role == 'var' else 'I')
                choices[ref]['row_class']  = 'header'
                for h_row in meta_dict.get('header_rows', []):
                    for col_d in h_row.get('cols', []):
                        cref = col_d.get('ref', '')
                        if cref and cref != ref:
                            role   = col_d.get('role', 'ignore')
                            t_role = 'L' if role == 'label' else ('V' if role == 'var' else 'I')
                            choices[cref] = {'choice': 'T-HEAD', 'anchor': ref,
                                             'table_role': t_role, 'row_class': 'header'}

                # Mark T-HEAD cells for footer rows
                for f_row in meta_dict.get('footer_rows', []):
                    for col_d in f_row.get('cols', []):
                        cref = col_d.get('ref', '')
                        if cref and cref != ref:
                            role   = col_d.get('role', 'ignore')
                            t_role = 'L' if role == 'label' else ('V' if role == 'var' else 'I')
                            choices[cref] = {'choice': 'T-HEAD', 'anchor': ref,
                                             'table_role': t_role, 'row_class': 'footer'}

                # Mark T-DATA cells (all non-header/footer/skip rows in range)
                start_col_ = meta_dict['start_col']
                end_col_   = meta_dict['end_col']
                _data_vars = meta_dict.get('data_vars') or []
                for rt_row, rt_type in meta_dict.get('row_types', {}).items():
                    if rt_type == 'D':
                        for c_ in range(start_col_, end_col_ + 1):
                            ci     = c_ - start_col_
                            dv     = _data_vars[ci] if ci < len(_data_vars) else {}
                            dv_r   = dv.get('role', 'ignore')
                            t_role = 'V' if dv_r == 'var' else ('L' if dv_r == 'label' else 'I')
                            dref   = _cell_ref(rt_row, c_)
                            if dref not in choices:
                                choices[dref] = {'choice': 'T-DATA', 'anchor': ref,
                                                 'table_role': t_role, 'row_class': 'data'}

                # Mark the table's bottom-right cell as the end ref (◀ glyph in grid)
                _end_ref_mark = _cell_ref(meta_dict['end_row'], meta_dict['end_col'])
                if _end_ref_mark in choices:
                    choices[_end_ref_mark]['is_table_end'] = True

            else:
                # Minimal anchor-only (modal not yet confirmed; placeholder)
                choices[ref] = {
                    'choice': 'T',
                    'name':   table_name,
                    'mult':   mult,
                }

        if note:
            notes[ref] = note
        elif ref in notes and not note:
            # Only delete note if explicitly cleared (empty string sent)
            if 'notes' in fields:
                notes.pop(ref, None)

        # Log classify event
        info = choices.get(ref, {})
        log_detail = f'{ref}:{action} name={info.get("name", "")!r}'
        if note:
            log_detail += f' note={note!r}'
        _STATE['log'].write('CLASSIFY', log_detail)

        # Detect same-name collisions across roles (lbl vs var).
        # lbl: and var: are separate namespaces in the engine — same name is
        # technically allowed — but warn so the author is aware of the overlap.
        name_warning: str | None = None
        new_name = info.get('name', '')
        new_role = action  # 'L' or 'V'
        if new_name and new_role in ('L', 'V'):
            opposite = 'V' if new_role == 'L' else 'L'
            for other_ref, other_meta in choices.items():
                if other_ref == ref:
                    continue
                if other_meta.get('choice') == opposite and other_meta.get('name') == new_name:
                    role_word = 'value' if opposite == 'V' else 'label'
                    name_warning = (
                        f"Name '{new_name}' is also used by {role_word} cell {other_ref}. "
                        f"lbl: and var: are separate namespaces so both will exist in the pattern, "
                        f"but consider using distinct names to avoid ambiguity."
                    )
                    break

        response: dict = {'ok': True, 'ref': ref, 'action': action, 'stats': _build_stats()}
        if name_warning:
            response['name_warning'] = name_warning
        return JSONResponse(response)

    @app.post('/api/classify-batch')
    async def api_classify_batch(request: Request):
        """Classify a rectangular range of cells with a single action.

        Body: {"refs": ["A1","B1",...], "action": "L"|"V"|"C"|"I"|"CLEAR"}
        Returns: {"ok": true, "n_classified": N, "skipped": [...], "stats": {...}}

        Each cell gets auto-derived defaults (name from cell value, type inferred).
        T action is not supported for batch — use the table modal for table anchors.
        """
        body   = await request.json()
        refs   = [r.upper() for r in body.get('refs', []) if isinstance(r, str)]
        action = body.get('action', '')

        if not refs:
            raise HTTPException(400, 'refs list is empty')
        if action not in ('L', 'V', 'I', 'CLEAR'):
            raise HTTPException(400, f'action must be L, V, I, or CLEAR (got {action!r}); '
                                    'T is not supported for batch classification')

        _, merge_skip = _build_merge_info(_STATE['ws'])
        choices: dict = _STATE['choices']
        ws_: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']

        # Push a single undo snapshot covering the whole batch (keyed to first ref)
        if refs:
            _push_undo(refs[0])

        skipped: list[str] = []
        classified: list[str] = []

        for ref in refs:
            if ref in merge_skip:
                skipped.append(ref)
                continue

            try:
                row_, col_ = _parse_ref(ref)
                cell_value = ws_.cell(row=row_, column=col_).value
            except (ValueError, IndexError):
                skipped.append(ref)
                continue
            except Exception:
                skipped.append(ref)
                continue

            if action == 'CLEAR':
                choices.pop(ref, None)
                _STATE['notes'].pop(ref, None)
            elif action == 'L':
                choices[ref] = {
                    'choice':   'L',
                    'name':     _slugify(str(cell_value or '')) + '_label',
                    'ltype':    'string',
                    'lmatch':   str(cell_value or ''),
                    'lbl_mode': '',
                }
            elif action == 'V':
                choices[ref] = {
                    'choice':      'V',
                    'name':        _slugify(str(cell_value or '')),
                    'ftype':       _infer_cell_type(ws_.cell(row=row_, column=col_)) if row_ else 'string',
                    'match':       '.*',
                    'col_a_extra': '',
                }
            elif action == 'I':
                choices[ref] = {'choice': 'I'}

            classified.append(ref)

        _STATE['log'].write('CLASSIFY-BATCH', f'{action} × {len(classified)} cells; skipped {len(skipped)}')
        return JSONResponse({
            'ok':           True,
            'n_classified': len(classified),
            'classified':   classified,
            'skipped':      skipped,
            'stats':        _build_stats(),
        })

    @app.post('/api/undo')
    async def api_undo():
        stack: list = _STATE.get('undo_stack', [])
        if not stack:
            return JSONResponse({'ok': False, 'msg': 'Nothing to undo'})
        snapshot = stack.pop()
        _STATE['choices'] = snapshot['choices']
        _STATE['notes']   = snapshot['notes']
        _STATE['log'].write('UNDO', f"restored {snapshot['ref']}")
        return JSONResponse({'ok': True, 'ref': snapshot['ref'], 'stats': _build_stats()})

    @app.post('/api/note')
    async def api_note(request: Request):
        body = await request.json()
        ref  = body.get('ref', '').upper()
        note = body.get('note', '').strip()
        if not ref:
            raise HTTPException(400, 'ref required')
        if note:
            _STATE['notes'][ref] = note
        else:
            _STATE['notes'].pop(ref, None)
        return JSONResponse({'ok': True})

    @app.get('/api/table-range/{start_ref}/{end_ref}')
    async def api_table_range(start_ref: str, end_ref: str):
        """Return all cell values in the rectangular range start_ref:end_ref.

        Rows are ordered top-to-bottom; each row carries its current choice
        (T / T-HEAD / T-DATA / '') so the modal can pre-assign row types.
        """
        start_ref = start_ref.upper()
        end_ref   = end_ref.upper()
        ws: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']
        try:
            sr, sc = _parse_ref(start_ref)
            er, ec = _parse_ref(end_ref)
        except Exception:
            raise HTTPException(400, f'Invalid refs {start_ref!r}:{end_ref!r}')

        r0, r1 = min(sr, er), max(sr, er)
        c0, c1 = min(sc, ec), max(sc, ec)

        # DoS guard: a large range request (even via a CSRF <img> tag) would
        # spin in a near-infinite loop.  Cap the rectangle before iterating.
        _MAX_TABLE_RANGE_CELLS = 10_000
        row_count = r1 - r0 + 1
        col_count = c1 - c0 + 1
        if row_count * col_count > _MAX_TABLE_RANGE_CELLS:
            raise HTTPException(
                400,
                f'Requested range {row_count}×{col_count} exceeds the '
                f'{_MAX_TABLE_RANGE_CELLS}-cell limit. Select a smaller range.',
            )

        choices = _STATE['choices']

        rows_out = []
        for r in range(r0, r1 + 1):
            cells_out = []
            for c in range(c0, c1 + 1):
                cref = _cell_ref(r, c)
                cell = ws.cell(row=r, column=c)
                ch   = choices.get(cref, {}).get('choice', '')
                cells_out.append({
                    'ref':          cref,
                    'value':        _cell_display(cell.value, 60),
                    'raw':          str(cell.value).strip() if cell.value is not None else '',
                    'inferred_type': _infer_cell_type(cell),
                    'choice':       ch,
                })
            rows_out.append({'sheet_row': r, 'cells': cells_out})

        return JSONResponse({
            'start_ref': start_ref,
            'end_ref':   end_ref,
            'start_row': r0, 'end_row': r1,
            'start_col': c0, 'end_col': c1,
            'col_count': c1 - c0 + 1,
            'row_count': r1 - r0 + 1,
            'rows':      rows_out,
        })

    @app.get('/api/table-scan/{ref}')
    async def api_table_scan(ref: str):
        """Auto-detect table columns starting from the anchor cell.

        For LR direction: scans rightward along the anchor row.
        For TD direction: scans downward along the anchor column.
        Returns suggested column configs (header text, suggested field name & type).
        """
        ref = ref.upper()
        ws: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']
        direction = _STATE['state'].direction
        try:
            anchor_row_, anchor_col_ = _parse_ref(ref)
        except Exception:
            raise HTTPException(404, f'Invalid ref {ref!r}')

        columns = []
        if direction == 'LR':
            c = anchor_col_
            while c <= ws.max_column:
                cell = ws.cell(row=anchor_row_, column=c)
                val  = cell.value
                if val is None and c > anchor_col_:
                    break
                cref = _cell_ref(anchor_row_, c)
                text = str(val).strip() if val is not None else ''
                columns.append({
                    'ref':            cref,
                    'col_letter':     get_column_letter(c),
                    'header':         text,
                    'suggested_field': _slugify(text) if text else f'col_{get_column_letter(c).lower()}',
                    'suggested_type': _infer_cell_type(cell),
                })
                c += 1
        else:
            r = anchor_row_
            while r <= ws.max_row:
                cell = ws.cell(row=r, column=anchor_col_)
                val  = cell.value
                if val is None and r > anchor_row_:
                    break
                cref = _cell_ref(r, anchor_col_)
                text = str(val).strip() if val is not None else ''
                columns.append({
                    'ref':            cref,
                    'row':            r,
                    'header':         text,
                    'suggested_field': _slugify(text) if text else f'row_{r}',
                    'suggested_type': _infer_cell_type(cell),
                })
                r += 1

        # Count non-empty data rows after the header row (preview only)
        data_count = 0
        if direction == 'LR' and columns:
            r = anchor_row_ + 1
            while r <= ws.max_row:
                if ws.cell(row=r, column=anchor_col_).value is None:
                    break
                data_count += 1
                r += 1

        return JSONResponse({
            'anchor':    ref,
            'direction': direction,
            'columns':   columns,
            'data_rows': data_count,
        })

    @app.get('/api/table-context/{ref}')
    async def api_table_context(ref: str):
        """Given any T/T-HEAD/T-DATA cell, return the anchor ref + full table config.

        Returns ``_web_row_configs`` so the table modal can pre-populate for editing.
        """
        ref     = ref.upper()
        choices = _STATE['choices']
        meta    = choices.get(ref, {})
        choice  = meta.get('choice', '')

        if choice == 'T':
            anchor_ref = ref
        elif choice in ('T-HEAD', 'T-DATA'):
            anchor_ref = meta.get('anchor', '')
        else:
            raise HTTPException(404, f'{ref} is not a table cell (choice={choice!r})')

        anchor_meta = choices.get(anchor_ref, {})
        return JSONResponse({
            'anchor':          anchor_ref,
            'name':            anchor_meta.get('name', ''),
            'mult':            anchor_meta.get('mult', '*'),
            'end_ref':         anchor_meta.get('end_ref', anchor_ref),
            'start_ref':       anchor_meta.get('start_ref', anchor_ref),
            # Full row configs for modal edit mode
            '_web_row_configs': anchor_meta.get('_web_row_configs', []),
            # Legacy columns (backward compat)
            'columns':         anchor_meta.get('columns', []),
            'header_rows':     anchor_meta.get('header_rows', []),
            'data_vars':       anchor_meta.get('data_vars', []),
            'footer_rows':     anchor_meta.get('footer_rows', []),
        })

    def _schema_summary() -> str:
        """Return a single-line summary of the current choices for SCHEMA_SUMMARY log events."""
        ch = _STATE['choices']
        lbl_count = sum(1 for v in ch.values() if v.get('choice') == 'L')
        var_count = sum(1 for v in ch.values() if v.get('choice') == 'V')
        # Collect all modifiers (col_a_extra tokens) across var fields
        from collections import Counter
        mod_counts: Counter = Counter()
        for v in ch.values():
            if v.get('choice') == 'V':
                extra = v.get('col_a_extra', '')
                if extra:
                    for tok in extra.split(':'):
                        if tok:
                            mod_counts[tok] += 1
        mods_str = ' '.join(f'{k}×{n}' for k, n in sorted(mod_counts.items())) or 'none'
        return (
            f'lbl={lbl_count} var={var_count} total={lbl_count + var_count} '
            f'mods={mods_str}'
        )

    def _make_csv() -> str:
        """Generate pattern CSV from current session state."""
        st: WizardState = _STATE['state']
        ws_     = _STATE['ws']
        cells   = _build_cell_order(ws_, st.direction)
        return _choices_to_csv(
            ws_,
            _STATE['choices'],
            cells,
            st.direction,
            st.sheet_name or ws_.title,
            ignore_case_labels=st.ignore_case_labels,
            ignore_case_values=st.ignore_case_values,
            currency_sign=st.currency_sign,
            trim_whitespace_labels=st.trim_whitespace_labels,
            trim_whitespace_values=st.trim_whitespace_values,
            lbl_match=st.lbl_match,
            var_match=st.var_match,
            empty_aliases=st.empty_aliases,
        )

    @app.get('/api/preview')
    async def api_preview():
        """Return the current pattern as CSV text."""
        try:
            return PlainTextResponse(_make_csv(), media_type='text/plain')
        except Exception as exc:
            raise HTTPException(500, str(exc))

    def _default_save_stem() -> str:
        """Return '{data_stem}_{safe_sheet_name}' for use in pattern filenames."""
        xlsx_path_ = Path(_STATE['xlsx_path'])
        sheet_name = _STATE['ws'].title
        safe_sheet = _slugify(sheet_name) or 'sheet'
        return f'{xlsx_path_.stem}_{safe_sheet}'

    @app.post('/api/save')
    async def api_save(request: Request):
        """Generate and save the pattern file next to the input xlsx."""
        body       = await request.json()
        xlsx_path_ = Path(_STATE['xlsx_path'])

        # Determine output path — sheet name is embedded in the default filename
        out_name = Path(body.get('filename', '')).name  # strip any directory components
        if not out_name:
            out_name = _default_save_stem() + '_pattern-from-web.csv'
        out_path = xlsx_path_.parent / out_name

        # Overwrite guard: return exists=True so the browser can confirm
        if out_path.exists() and not body.get('confirm_overwrite'):
            return JSONResponse({
                'ok':     False,
                'exists': True,
                'path':   str(out_path),
            })

        try:
            csv_text = _make_csv()
            out_path.write_text(csv_text, encoding='utf-8')
        except Exception as exc:
            raise HTTPException(500, str(exc))

        _STATE['log'].write('SCHEMA_SUMMARY', _schema_summary())
        _STATE['log'].write('SAVE', f'path={out_path} rows={len(csv_text.splitlines())}')
        return JSONResponse({
            'ok':       True,
            'saved_to': str(out_path),
            'rows':     len(csv_text.splitlines()),
        })

    @app.post('/api/save-xlsx')
    async def api_save_xlsx(request: Request):
        """Generate and save the pattern as a coloured .xlsx file."""
        import csv as _csv
        import io as _io
        body       = await request.json()
        xlsx_path_ = Path(_STATE['xlsx_path'])

        out_name = Path(body.get('filename', '')).name  # strip any directory components
        if not out_name:
            out_name = _default_save_stem() + '_pattern-from-web.xlsx'
        out_path = xlsx_path_.parent / out_name

        # Overwrite guard
        if out_path.exists() and not body.get('confirm_overwrite'):
            return JSONResponse({
                'ok':     False,
                'exists': True,
                'path':   str(out_path),
            })

        try:
            csv_text = _make_csv()
            # Parse CSV rows
            rows = list(_csv.reader(_io.StringIO(csv_text)))
            # Write xlsx
            out_wb = openpyxl.Workbook()
            out_ws = out_wb.active
            out_ws.title = 'pattern'
            for row in rows:
                out_ws.append(row)
            out_wb.save(str(out_path))
            # Colourize
            try:
                from .pattern_colors import colorize_pattern_file
                colorize_pattern_file(str(out_path))
            except Exception:
                pass  # colour is cosmetic — don't fail on it
        except Exception as exc:
            raise HTTPException(500, str(exc))

        _STATE['log'].write('SCHEMA_SUMMARY', _schema_summary())
        _STATE['log'].write('SAVE', f'path={out_path} format=xlsx rows={len(rows)}')
        return JSONResponse({
            'ok':       True,
            'saved_to': str(out_path),
            'rows':     len(rows),
        })

    @app.get('/api/cell/{ref}')
    async def api_cell(ref: str):
        ref = ref.upper()
        ws  = _STATE['ws']
        try:
            row, col = _parse_ref(ref)
            cell     = ws.cell(row=row, column=col)
        except Exception:
            raise HTTPException(404, f'Invalid cell ref {ref!r}')
        choice_info = _STATE['choices'].get(ref, {})
        ex_col_a_extra = choice_info.get('col_a_extra', '')
        ex_var_mode, ex_modifiers = _col_a_extra_to_parts(ex_col_a_extra)
        _all_img = _STATE.get('image_cells', {})
        img_info = _all_img.get(ws.title, {}).get(ref)
        return JSONResponse({
            'ref':          ref,
            'row':          row,
            'col':          col,
            'col_letter':   get_column_letter(col),
            'value':        _cell_display(cell.value, 200),
            'raw':          str(cell.value) if cell.value is not None else '',
            'has_image':    img_info is not None,
            'image_count':  img_info['count'] if img_info else 0,
            'image_suspicious': img_info.get('suspicious', False) if img_info else False,
            'image_mimes':  img_info.get('mimes', []) if img_info else [],
            'inferred_type': 'image' if img_info is not None else _infer_cell_type(cell),
            'choice':       choice_info.get('choice', ''),
            'anchor':       choice_info.get('anchor', ''),   # set for T-HEAD / T-DATA
            'name':         choice_info.get('name', ''),
            'ltype':        choice_info.get('ltype', ''),
            'lmatch':       choice_info.get('lmatch', ''),
            'lbl_mode':     choice_info.get('lbl_mode', ''),
            'ftype':        choice_info.get('ftype', ''),
            'match':        choice_info.get('match', ''),
            'var_mode':     ex_var_mode,
            'modifiers':    ex_modifiers,
            'note':         _STATE['notes'].get(ref, ''),
            'has_preview':  ref in _STATE.get('extracted_images', {}),
        })

    @app.get('/api/image/{ref}')
    async def api_image(ref: str):
        """Serve an extracted embedded image for the given cell reference."""
        from fastapi.responses import FileResponse
        ref = ref.upper()
        extracted = _STATE.get('extracted_images', {})
        img_path = extracted.get(ref)
        if not img_path or not Path(img_path).is_file():
            raise HTTPException(404, 'No extracted image for this cell')
        # Infer content type from extension
        ext = Path(img_path).suffix.lower()
        ct_map = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                  '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp',
                  '.tiff': 'image/tiff', '.svg': 'image/svg+xml'}
        media_type = ct_map.get(ext, 'application/octet-stream')
        headers: dict[str, str] = {'X-Content-Type-Options': 'nosniff'}
        if ext == '.svg':
            # SVG can carry inline scripts; sandbox it so navigating directly
            # to the endpoint cannot execute scripts in the wizard's origin.
            headers['Content-Security-Policy'] = "default-src 'none'; sandbox"
        return FileResponse(img_path, media_type=media_type, headers=headers)

    @app.post('/api/extract')
    async def api_extract():
        """Run extraction with the current pattern; return result + approximate provenance.

        If no pattern file was pre-loaded (-p flag), generates the pattern
        from the current wizard choices on the fly.
        """
        xlsx_path_ = _STATE.get('xlsx_path')
        tmp_csv_path: str | None = None
        try:
            import grepxcel as _gx
            import tempfile

            pattern_path_ = _STATE.get('pattern_path')
            if pattern_path_:
                result = _gx.extract(pattern_path_, xlsx_path_, output_format='nested')
            else:
                # No pre-loaded pattern — generate CSV from current choices
                csv_text = _make_csv()
                # Check that the CSV has at least one extraction instruction
                start_idx = csv_text.find('START:')
                end_idx   = csv_text.find('END:')
                has_steps = (
                    start_idx >= 0 and end_idx > start_idx
                    and csv_text[start_idx + 6: end_idx].strip()
                )
                if not has_steps:
                    return JSONResponse({
                        'ok': False,
                        'error': 'No fields classified yet — classify some cells first, then run extraction.',
                    })
                with tempfile.NamedTemporaryFile(
                    mode='w', suffix='.csv', delete=False, encoding='utf-8'
                ) as tf:
                    tf.write(csv_text)
                    tmp_csv_path = tf.name
                result = _gx.extract(tmp_csv_path, xlsx_path_, output_format='nested')
            safe = _to_json_safe(result)
        except Exception as exc:
            _STATE['log'].write('EXTRACT', f'ok=False error={exc}')
            return JSONResponse({'ok': False, 'error': str(exc)})
        finally:
            if tmp_csv_path:
                import os as _os
                try:
                    _os.unlink(tmp_csv_path)
                except OSError:
                    pass

        # Provenance: field name → [cell refs] derived from the current choices dict.
        # Using the leaf name (what the user typed in the classify form) as the key,
        # which matches the leaf key in the nested extraction result.
        provenance: dict[str, list[str]] = {}
        for ref, info in _STATE.get('choices', {}).items():
            name = info.get('name', '')
            if name:
                provenance.setdefault(name, []).append(ref)

        _STATE['log'].write('EXTRACT', f'ok=True fields={len(provenance)}')
        return JSONResponse({'ok': True, 'result': safe, 'provenance': provenance})

    @app.post('/api/load-pattern')
    async def api_load_pattern(request: Request):
        """Upload a pattern file and reload the wizard's preloaded choices.

        Accepts multipart form data with a single ``file`` field.  The uploaded
        pattern (.xlsx or .csv) is saved to a temp file, preloaded against the
        current data sheet, and the session choices are replaced.  The browser
        should reload after a successful response to re-render the updated grid.
        """
        import tempfile
        import os as _os
        from starlette.datastructures import UploadFile as _UploadFile

        form = await request.form()
        upload = form.get('file')
        if upload is None or not hasattr(upload, 'read'):
            raise HTTPException(400, 'No file uploaded')

        filename = getattr(upload, 'filename', 'pattern.csv') or 'pattern.csv'
        suffix = '.xlsx' if filename.lower().endswith('.xlsx') else '.csv'

        tmp_path: str | None = None
        try:
            data = await upload.read()
            _MAX_PATTERN_BYTES = 5 * 1024 * 1024  # 5 MB — same as validate_file default
            if len(data) > _MAX_PATTERN_BYTES:
                raise HTTPException(413, 'Pattern file exceeds 5 MB limit')
            with tempfile.NamedTemporaryFile(
                suffix=suffix, delete=False
            ) as tf:
                tf.write(data)
                tmp_path = tf.name

            from .security import validate_pattern_file
            validate_pattern_file(tmp_path)
            ws_ = _STATE['ws']
            choices, cfg, warnings = _preload_from_pattern(ws_, tmp_path)

            # Rebuild state from preloaded choices so Config tab reflects the pattern
            _STATE['choices'] = choices
            _STATE['preload_warnings'] = warnings

            # Update direction/config from the loaded pattern's global config
            st: WizardState = _STATE['state']
            st.direction              = cfg.get('direction', st.direction)
            st.ignore_case_labels     = cfg.get('ignore_case_labels', st.ignore_case_labels)
            st.ignore_case_values     = cfg.get('ignore_case_values', st.ignore_case_values)
            st.trim_whitespace_labels = cfg.get('trim_whitespace_labels', st.trim_whitespace_labels)
            st.trim_whitespace_values = cfg.get('trim_whitespace_values', st.trim_whitespace_values)
            st.currency_sign          = cfg.get('currency_sign', st.currency_sign)
            st.lbl_match              = cfg.get('lbl_match', st.lbl_match)
            st.var_match              = cfg.get('var_match', st.var_match)
            st.empty_aliases          = cfg.get('empty_aliases', st.empty_aliases)

            # Mark the loaded pattern as the active pattern for future extractions
            _STATE['pattern_path'] = tmp_path   # keep temp alive for this session

            _STATE['log'].write(
                'LOAD_PATTERN',
                f'file={filename} choices={len(choices)} warnings={len(warnings)}',
            )

            return JSONResponse({
                'ok': True,
                'n_choices': len(choices),
                'warnings': warnings,
            })
        except Exception as exc:
            if tmp_path:
                try:
                    _os.unlink(tmp_path)
                except OSError:
                    pass
            raise HTTPException(500, str(exc))

    @app.post('/api/load-pattern-by-path')
    async def api_load_pattern_by_path(request: Request):
        """Load a pattern from a server-side path (from /api/list-patterns listing)."""
        body = await request.json()
        path_str = body.get('path', '')
        if not path_str:
            raise HTTPException(400, 'Missing path')

        pat_path = Path(path_str)
        # Security: path must be in the same directory as the data file
        data_dir = Path(_STATE['xlsx_path']).parent.resolve()
        try:
            pat_resolved = pat_path.resolve()
        except Exception:
            raise HTTPException(400, 'Invalid path')
        if pat_resolved.parent != data_dir:
            raise HTTPException(403, 'Pattern path must be in the same folder as the data file')
        if not pat_resolved.exists():
            raise HTTPException(404, f'Pattern file not found: {pat_resolved.name}')

        try:
            from .security import validate_pattern_file
            validate_pattern_file(str(pat_resolved))
            ws_ = _STATE['ws']
            choices, cfg, warnings = _preload_from_pattern(ws_, str(pat_resolved))

            _STATE['choices'] = choices
            _STATE['preload_warnings'] = warnings

            st: WizardState = _STATE['state']
            st.direction              = cfg.get('direction', st.direction)
            st.ignore_case_labels     = cfg.get('ignore_case_labels', st.ignore_case_labels)
            st.ignore_case_values     = cfg.get('ignore_case_values', st.ignore_case_values)
            st.trim_whitespace_labels = cfg.get('trim_whitespace_labels', st.trim_whitespace_labels)
            st.trim_whitespace_values = cfg.get('trim_whitespace_values', st.trim_whitespace_values)
            st.currency_sign          = cfg.get('currency_sign', st.currency_sign)
            st.lbl_match              = cfg.get('lbl_match', st.lbl_match)
            st.var_match              = cfg.get('var_match', st.var_match)
            st.empty_aliases          = cfg.get('empty_aliases', st.empty_aliases)
            _STATE['pattern_path'] = str(pat_resolved)

            _STATE['log'].write(
                'LOAD_PATTERN',
                f'file={pat_resolved.name} choices={len(choices)} warnings={len(warnings)}',
            )
            return JSONResponse({
                'ok': True,
                'n_choices': len(choices),
                'warnings': warnings,
            })
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(500, str(exc))

    @app.post('/api/shutdown')
    async def api_shutdown(request: Request):
        """Graceful shutdown — called by the browser when user clicks 'Done'."""
        # CSRF guard: form POSTs are simple requests (no preflight). Reject any
        # cross-origin Origin header that is not localhost / 127.0.0.1.
        origin = request.headers.get('origin', '')
        if origin and not (
            origin.startswith('http://localhost:')
            or origin.startswith('http://127.0.0.1:')
        ):
            raise HTTPException(403, 'Cross-origin shutdown rejected')
        # Clean up any uploaded temp pattern file before exiting.
        import tempfile as _tmpmod
        _tmp_pat = _STATE.get('pattern_path')
        if _tmp_pat and _tmp_pat.startswith(_tmpmod.gettempdir()):
            try:
                os.unlink(_tmp_pat)
            except OSError:
                pass
        _STATE['log'].close(_STATE.get('choices', {}), _STATE.get('notes', {}))
        def _stop():
            time.sleep(0.3)
            os._exit(0)
        threading.Thread(target=_stop, daemon=True).start()
        return JSONResponse({'ok': True})

    return app


def _get_version() -> str:
    try:
        from grepxcel import __version__
        return __version__
    except Exception:
        return '?'


def _port_in_use(port: int) -> bool:
    """Return True if *port* is already bound on localhost."""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex(('127.0.0.1', port)) == 0


def _pids_on_port(port: int) -> list[tuple[int, str]]:
    """Return [(pid, cmdline), …] for processes listening on *port*.

    Uses ``ss`` (Linux) then ``lsof`` (macOS/fallback), then falls back
    gracefully to an empty list when neither is available.
    """
    import subprocess, shutil
    results: list[tuple[int, str]] = []

    # ── Try ss (Linux iproute2) ────────────────────────────────────────────────
    if shutil.which('ss'):
        try:
            out = subprocess.check_output(
                ['ss', '-tlnp', f'sport = :{port}'],
                text=True, stderr=subprocess.DEVNULL,
            )
            for line in out.splitlines():
                # ss output: "LISTEN  0  128  127.0.0.1:8765  *:*  users:(("grepxcel",pid=12345,fd=7))"
                if f':{port}' in line and 'pid=' in line:
                    import re
                    for m in re.finditer(r'pid=(\d+)', line):
                        pid = int(m.group(1))
                        try:
                            cmd = Path(f'/proc/{pid}/cmdline').read_text().replace('\x00', ' ').strip()
                        except OSError:
                            cmd = f'pid {pid}'
                        results.append((pid, cmd))
        except (subprocess.SubprocessError, FileNotFoundError):
            pass

    # ── Try lsof (macOS / BSD) ─────────────────────────────────────────────────
    if not results and shutil.which('lsof'):
        try:
            out = subprocess.check_output(
                ['lsof', '-ti', f'tcp:{port}'],
                text=True, stderr=subprocess.DEVNULL,
            )
            for line in out.splitlines():
                pid = int(line.strip())
                try:
                    cmd = Path(f'/proc/{pid}/cmdline').read_text().replace('\x00', ' ').strip()
                except OSError:
                    try:
                        cmd = subprocess.check_output(
                            ['ps', '-p', str(pid), '-o', 'command='],
                            text=True, stderr=subprocess.DEVNULL,
                        ).strip()
                    except subprocess.SubprocessError:
                        cmd = f'pid {pid}'
                results.append((pid, cmd))
        except (subprocess.SubprocessError, ValueError, FileNotFoundError):
            pass

    return results


def _handle_port_conflict(port: int) -> None:
    """If *port* is taken, show who holds it and ask the user to kill or abort."""
    import signal
    if not _port_in_use(port):
        return  # all clear

    pids = _pids_on_port(port)
    print(f'\n⚠️  Port {port} is already in use.', file=sys.stderr)
    if pids:
        for pid, cmd in pids:
            short = cmd[:80] + ('…' if len(cmd) > 80 else '')
            print(f'   PID {pid}: {short}', file=sys.stderr)
    else:
        print('   (Could not identify the process holding the port.)', file=sys.stderr)

    try:
        answer = input(
            f'\nKill the existing process and restart on port {port}? [y/N] '
        ).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print('\nAborted.', file=sys.stderr)
        sys.exit(1)

    if answer not in ('y', 'yes'):
        print(
            f'Tip: use --port <other> to start on a different port.',
            file=sys.stderr,
        )
        sys.exit(1)

    # Kill each identified process
    if pids:
        for pid, _ in pids:
            try:
                os.kill(pid, signal.SIGTERM)
                print(f'   Sent SIGTERM to PID {pid}', file=sys.stderr)
            except ProcessLookupError:
                pass
        # Wait for port to free (up to 3 s)
        for _ in range(30):
            time.sleep(0.1)
            if not _port_in_use(port):
                break
        else:
            # Hard-kill if still up
            for pid, _ in pids:
                try:
                    os.kill(pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
            time.sleep(0.3)
    else:
        # No PID found; wait a moment hoping the OS frees the port
        print('   Waiting for port to free…', file=sys.stderr)
        time.sleep(2)

    if _port_in_use(port):
        print(
            f'Port {port} is still occupied. Try --port <other>.',
            file=sys.stderr,
        )
        sys.exit(1)

    print(f'   Port {port} is now free.\n', file=sys.stderr)


def run(
    xlsx_path: str,
    pattern_path: str | None = None,
    port: int = 8765,
    open_browser: bool = True,
    max_rows: int = 150,
    max_cols: int = 40,
    sheet: str | None = None,
    max_file_mb: float = 5,
    max_uncompressed_mb: float = 50,
) -> None:
    """Start the web wizard server and (optionally) open the browser."""
    if not _WEB_OK:
        print(
            'Web wizard requires FastAPI and uvicorn.\n'
            'Install with:  pip install "grepxcel[web]"',
            file=sys.stderr,
        )
        sys.exit(1)

    _handle_port_conflict(port)

    app = create_app(
        xlsx_path,
        pattern_path,
        max_rows=max_rows,
        max_cols=max_cols,
        sheet=sheet,
        max_file_mb=max_file_mb,
        max_uncompressed_mb=max_uncompressed_mb,
    )
    url = f'http://localhost:{port}'

    if open_browser:
        def _open():
            time.sleep(0.8)
            webbrowser.open(url)
        threading.Thread(target=_open, daemon=True).start()

    log_path = (_STATE.get('log') or type('', (), {'path': None})()).path
    print(f'\ngrepxcel Web Wizard')
    print(f'  File  : {xlsx_path}')
    if pattern_path:
        print(f'  Pattern: {pattern_path}')
    print(f'  URL   : {url}')
    if log_path:
        print(f'  Log   : {log_path}')
    print(f'\n  Press Ctrl+C to stop.\n')

    uvicorn.run(app, host='127.0.0.1', port=port, log_level='warning')
