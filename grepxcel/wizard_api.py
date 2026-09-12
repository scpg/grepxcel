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
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

# ── Reuse existing pure functions ─────────────────────────────────────────────
from grepxcel.wizard import WizardState
from grepxcel.wizard_tui import (
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
            self._fh.write('# CLASSIFY ref:ACTION name=… — cell classified\n')
            self._fh.write('# CONFIG   direction=… — global config changed\n')
            self._fh.write('# UNDO     ref — last classification reversed\n')
            self._fh.write('# SAVE     path — pattern file written to disk\n')
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
    s = str(value)
    if len(s) > max_len:
        return s[:max_len] + '…'
    return s


def _cell_type_display(value: Any) -> str:
    if value is None:
        return 'empty'
    if isinstance(value, bool):
        return 'boolean'
    if isinstance(value, (int, float)):
        return 'number'
    if isinstance(value, datetime):
        return 'date'
    return 'string'


def _build_sheet_data() -> dict:
    """Serialize the active worksheet into a JSON-friendly structure."""
    ws: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']
    choices: dict = _STATE['choices']
    notes: dict = _STATE['notes']
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    # Cap to reasonable display size
    display_rows = min(max_row, _STATE.get('max_rows', 150))
    display_cols = min(max_col, _STATE.get('max_cols', 40))

    rows = []
    for r in range(1, display_rows + 1):
        row = []
        for c in range(1, display_cols + 1):
            ref = _cell_ref(r, c)
            cell = ws.cell(row=r, column=c)
            choice_info = choices.get(ref, {})
            row.append({
                'ref':     ref,
                'row':     r,
                'col':     c,
                'col_letter': get_column_letter(c),
                'value':   _cell_display(cell.value),
                'raw':     str(cell.value) if cell.value is not None else '',
                'type':    _cell_type_display(cell.value),
                'choice':  choice_info.get('choice', ''),
                'name':    choice_info.get('name', ''),
                'note':    notes.get(ref, ''),
                'empty':   cell.value is None,
            })
        rows.append(row)

    col_letters = [get_column_letter(c) for c in range(1, display_cols + 1)]

    return {
        'rows':        rows,
        'max_row':     display_rows,
        'max_col':     display_cols,
        'col_letters': col_letters,
        'total_rows':  max_row,
        'total_cols':  max_col,
        'choices':     choices,
        'notes':       notes,
    }


def _build_stats() -> dict:
    choices = _STATE.get('choices', {})
    counts: dict[str, int] = {}
    for info in choices.values():
        ch = info.get('choice', '')
        counts[ch] = counts.get(ch, 0) + 1
    ws = _STATE.get('ws')
    total = sum(
        1 for r in range(1, (ws.max_row or 1) + 1)
        for c in range(1, (ws.max_column or 1) + 1)
        if ws.cell(row=r, column=c).value is not None
    ) if ws else 0
    return {
        'L': counts.get('L', 0),
        'V': counts.get('V', 0),
        'C': counts.get('C', 0),
        'T': counts.get('T', 0),
        'I': counts.get('I', 0),
        'total': total,
        'classified': sum(counts.values()),
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
) -> 'FastAPI':
    """Create and return the FastAPI application for the web wizard."""
    if not _WEB_OK:
        raise ImportError(
            'Web wizard requires FastAPI and uvicorn. '
            'Install with: pip install "grepxcel[web]"'
        )

    # ── Load workbook ─────────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    # ── Pre-populate from existing pattern ────────────────────────────────────
    state = WizardState(sheet_name=ws.title)
    choices: dict[str, dict] = {}
    notes: dict[str, str] = {}

    if pattern_path and Path(pattern_path).exists():
        try:
            loaded_choices, preload_cfg, _warnings = _preload_from_pattern(ws, pattern_path)
            choices.update(loaded_choices)
            if preload_cfg.get('direction'):
                state.direction       = preload_cfg['direction']
            if preload_cfg.get('lbl_match'):
                state.lbl_match       = preload_cfg['lbl_match']
            if preload_cfg.get('var_match'):
                state.var_match       = preload_cfg['var_match']
            if preload_cfg.get('ignore_case') is not None:
                state.ignore_case     = preload_cfg['ignore_case']
            if preload_cfg.get('trim_whitespace') is not None:
                state.trim_whitespace = preload_cfg['trim_whitespace']
            if preload_cfg.get('currency_sign'):
                state.currency_sign   = preload_cfg['currency_sign']
            if preload_cfg.get('empty_aliases'):
                state.empty_aliases   = preload_cfg['empty_aliases']
        except Exception:
            pass

    # ── Session log ───────────────────────────────────────────────────────────
    session_log = _SessionLog(xlsx_path, ws.title)

    # ── Module-level session ──────────────────────────────────────────────────
    _STATE.clear()
    _STATE.update({
        'xlsx_path':   xlsx_path,
        'pattern_path': pattern_path,
        'wb':          wb,
        'ws':          ws,
        'state':       state,
        'choices':     choices,
        'notes':       notes,
        'undo_stack':  [],
        'max_rows':    max_rows,
        'max_cols':    max_cols,
        'log':         session_log,
    })

    # Log initial config
    st0 = state
    session_log.write('CONFIG',
        f'direction={st0.direction} ignore_case={st0.ignore_case} '
        f'trim_whitespace={st0.trim_whitespace} currency_sign={st0.currency_sign!r} '
        f'lbl_match={st0.lbl_match!r} var_match={st0.var_match!r} '
        f'aliases={st0.empty_aliases}'
    )
    if pattern_path:
        session_log.write('PRELOAD',
            f'pattern={pattern_path} cells_loaded={len(choices)}'
        )

    # ── Jinja2 env ────────────────────────────────────────────────────────────
    tpl_dir = Path(__file__).parent / 'templates'
    jinja_env = Environment(loader=FileSystemLoader(str(tpl_dir)),
                            autoescape=True)

    # ── FastAPI app ───────────────────────────────────────────────────────────
    app = FastAPI(title='grepxcel Web Wizard', docs_url=None, redoc_url=None)

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

    @app.get('/api/state')
    async def api_state():
        st: WizardState = _STATE['state']
        return JSONResponse({
            'config': {
                'direction':      st.direction,
                'ignore_case':    st.ignore_case,
                'trim_whitespace': st.trim_whitespace,
                'currency_sign':  st.currency_sign,
                'lbl_match':      st.lbl_match,
                'var_match':      st.var_match,
                'empty_aliases':  st.empty_aliases,
            },
            'choices': _STATE['choices'],
            'notes':   _STATE['notes'],
            'stats':   _build_stats(),
        })

    @app.post('/api/config')
    async def api_config(request: Request):
        body = await request.json()
        st: WizardState = _STATE['state']
        st.direction       = body.get('direction',       st.direction)
        st.ignore_case     = body.get('ignore_case',     st.ignore_case)
        st.trim_whitespace = body.get('trim_whitespace', st.trim_whitespace)
        st.currency_sign   = body.get('currency_sign',   st.currency_sign)
        st.lbl_match       = body.get('lbl_match',       st.lbl_match)
        st.var_match       = body.get('var_match',       st.var_match)
        st.empty_aliases   = body.get('empty_aliases',   st.empty_aliases)
        _STATE['log'].write('CONFIG',
            f'direction={st.direction} ignore_case={st.ignore_case} '
            f'trim_whitespace={st.trim_whitespace} currency_sign={st.currency_sign!r} '
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

        if not ref or action not in ('L', 'V', 'C', 'I', 'T', 'CLEAR'):
            raise HTTPException(400, f'Invalid ref={ref!r} or action={action!r}')

        _push_undo(ref)
        choices = _STATE['choices']
        notes   = _STATE['notes']

        if action == 'CLEAR':
            choices.pop(ref, None)
            notes.pop(ref, None)
            _STATE['log'].write('CLASSIFY', f'{ref}:CLEAR')
            return JSONResponse({'ok': True, 'ref': ref, 'action': 'CLEAR'})

        ws: openpyxl.worksheet.worksheet.Worksheet = _STATE['ws']
        try:
            row, col = _parse_ref(ref)
            cell_value = ws.cell(row=row, column=col).value
        except Exception:
            cell_value = None

        if action == 'L':
            choices[ref] = {
                'choice':   'L',
                'name':     fields.get('name', _slugify(str(cell_value or '')) + '_label'),
                'ltype':    fields.get('type', 'string'),
                'lmatch':   fields.get('match', str(cell_value or '')),
                'lbl_mode': fields.get('match_mode', ''),
            }
        elif action == 'V':
            var_mode_raw  = fields.get('match_mode', '(default)')
            modifiers_raw = fields.get('modifiers', 'none')
            col_a_extra   = _col_a_extra_from_parts(var_mode_raw, modifiers_raw)
            choices[ref] = {
                'choice':      'V',
                'name':        fields.get('name', _slugify(str(cell_value or ''))),
                'ftype':       fields.get('type', _infer_cell_type(ws.cell(row=row, column=col))),
                'match':       fields.get('match', '.*'),
                'col_a_extra': col_a_extra,
            }
        elif action == 'C':
            choices[ref] = {
                'choice': 'C',
                'name':   fields.get('name', _slugify(str(cell_value or ''))),
            }
        elif action == 'I':
            choices[ref] = {'choice': 'I'}
        elif action == 'T':
            # Simplified table: mark anchor cell as T; full multi-step is TUI-only in v1
            choices[ref] = {
                'choice': 'T',
                'name':   fields.get('name', _slugify(str(cell_value or '')) + '_table'),
                '_web_simplified': True,
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

        return JSONResponse({
            'ok':     True,
            'ref':    ref,
            'action': action,
            'stats':  _build_stats(),
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
            ignore_case=st.ignore_case,
            currency_sign=st.currency_sign,
            trim_whitespace=st.trim_whitespace,
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

    @app.post('/api/save')
    async def api_save(request: Request):
        """Generate and save the pattern file next to the input xlsx."""
        body       = await request.json()
        xlsx_path_ = Path(_STATE['xlsx_path'])

        # Determine output path
        out_name = body.get('filename', '')
        if not out_name:
            stem     = xlsx_path_.stem
            out_name = stem + '_pattern-from-web.csv'
        out_path = xlsx_path_.parent / out_name

        try:
            csv_text = _make_csv()
            out_path.write_text(csv_text, encoding='utf-8')
        except Exception as exc:
            raise HTTPException(500, str(exc))

        _STATE['log'].write('SAVE', f'path={out_path} rows={len(csv_text.splitlines())}')
        return JSONResponse({
            'ok':       True,
            'saved_to': str(out_path),
            'rows':     len(csv_text.splitlines()),
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
        return JSONResponse({
            'ref':          ref,
            'row':          row,
            'col':          col,
            'col_letter':   get_column_letter(col),
            'value':        _cell_display(cell.value, 200),
            'raw':          str(cell.value) if cell.value is not None else '',
            'inferred_type': _infer_cell_type(cell),
            'choice':       choice_info.get('choice', ''),
            'name':         choice_info.get('name', ''),
            'ltype':        choice_info.get('ltype', ''),
            'lmatch':       choice_info.get('lmatch', ''),
            'lbl_mode':     choice_info.get('lbl_mode', ''),
            'ftype':        choice_info.get('ftype', ''),
            'match':        choice_info.get('match', ''),
            'var_mode':     ex_var_mode,
            'modifiers':    ex_modifiers,
            'note':         _STATE['notes'].get(ref, ''),
        })

    @app.get('/api/shutdown')
    async def api_shutdown():
        """Graceful shutdown — called by the browser when user clicks 'Done'."""
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


def run(
    xlsx_path: str,
    pattern_path: str | None = None,
    port: int = 8765,
    open_browser: bool = True,
    max_rows: int = 150,
    max_cols: int = 40,
) -> None:
    """Start the web wizard server and (optionally) open the browser."""
    if not _WEB_OK:
        print(
            'Web wizard requires FastAPI and uvicorn.\n'
            'Install with:  pip install "grepxcel[web]"',
            file=sys.stderr,
        )
        sys.exit(1)

    app = create_app(xlsx_path, pattern_path, max_rows=max_rows, max_cols=max_cols)
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
