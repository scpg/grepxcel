"""Full-screen TUI wizard for grepxcel.

Requires textual: pip install 'grepxcel[wizard]'

Design principles
-----------------
* push_screen(modal, callback) throughout — no push_screen_wait, no workers needed.
* _choices dict is the single source of truth; WizardState is built on save.
* Panel has 4 fixed zones: CELL · CLASSIFY · NAVIGATE · LEGEND+STATS
* Color system: green=Label  yellow=Value  blue=Header  magenta=Table  dim=Ignore
* Highlights: H key marks all unclassified non-empty cells in amber; auto-clears on
  any navigation or classification.
* Undo: Ctrl+Z / undo stack pops the last classification and reverts _choices entry.
* Web-service ready: WizardState is JSON-serialisable via dataclasses.asdict().
"""
from __future__ import annotations

import csv
import datetime
import io
import os
import sys
from typing import Any

try:
    from textual.app import App, ComposeResult
    from textual.binding import Binding
    from textual.containers import Horizontal, Vertical
    from textual.screen import ModalScreen
    from textual.widgets import DataTable, Footer, Header, Input, Label, Select, Static
    from rich.text import Text as RichText
    _TEXTUAL_OK = True
except ImportError:
    _TEXTUAL_OK = False

from .wizard import (
    WizardState,
    _slugify,
    _propose_type,
    _write_pattern,
    _build_cell_order,
    _cell_ref,
    _col_label,
    _parse_cell_ref,
    _find_next_nonempty,
    _find_prev_nonempty,
    _save_history,
    _load_history,
    _detect_template,
)


# ── Style constants ────────────────────────────────────────────────────────────

_STYLE: dict[str, str] = {
    'L':      'bold green',
    'C':      'bold blue',
    'V':      'bold yellow',
    'T':      'bold magenta',
    'T-HEAD': 'magenta',          # non-anchor column header cells
    'I':      'dim',
    'PENDING': 'bold black on dark_goldenrod',
}

_CHOICE_COLOR = {
    'L':      'green',
    'C':      'blue',
    'V':      'yellow',
    'T':      'magenta',
    'T-HEAD': 'magenta',
    'I':      'dim',
}

_CHOICE_NAME = {
    'L':      'Label',
    'C':      'Header',
    'V':      'Value',
    'T':      'Table anchor',
    'T-HEAD': 'Table column',
    'I':      'Ignore',
}

_SEP = '─' * 42   # visual divider for panel zones


def _trunc(value: Any, width: int = 20) -> str:
    """Truncate to width chars, adding … when text is cut."""
    if value is None:
        return ''
    s = str(value)
    return (s[:width - 1] + '…') if len(s) > width else s


def _styled(value: Any, choice: str) -> 'RichText':
    if value is None:
        # Empty classified cell: show a colored dot so the user can see it was classified.
        # An empty RichText with a bold style renders as nothing — the dot is essential.
        marker = '○' if choice == 'I' else '●'
        return RichText(marker, style=_STYLE.get(choice, ''))
    return RichText(_trunc(value), style=_STYLE.get(choice, ''))


# ── Pattern builder ────────────────────────────────────────────────────────────

def _build_state_from_choices(
    ws,
    choices: dict[str, dict],
    cells: list[tuple[int, int]],
    direction: str,
    sheet_name: str,
) -> WizardState:
    """Build a WizardState from the choices dict, in cell-scan order.

    This is the single aggregation point — keeping it here (not in action
    methods) means undo/reclassify only need to update _choices, and the
    pattern is always consistent on save.

    Web-service note: this is the function a REST handler would call after
    the client submits the final classification list.
    """
    state = WizardState(direction=direction, sheet_name=sheet_name)
    seen_t_anchors: set[str] = set()
    for r, c in cells:
        ref  = _cell_ref(r, c)
        meta = choices.get(ref)
        if not meta:
            continue
        choice = meta.get('choice', '')
        value  = ws.cell(row=r, column=c).value
        name   = meta.get('name', '')
        if choice == 'L':
            state.lbl_defs.append((name, 'string', str(value) if value is not None else ''))
            state.body_rows.append(['cell:1', name])
        elif choice == 'C':
            state.lbl_defs.append((name, 'string', str(value) if value is not None else ''))
            state.body_rows.append(['cell:1', name])
        elif choice == 'V':
            state.var_defs.append((name, meta.get('ftype', 'string'), meta.get('match', '.*')))
            state.body_rows.append(['cell:1', name])
        elif choice == 'T':
            # Each table is emitted once from its anchor cell; T-HEAD cells are skipped.
            if ref in seen_t_anchors:
                continue
            seen_t_anchors.add(ref)
            mult      = meta.get('mult', '*')
            cols      = meta.get('columns', [])
            lbl_names = []
            var_names = []
            for col in cols:
                lbl_name   = col.get('lbl_name', 'IGNORE')
                var_name   = col.get('var_name', 'IGNORE')
                var_type   = col.get('var_type', 'string')
                var_match  = col.get('var_match', '.*')
                cell_val   = col.get('cell_value', '')
                state.lbl_defs.append((lbl_name, 'string', cell_val))
                state.var_defs.append((var_name, var_type, var_match))
                lbl_names.append(lbl_name)
                var_names.append(var_name)
            state.body_rows.append([f'table:{mult}'])
            state.body_rows.append(['', 'HEADER:1'] + lbl_names)
            state.body_rows.append(['', 'DATA:*'] + var_names)
        elif choice == 'T-HEAD':
            pass  # handled by the anchor cell above
        elif choice == 'I':
            state.body_rows.append(['cell:1', 'IGNORE'])
    return state


def _choices_to_csv(ws, choices, cells, direction, sheet_name) -> str:
    state = _build_state_from_choices(ws, choices, cells, direction, sheet_name)
    buf = io.StringIO()
    w   = csv.writer(buf)
    w.writerow(['config:', 'read.direction', state.direction])
    for name, typ, text in state.lbl_defs:
        w.writerow(['lbl:', name, typ, text])
    for name, typ, match in state.var_defs:
        w.writerow(['var:', name, typ, match])
    w.writerow(['START:'])
    for row in state.body_rows:
        w.writerow(row)
    w.writerow(['END:'])
    return buf.getvalue()


if _TEXTUAL_OK:

    # ── Modals ─────────────────────────────────────────────────────────────────

    class _FieldsModal(ModalScreen):
        """Multi-field text input.
        ENTER moves between fields; last ENTER submits → dismiss(list[str]).
        ESC → dismiss(None).
        """
        DEFAULT_CSS = """
        _FieldsModal              { align: center middle; }
        _FieldsModal > #dialog    { background: $surface; border: thick $primary;
                                    width: 64; height: auto; padding: 1 2; }
        _FieldsModal Label.title  { text-style: bold; margin-bottom: 1; }
        _FieldsModal Label.lbl    { color: $text-muted; margin-top: 1; }
        _FieldsModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        def __init__(self, title: str, fields: list[tuple[str, str]]) -> None:
            super().__init__()
            self._title  = title
            self._fields = fields

        def compose(self) -> ComposeResult:
            with Vertical(id='dialog'):
                yield Label(self._title, classes='title')
                for i, (lbl, default) in enumerate(self._fields):
                    yield Label(lbl, classes='lbl')
                    yield Input(value=default, id=f'f{i}')
                yield Label('ENTER = confirm  •  ESC = cancel', classes='hint')

        def on_mount(self) -> None:
            try:
                self.query_one('#f0', Input).focus()
            except Exception:
                pass

        def on_input_submitted(self, event: Input.Submitted) -> None:
            inputs = list(self.query(Input))
            try:
                idx = inputs.index(event.input)
            except ValueError:
                idx = len(inputs) - 1
            if idx < len(inputs) - 1:
                inputs[idx + 1].focus()
            else:
                self._submit()

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(None)

        def _submit(self) -> None:
            inputs = list(self.query(Input))
            self.dismiss([
                inp.value if inp.value else default
                for inp, (_, default) in zip(inputs, self._fields)
            ])


    class _GotoModal(ModalScreen):
        DEFAULT_CSS = """
        _GotoModal              { align: center middle; }
        _GotoModal > #dialog    { background: $surface; border: thick $primary;
                                  width: 44; height: auto; padding: 1 2; }
        _GotoModal Label.title  { text-style: bold; margin-bottom: 1; }
        _GotoModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        def compose(self) -> ComposeResult:
            with Vertical(id='dialog'):
                yield Label('Go to cell', classes='title')
                yield Input(placeholder='e.g. B5', id='ref')
                yield Label('ENTER = jump  •  ESC = cancel', classes='hint')

        def on_mount(self) -> None:
            self.query_one('#ref', Input).focus()

        def on_input_submitted(self, event: Input.Submitted) -> None:
            self.dismiss(event.value.strip().upper() or None)

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(None)


    class _ConfigModal(ModalScreen):
        DEFAULT_CSS = """
        _ConfigModal              { align: center middle; }
        _ConfigModal > #dialog    { background: $surface; border: thick $primary;
                                    width: 72; height: auto; padding: 1 3; }
        _ConfigModal Label.title  { text-style: bold; margin-bottom: 1; }
        _ConfigModal Label.sect   { text-style: bold; margin-top: 1; }
        _ConfigModal Label.desc   { color: $text-muted; margin-bottom: 1; }
        _ConfigModal Select       { margin-bottom: 1; }
        _ConfigModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        def __init__(self, is_template: bool) -> None:
            super().__init__()
            self._is_template = is_template

        def compose(self) -> ComposeResult:
            note = '  [yellow]⚑ template detected[/yellow]' if self._is_template else ''
            with Vertical(id='dialog'):
                yield Label(f'Wizard configuration{note}', classes='title')
                yield Label('Scan direction', classes='sect')
                yield Label('In which direction does the data read?', classes='desc')
                yield Select(
                    options=[('Left → Right  (LR)', 'LR'), ('Top → Down  (TD)', 'TD')],
                    value='LR', id='dir',
                )
                yield Label('Template mode', classes='sect')
                yield Label('Empty cells after labels are treated as variable slots',
                            classes='desc')
                yield Select(
                    options=[('Yes — empty slots are variables', 'yes'),
                              ('No — classify each cell manually', 'no')],
                    value='yes' if self._is_template else 'no',
                    id='tpl',
                )
                yield Label('ENTER = start  •  ESC = cancel', classes='hint')

        def on_key(self, event) -> None:
            if event.key == 'enter':
                try:
                    direction = str(self.query_one('#dir', Select).value)
                    template  = str(self.query_one('#tpl', Select).value) == 'yes'
                except Exception:
                    direction, template = 'LR', self._is_template
                self.dismiss({'direction': direction, 'template': template})
            elif event.key == 'escape':
                self.dismiss(None)


    class _ZoomModal(ModalScreen):
        """Read-only full-content view of a cell (the table truncates to 20 chars)."""
        DEFAULT_CSS = """
        _ZoomModal              { align: center middle; }
        _ZoomModal > #dialog    { background: $surface; border: thick $primary;
                                  width: 72; height: auto; padding: 1 2; }
        _ZoomModal Label.title  { text-style: bold; color: $accent; margin-bottom: 1; }
        _ZoomModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        def __init__(self, ref: str, value: Any, meta: dict | None) -> None:
            super().__init__()
            self._ref   = ref
            self._value = value
            self._meta  = meta

        def compose(self) -> ComposeResult:
            display = '' if self._value is None else str(self._value)
            status  = ''
            if self._meta:
                ch    = self._meta.get('choice', '')
                cname = _CHOICE_NAME.get(ch, ch)
                ccol  = _CHOICE_COLOR.get(ch, 'white')
                name  = self._meta.get('name', '')
                status = f'\n[bold {ccol}]✓ {cname}[/bold {ccol}]' + (f'  [dim]{name}[/dim]' if name else '')
            with Vertical(id='dialog'):
                yield Label(f'Cell {self._ref}', classes='title')
                yield Static(f'[white]{display}[/white]{status}')
                yield Label('ESC / ENTER to close', classes='hint')

        def on_key(self, event) -> None:
            if event.key in ('escape', 'enter', 'space'):
                self.dismiss(None)


    class _PreviewModal(ModalScreen):
        """Show the current pattern as it would be written to CSV."""
        DEFAULT_CSS = """
        _PreviewModal              { align: center middle; }
        _PreviewModal > #dialog    { background: $surface; border: thick $primary;
                                     width: 80; height: 28; padding: 1 2;
                                     overflow-y: auto; }
        _PreviewModal Label.title  { text-style: bold; color: $accent; margin-bottom: 1; }
        _PreviewModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        def __init__(self, csv_text: str) -> None:
            super().__init__()
            self._csv = csv_text

        def compose(self) -> ComposeResult:
            with Vertical(id='dialog'):
                yield Label('Pattern preview (current state)', classes='title')
                yield Static(self._csv)
                yield Label('ESC to close', classes='hint')

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(None)


    class _HelpModal(ModalScreen):
        DEFAULT_CSS = """
        _HelpModal              { align: center middle; }
        _HelpModal > #dialog    { background: $surface; border: thick $primary;
                                  width: 72; height: auto; max-height: 40;
                                  padding: 1 2; overflow-y: auto; }
        _HelpModal Label.title  { text-style: bold; color: $accent; margin-bottom: 1; }
        _HelpModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        _HELP = """\
[bold cyan]CLASSIFY[/bold cyan]

  [bold green]L[/bold green]  [bold]Label[/bold]  — A text cell that identifies a nearby value.
          e.g. "Invoice No:" → the next cell is the invoice number.
          Pressing L also suggests the next empty/adjacent cell as a Value.

  [bold blue]C[/bold blue]  [bold]Header[/bold] — Section title or navigation marker; no value follows.
          e.g. "APPROVED BY DEPT" used as a structural heading.

  [bold yellow]V[/bold yellow]  [bold]Value[/bold]  — Extract this cell's content (a variable).
          e.g. the actual invoice number, a date, a name.

  [bold magenta]T[/bold magenta]  [bold]Table[/bold]  — Mark the start of a repeating data-row block.
          Next, mark each column in this row with V to define columns.

  [dim]I[/dim]  [bold]Ignore[/bold] — Skip this cell; emit no pattern entry for it.

  [bold]R[/bold]  [bold]Remove[/bold] — Clear the classification of the current cell so you
          can reclassify it. (Also undoable with Ctrl+Z.)

[bold cyan]NAVIGATE[/bold cyan]

  [bold]N[/bold]  Next non-empty cell in scan order
  [bold]P[/bold]  Prev non-empty cell in scan order
  [bold]U[/bold]  Jump to next [italic]unclassified[/italic] non-empty cell (wraps around)
  [bold]G[/bold]  Go to a specific cell reference (e.g. D11)
  [bold]H[/bold]  Highlight all pending (unclassified) cells in amber
       Navigate or press H again to clear the highlight

[bold cyan]OTHER[/bold cyan]

  [bold]ENTER[/bold]     Auto-accept the proposed classification
  [bold]Space[/bold]     Zoom — view the full untruncated cell content
  [bold]F3[/bold]        Preview the current pattern CSV
  [bold]F1[/bold]        This help screen
  [bold]Ctrl+Z[/bold]    Undo last classification
  [bold]Ctrl+D[/bold]    Toggle dark / light mode
  [bold]^P[/bold]        Command palette (search all actions by name)
  [bold]E[/bold]         End wizard and save the pattern
  [bold]Ctrl+Q[/bold]    Cancel without saving

[bold cyan]LEGEND[/bold cyan]

  [bold green]●[/bold green] green   = Label (L)    [bold yellow]●[/bold yellow] yellow  = Value (V)
  [bold blue]●[/bold blue] blue    = Header (C)  [bold magenta]●[/bold magenta] magenta = Table (T)
  [dim]○[/dim] dim     = Ignore (I)   white   = not yet classified\
"""

        def compose(self) -> ComposeResult:
            with Vertical(id='dialog'):
                yield Label('grepxcel wizard — Help  (F1)', classes='title')
                yield Static(self._HELP)
                yield Label('ESC to close', classes='hint')

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(None)


    class _LogModal(ModalScreen):
        """Session event log — hidden support tool, triggered by F12."""
        DEFAULT_CSS = """
        _LogModal              { align: center middle; }
        _LogModal > #dialog    { background: $surface; border: thick $warning;
                                 width: 90; height: 34; padding: 1 2;
                                 overflow-y: auto; }
        _LogModal Label.title  { text-style: bold; color: $warning; margin-bottom: 1; }
        _LogModal Label.hint   { color: $text-muted; margin-top: 1; }
        """

        def __init__(self, lines: list[str], data_file: str) -> None:
            super().__init__()
            self._lines     = lines
            self._data_file = data_file
            self._wrote: str | None = None

        def compose(self) -> ComposeResult:
            fname   = os.path.basename(self._data_file)
            content = '\n'.join(self._lines) if self._lines else '[dim](no events yet)[/dim]'
            with Vertical(id='dialog'):
                yield Label(f'Session event log — {fname}', classes='title')
                yield Static(content)
                yield Label('ESC to close  •  W to write log to file', classes='hint')

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(self._wrote)
            elif event.key == 'w':
                ts       = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
                stem     = os.path.splitext(os.path.basename(self._data_file))[0]
                log_path = os.path.join(
                    os.path.dirname(os.path.abspath(self._data_file)),
                    f'grepxcel-wizard-log-{stem}-{ts}.txt',
                )
                with open(log_path, 'w', encoding='utf-8') as fh:
                    fh.write('\n'.join(self._lines))
                self._wrote = log_path
                # Update hint to confirm
                try:
                    self.query_one(Label.hint if False else 'Label.hint', Label).update(
                        f'[bold green]Written:[/bold green] {log_path}  •  ESC to close'
                    )
                except Exception:
                    pass


    # ── Main application ───────────────────────────────────────────────────────

    class _TableSetupModal(ModalScreen):
        """Table definition — step 1: name, header-row range, multiplicity.

        Dismissed with {'name', 'range_str', 'mult'} or None on cancel.
        The range_str is validated here (must parse to a single row).
        """
        DEFAULT_CSS = """
        _TableSetupModal              { align: center middle; }
        _TableSetupModal > #dialog    { background: $surface; border: thick $primary;
                                        width: 76; height: auto; padding: 1 3; }
        _TableSetupModal Label.title  { text-style: bold; color: $accent; margin-bottom: 1; }
        _TableSetupModal Label.sect   { text-style: bold; margin-top: 1; }
        _TableSetupModal Label.desc   { color: $text-muted; margin-bottom: 1; }
        _TableSetupModal Label.err    { color: $error; margin-top: 1; }
        _TableSetupModal Label.hint   { color: $text-muted; margin-top: 1; }
        _TableSetupModal Input        { margin-bottom: 1; }
        """

        def __init__(self, default_name: str, default_range: str) -> None:
            super().__init__()
            self._default_name  = default_name
            self._default_range = default_range

        def compose(self) -> ComposeResult:
            with Vertical(id='dialog'):
                yield Label('[bold magenta]Table[/bold magenta] — repeating row block',
                            classes='title')
                yield Label('Table name', classes='sect')
                yield Label('Short identifier used in the output JSON key', classes='desc')
                yield Input(value=self._default_name, id='name')
                yield Label('Header row range', classes='sect')
                yield Label('Cells containing the column labels  (e.g. B11:F11)',
                            classes='desc')
                yield Input(value=self._default_range, id='range',
                            placeholder='e.g. B11:F11')
                yield Label('Multiplicity', classes='sect')
                yield Label(
                    '*=any instances  ·  1=exactly one  ·  {n,m}=bounded range',
                    classes='desc')
                yield Input(value='*', id='mult', placeholder='*')
                yield Label(id='errmsg', classes='err')
                yield Label('ENTER = next field  •  ESC = cancel', classes='hint')

        def on_mount(self) -> None:
            self.query_one('#name', Input).focus()

        def on_input_submitted(self, event: Input.Submitted) -> None:
            inputs = list(self.query(Input))
            try:
                idx = inputs.index(event.input)
            except ValueError:
                idx = len(inputs) - 1
            if idx < len(inputs) - 1:
                inputs[idx + 1].focus()
            else:
                self._validate_and_submit()

        def _validate_and_submit(self) -> None:
            name       = self.query_one('#name',  Input).value.strip()
            range_str  = self.query_one('#range', Input).value.strip().upper()
            mult_raw   = self.query_one('#mult',  Input).value.strip() or '*'
            errlbl     = self.query_one('#errmsg', Label)
            if not name:
                errlbl.update('Table name cannot be empty.')
                self.query_one('#name', Input).focus()
                return
            parts = range_str.split(':')
            if len(parts) != 2:
                errlbl.update('Range must be like B11:F11 (two cell refs separated by :)')
                self.query_one('#range', Input).focus()
                return
            start = _parse_cell_ref(parts[0].strip())
            end   = _parse_cell_ref(parts[1].strip())
            if not start or not end:
                errlbl.update('Could not parse the range — use the format B11:F11.')
                self.query_one('#range', Input).focus()
                return
            if start[0] != end[0]:
                errlbl.update(
                    'The header range must be a single row  '
                    '(same row number on both sides).'
                )
                self.query_one('#range', Input).focus()
                return
            if start[1] > end[1]:
                errlbl.update('Start column must be ≤ end column.')
                self.query_one('#range', Input).focus()
                return
            # Normalise multiplicity
            if mult_raw == '1':
                mult = '1'
            elif mult_raw == '*':
                mult = '*'
            elif mult_raw.startswith('{') and mult_raw.endswith('}'):
                mult = mult_raw
            elif '..' in mult_raw:
                a, b = mult_raw.split('..', 1)
                mult = f'{{{a.strip()},{b.strip()}}}'
            elif mult_raw.isdigit():
                mult = mult_raw
            else:
                mult = '*'
            errlbl.update('')
            self.dismiss({'name': name, 'range_str': range_str,
                          'mult': mult, 'start': start, 'end': end})

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(None)


    class _Panel(Static):
        """Scrollable right-hand classification panel."""


    class WizardTUIApp(App):
        """Full-screen interactive pattern wizard."""

        CSS = """
        Screen  { layers: default modal; }
        #main   { height: 1fr; }

        DataTable {
            width: 1fr;
            height: 1fr;
        }

        _Panel {
            width: 46;
            height: 1fr;
            border-left: solid $primary-darken-1;
            padding: 0 1;
            overflow-y: auto;
        }

        Footer { height: 1; }
        """

        BINDINGS = [
            # Classify (shown in footer)
            Binding('l', 'act_L',  'Label',       show=True),
            Binding('c', 'act_C',  'Header',       show=True),
            Binding('v', 'act_V',  'Value',        show=True),
            Binding('t', 'act_T',  'Table',        show=True),
            Binding('i', 'act_I',  'Ignore',       show=True),
            # Navigate (shown in footer)
            Binding('n', 'nav_next',  'Next',      show=True),
            Binding('p', 'nav_prev',  'Prev',      show=True),
            Binding('g', 'nav_goto',  'Goto',      show=True),
            Binding('e', 'end_save',  'End & Save', show=True),
            # Classify extras (hidden, discoverable via ^P palette)
            Binding('r',      'act_R',             'Remove classif.', show=False),
            # Navigate extras (hidden)
            Binding('u',      'nav_unclassified',  'Next unclassified', show=False),
            Binding('h',      'highlight_pending', 'Highlight pending', show=False),
            # Other (hidden)
            Binding('enter',  'accept',             'Auto-accept',     show=False),
            Binding('space',  'zoom',               'Zoom cell',       show=False),
            Binding('f1',     'show_help',          'Help',            show=False),
            Binding('f3',     'preview',            'Pattern preview', show=False),
            Binding('ctrl+z', 'undo',               'Undo',            show=False),
            Binding('ctrl+q', 'cancel',             'Cancel',          show=False),
            Binding('q',      'cancel',             'Quit',            show=False),
            # F12: hidden session log for debugging / support — not shown anywhere
            Binding('f12',    'show_log',           'Session log',     show=False),
        ]

        def __init__(self, ws, state: WizardState, data_file: str) -> None:
            super().__init__()
            self._ws        = ws
            self._state     = state   # used for direction / sheet_name only; pattern built at save
            self._data_file = data_file

            self._max_row = ws.max_row or 1
            self._max_col = ws.max_column or 1

            self._history  = _load_history(data_file)
            # Primary state: all classifications live here
            self._choices: dict[str, dict] = {}
            self._last_label_base: str | None = None

            self._ws_row = 1
            self._ws_col = 1
            self._is_template     = False
            self._cells: list[tuple[int, int]] = []
            self._initialized     = False   # NOTE: must not be named _ready (conflicts with App._ready)
            self._total_nonempty  = 0

            # Undo / highlight state
            self._undo_stack: list[dict] = []
            self._highlighted: set[str]  = set()

            # Session event log (F12 debug tool — tracks all classify/navigate events)
            self._session_start = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self._event_log: list[str] = [
                f'grepxcel wizard session — {os.path.basename(data_file)}',
                f'Started: {self._session_start}',
            ]

        def _log(self, event_type: str, detail: str) -> None:
            ts   = datetime.datetime.now().strftime('%H:%M:%S')
            line = f'{ts}  {event_type:<10}  {detail}'
            self._event_log.append(line)

        def compose(self) -> ComposeResult:
            yield Header(show_clock=False)
            with Horizontal(id='main'):
                yield DataTable(id='sheet', cursor_type='cell', zebra_stripes=True)
                yield _Panel(id='panel')
            yield Footer()

        async def on_mount(self) -> None:
            fname = os.path.basename(self._data_file)
            self.title = f'grepxcel wizard — {fname} ({self._state.sheet_name})'
            is_tpl = _detect_template(self._ws, self._data_file)
            self.push_screen(_ConfigModal(is_tpl), self._on_config_done)

        def _on_config_done(self, cfg: dict | None) -> None:
            if cfg is None:
                self.exit(result=None)
                return
            self._state.direction = cfg['direction']
            self._is_template     = cfg['template']
            self._cells           = _build_cell_order(self._ws, self._state.direction)
            self._total_nonempty  = sum(
                1 for r, c in self._cells
                if self._ws.cell(row=r, column=c).value is not None
            )
            self._log('CONFIG',
                      f'direction={cfg["direction"]}  template={cfg["template"]}'
                      f'  sheet={self._state.sheet_name}'
                      f'  cells={self._total_nonempty} non-empty')
            self._populate_table()
            first = _find_next_nonempty(self._cells, self._ws, 0)
            if first is not None:
                r, c = self._cells[first]
                self._ws_row, self._ws_col = r, c
                self._move_cursor(r, c)
            self._initialized = True
            self._refresh_panel()

        # ── DataTable helpers ──────────────────────────────────────────────────

        def _populate_table(self) -> None:
            table = self.query_one('#sheet', DataTable)
            table.clear(columns=True)
            table.add_column('#', width=4, key='__rn__')
            for col in range(1, self._max_col + 1):
                table.add_column(_col_label(col), key=str(col))
            for row in range(1, self._max_row + 1):
                cells: list[Any] = [RichText(str(row), style='dim')]
                for col in range(1, self._max_col + 1):
                    v    = self._ws.cell(row=row, column=col).value
                    ref  = _cell_ref(row, col)
                    meta = self._choices.get(ref)
                    cells.append(
                        _styled(v, meta['choice']) if meta
                        else _trunc(v)
                    )
                table.add_row(*cells, key=str(row))
            table.fixed_columns = 1

        def _move_cursor(self, ws_row: int, ws_col: int) -> None:
            table = self.query_one('#sheet', DataTable)
            # DataTable col 0 = row-pin; col N = worksheet col N
            table.move_cursor(row=ws_row - 1, column=ws_col, animate=False)

        def _restyle_cell(self, ws_row: int, ws_col: int) -> None:
            table = self.query_one('#sheet', DataTable)
            ref   = _cell_ref(ws_row, ws_col)
            value = self._ws.cell(row=ws_row, column=ws_col).value
            meta  = self._choices.get(ref)
            new   = _styled(value, meta['choice']) if meta else _trunc(value)
            try:
                table.update_cell(str(ws_row), str(ws_col), new, update_width=False)
            except Exception:
                pass

        def on_data_table_cell_highlighted(
            self, event: DataTable.CellHighlighted
        ) -> None:
            """Update panel whenever the DataTable cursor moves to a new cell."""
            if not self._initialized:
                return
            dt_col = event.coordinate.column
            dt_row = event.coordinate.row
            if dt_col == 0:     # row-pin column — ignore
                return
            ws_row = dt_row + 1
            ws_col = dt_col     # DataTable col N == worksheet col N
            if ws_row == self._ws_row and ws_col == self._ws_col:
                return
            self._clear_highlights()
            self._ws_row = ws_row
            self._ws_col = ws_col
            self._refresh_panel()

        # ── Panel (4 zones) ────────────────────────────────────────────────────

        def _proposal(self) -> str:
            v = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            if v is None:
                return 'var:string' if (self._is_template and self._last_label_base) else 'skip'
            return _propose_type(v, ws=self._ws, row=self._ws_row, col=self._ws_col)

        def _refresh_panel(self) -> None:
            ref      = _cell_ref(self._ws_row, self._ws_col)
            raw      = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            proposal = self._proposal()
            meta     = self._choices.get(ref)
            prior    = self._history.get(ref)

            # ── Zone 1 — CURRENT CELL ─────────────────────────────────────────
            if raw is None:
                if self._is_template and self._last_label_base:
                    val_line = '[dim](empty — template slot)[/dim]'
                else:
                    val_line = '[dim](empty)[/dim]'
            else:
                s = str(raw)
                val_line = f'[white]"{s[:36]}{"…" if len(s) > 36 else ""}[/white]"'

            if proposal == 'label':
                prop_line = '[bold green]label[/bold green]'
            elif proposal.startswith('var:'):
                prop_line = f'[bold yellow]{proposal}[/bold yellow]'
            else:
                prop_line = f'[dim]{proposal}[/dim]'

            # Build status line — special handling for table cells
            if meta:
                ch   = meta['choice']
                ccol = _CHOICE_COLOR.get(ch, 'white')
                if ch == 'T':
                    tname  = meta.get('name', '')
                    trng   = meta.get('range', ref)
                    ncols  = len(meta.get('columns', []))
                    mult   = meta.get('mult', '*')
                    status = (f'[bold magenta]✓ Table anchor[/bold magenta]'
                              f'  [dim]{tname}[/dim]')
                    extra  = [
                        f'  Range:  [magenta]{trng}[/magenta]'
                        f'  ({ncols} col{"s" if ncols != 1 else ""}, mult={mult})',
                        '  [dim]R = remove whole table[/dim]',
                    ]
                elif ch == 'T-HEAD':
                    anchor_ref  = meta.get('anchor', '')
                    anchor_meta = self._choices.get(anchor_ref, {})
                    tname  = anchor_meta.get('name', '')
                    trng   = anchor_meta.get('range', anchor_ref)
                    # Find column index for this ref
                    col_idx = next(
                        (i for i, c in enumerate(anchor_meta.get('columns', []))
                         if c.get('ref') == ref),
                        -1
                    )
                    col_name = ''
                    if col_idx >= 0:
                        col_name = anchor_meta['columns'][col_idx].get('var_name', '')
                    status = (f'[magenta]✓ Table column[/magenta]'
                              f'  [dim]{tname}.{col_name}[/dim]')
                    extra  = [
                        f'  Table:  [magenta]{tname}[/magenta]  range {trng}',
                        '  [dim]R = remove whole table[/dim]',
                    ]
                else:
                    cname  = _CHOICE_NAME.get(ch, ch)
                    status = f'[bold {ccol}]✓ {cname}[/bold {ccol}]'
                    if 'name' in meta:
                        status += f'  [dim]{meta["name"]}[/dim]'
                    extra = []
            elif prior:
                pname  = _CHOICE_NAME.get(prior, prior)
                status = f'[magenta]↺ prev: {pname}[/magenta]'
                extra  = []
            else:
                status = '[dim]not classified[/dim]'
                extra  = []

            lines: list[str] = [
                f'[bold cyan]─ {ref} {"─" * (40 - len(ref))}[/bold cyan]',
                f'  {val_line}',
                f'  Proposal:  {prop_line}',
                f'  Status:    {status}',
            ] + extra
            if self._last_label_base and not meta:
                lines += [
                    f'  [dim magenta]← label: {self._last_label_base}_label[/dim magenta]',
                    f'  [dim cyan]  hint: variable → {self._last_label_base}[/dim cyan]',
                ]
            lines.append('')

            # ── Zone 2 — CLASSIFY ─────────────────────────────────────────────
            lines += [
                f'[bold cyan]─ Classify {"─" * 32}[/bold cyan]',
                '  [dim]ENTER[/dim]  auto-accept proposal',
                '  [bold green]L[/bold green]  Label  — text anchors a value',
                '  [bold blue]C[/bold blue]  Header — section title only',
                '  [bold yellow]V[/bold yellow]  Value  — extract this cell',
                '  [bold magenta]T[/bold magenta]  Table  — define range + column names',
                '  [dim]I[/dim]  Ignore — skip',
                '  [bold]R[/bold]  Remove classification',
                '',
            ]

            # ── Zone 3 — NAVIGATE ─────────────────────────────────────────────
            lines += [
                f'[bold cyan]─ Navigate {"─" * 32}[/bold cyan]',
                '  [bold]N[/bold]  Next non-empty',
                '  [bold]P[/bold]  Prev non-empty',
                '  [bold]U[/bold]  Next unclassified',
                '  [bold]G[/bold]  Go to cell (e.g. D11)',
                '  [bold red]E[/bold red]  End & save pattern',
                '',
                '  [dim]Space[/dim] Zoom  [dim]F3[/dim] Preview  [dim]F1[/dim] Help',
                '  [dim]H[/dim] Highlight pending  [dim]^Z[/dim] Undo',
                '  [dim]^D[/dim] Dark/light  [dim]^P[/dim] Palette',
                '',
            ]

            # ── Zone 4 — LEGEND + STATS ───────────────────────────────────────
            counts = {}
            for m in self._choices.values():
                c = m.get('choice', '')
                counts[c] = counts.get(c, 0) + 1

            done    = len(self._choices)
            total   = self._total_nonempty
            pct     = int(done / total * 100) if total else 0

            lines += [
                f'[bold cyan]─ Legend {"─" * 34}[/bold cyan]',
                '  [bold green]●[/bold green] green   = Label (L)',
                '  [bold yellow]●[/bold yellow] yellow  = Value (V)',
                '  [bold blue]●[/bold blue] blue    = Header (C)',
                '  [bold magenta]●[/bold magenta] magenta = Table (T)',
                '  [dim]○[/dim] dim     = Ignore (I)',
                '',
                f'[bold cyan]─ Stats {"─" * 35}[/bold cyan]',
                f'  Sheet:  {self._max_row} rows × {self._max_col} cols',
                f'  Done:   [bold]{done}[/bold] / {total} non-empty  ({pct}%)',
                f'  Labels: {counts.get("L", 0)}   Values: {counts.get("V", 0)}'
                f'   Hdrs: {counts.get("C", 0)}',
                f'  Tables: {counts.get("T", 0)}   Ignored: {counts.get("I", 0)}'
                f'   Dir: {self._state.direction}',
                f'  Template: {"[yellow]ON[/yellow]" if self._is_template else "[dim]off[/dim]"}',
            ]
            if self._undo_stack:
                lines.append(f'  [dim]Undo depth: {len(self._undo_stack)}[/dim]')

            self.query_one('#panel', _Panel).update('\n'.join(lines))

        # ── Table helpers ─────────────────────────────────────────────────────

        def _detect_table_end_col(self, ws_row: int, ws_col: int) -> int:
            """Return the last col in ws_row that has a value, starting from ws_col."""
            end = ws_col
            for c in range(ws_col, self._max_col + 1):
                if self._ws.cell(row=ws_row, column=c).value is not None:
                    end = c
                else:
                    break
            return end

        def _extract_table_columns(self,
                                   start_row: int, start_col: int,
                                   end_col: int) -> list[dict]:
            """Return column spec list from the header row range."""
            cols = []
            for c in range(start_col, end_col + 1):
                val  = self._ws.cell(row=start_row, column=c).value
                ref  = _cell_ref(start_row, c)
                slug = _slugify(str(val)) if val is not None else f'col{c}'
                cols.append({
                    'ref':        ref,
                    'row':        start_row,
                    'col':        c,
                    'cell_value': str(val) if val is not None else '',
                    'lbl_name':   f'col_{slug}_label',
                    'var_name':   slug,
                    'var_type':   'string',
                    'var_match':  '.*',
                })
            return cols

        def _commit_table(self, anchor_ref: str, name: str, mult: str,
                          range_str: str, cols: list[dict]) -> None:
            """Write T + T-HEAD entries to _choices and restyle all cells."""
            anchor_meta = {
                'choice':  'T',
                'name':    name,
                'mult':    mult,
                'range':   range_str,
                'columns': cols,
            }
            self._choices[anchor_ref] = anchor_meta
            parsed = _parse_cell_ref(anchor_ref)
            if parsed:
                self._restyle_cell(parsed[0], parsed[1])
            for col in cols[1:]:
                self._choices[col['ref']] = {'choice': 'T-HEAD', 'anchor': anchor_ref}
                self._restyle_cell(col['row'], col['col'])

        def _remove_table(self, anchor_ref: str, meta: dict) -> None:
            """Remove the anchor + all T-HEAD cells for a table."""
            refs_to_clear = [anchor_ref]
            refs_to_clear += [col['ref'] for col in meta.get('columns', [])[1:]]
            for ref in refs_to_clear:
                self._choices.pop(ref, None)
                p = _parse_cell_ref(ref)
                if p:
                    self._restyle_cell(p[0], p[1])

        # ── Highlights ────────────────────────────────────────────────────────

        def _clear_highlights(self) -> None:
            if not self._highlighted:
                return
            table = self.query_one('#sheet', DataTable)
            for ref in self._highlighted:
                parsed = _parse_cell_ref(ref)
                if not parsed:
                    continue
                r, c = parsed
                val  = self._ws.cell(row=r, column=c).value
                meta = self._choices.get(ref)
                disp = _styled(val, meta['choice']) if meta else ('' if val is None else str(val)[:20])
                try:
                    table.update_cell(str(r), str(c), disp, update_width=False)
                except Exception:
                    pass
            self._highlighted.clear()

        async def action_highlight_pending(self) -> None:
            if self._highlighted:           # toggle off
                self._log('HIGHLIGHT', f'cleared ({len(self._highlighted)} cells)')
                self._clear_highlights()
                return
            table = self.query_one('#sheet', DataTable)
            count = 0
            for r, c in self._cells:
                ref = _cell_ref(r, c)
                if ref in self._choices:    # already classified
                    continue
                val = self._ws.cell(row=r, column=c).value
                if val is None:             # empty cell
                    continue
                disp = RichText(str(val)[:20], style=_STYLE['PENDING'])
                try:
                    table.update_cell(str(r), str(c), disp, update_width=False)
                    self._highlighted.add(ref)
                    count += 1
                except Exception:
                    pass
            if count:
                self._log('HIGHLIGHT', f'{count} unclassified cells highlighted')
                self.notify(
                    f'{count} unclassified cells highlighted.  '
                    'Navigate or press H to clear.',
                    timeout=4,
                )
            else:
                self._log('HIGHLIGHT', 'all cells classified — nothing to highlight')
                self.notify('All non-empty cells are classified!', timeout=2)

        # ── Navigation ────────────────────────────────────────────────────────

        def _scan_idx(self) -> int:
            try:
                return self._cells.index((self._ws_row, self._ws_col))
            except ValueError:
                return -1

        def _advance(self) -> None:
            self._clear_highlights()
            nxt = _find_next_nonempty(self._cells, self._ws, self._scan_idx() + 1)
            if nxt is not None:
                r, c = self._cells[nxt]
                self._ws_row, self._ws_col = r, c
                self._move_cursor(r, c)
            else:
                self.notify('No more non-empty cells.', timeout=2)

        async def action_nav_next(self) -> None:
            self._clear_highlights()
            self._advance()

        async def action_nav_prev(self) -> None:
            self._clear_highlights()
            idx = self._scan_idx()
            prv = _find_prev_nonempty(self._cells, self._ws, idx - 1)
            if prv is not None:
                r, c = self._cells[prv]
                self._ws_row, self._ws_col = r, c
                self._move_cursor(r, c)
            else:
                self.notify('No previous non-empty cell.', timeout=2)

        async def action_nav_unclassified(self) -> None:
            """Jump to the next unclassified non-empty cell, wrapping around."""
            self._clear_highlights()
            start = self._scan_idx() + 1
            # Forward pass
            for i in range(start, len(self._cells)):
                r, c = self._cells[i]
                if self._ws.cell(row=r, column=c).value is not None:
                    if _cell_ref(r, c) not in self._choices:
                        self._ws_row, self._ws_col = r, c
                        self._move_cursor(r, c)
                        return
            # Wrap to beginning
            for i in range(0, start):
                r, c = self._cells[i]
                if self._ws.cell(row=r, column=c).value is not None:
                    if _cell_ref(r, c) not in self._choices:
                        self._ws_row, self._ws_col = r, c
                        self._move_cursor(r, c)
                        self.notify('Wrapped to first unclassified cell.', timeout=2)
                        return
            self.notify('All non-empty cells are classified!', timeout=2)

        async def action_nav_goto(self) -> None:
            def _on_ref(ref: str | None) -> None:
                self._clear_highlights()
                if not ref:
                    return
                parsed = _parse_cell_ref(ref)
                if parsed:
                    wr, wc = parsed
                    if 1 <= wr <= self._max_row and 1 <= wc <= self._max_col:
                        self._ws_row, self._ws_col = wr, wc
                        self._move_cursor(wr, wc)
                        self._log('GOTO', f'→ {ref}')
                    else:
                        self.notify(f'Cell {ref} is out of range.', timeout=2)
                        self._log('GOTO-X', f'{ref}  out of range')
                else:
                    self.notify(f'Invalid reference: {ref}', timeout=2)
                    self._log('GOTO-X', f'"{ref}"  invalid reference')
            self.push_screen(_GotoModal(), _on_ref)

        # ── Classification helpers ─────────────────────────────────────────────

        def _push_undo(self, ref: str) -> None:
            self._undo_stack.append({
                'ref':             ref,
                'prev_choice':     self._choices.get(ref),  # dict | None
                'prev_label_base': self._last_label_base,
            })

        def _commit(self, ref: str, meta: dict) -> None:
            self._choices[ref] = meta
            parsed = _parse_cell_ref(ref)
            if parsed:
                self._restyle_cell(parsed[0], parsed[1])

        # ── Classification actions ─────────────────────────────────────────────

        async def action_accept(self) -> None:
            p = self._proposal()
            if p == 'label':
                await self.action_act_L()
            elif p.startswith('var:') or (p == 'skip' and self._is_template
                                          and self._last_label_base):
                await self.action_act_V()
            else:
                self._advance()

        async def action_act_L(self) -> None:
            value = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            ref   = _cell_ref(self._ws_row, self._ws_col)
            slug  = _slugify(str(value)) if value is not None else 'label'
            default = f'{slug}_label'

            def _done(result: list[str] | None) -> None:
                self._clear_highlights()
                if result is None:
                    self._log('LABEL-X', f'{ref}  cancelled')
                    return
                name = result[0]
                old  = self._choices.get(ref, {}).get('choice')
                self._push_undo(ref)
                base = name[:-6] if name.endswith('_label') else name
                self._last_label_base = base
                self._commit(ref, {'choice': 'L', 'name': name})
                reclassify = f'  (was {old})' if old else ''
                rawval = '' if value is None else f'  "{str(value)[:30]}"'
                self._log('LABEL', f'{ref}{rawval}  →  {name}{reclassify}')
                self._refresh_panel()
                self._advance()

            self.push_screen(
                _FieldsModal(
                    '[bold green]Label[/bold green] — text that identifies a nearby value',
                    [('Label anchor name', default)],
                ),
                _done,
            )

        async def action_act_C(self) -> None:
            value = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            ref   = _cell_ref(self._ws_row, self._ws_col)
            slug  = _slugify(str(value)) if value is not None else 'header'

            def _done(result: list[str] | None) -> None:
                self._clear_highlights()
                if result is None:
                    self._log('HEADER-X', f'{ref}  cancelled')
                    return
                name = result[0]
                old  = self._choices.get(ref, {}).get('choice')
                self._push_undo(ref)
                self._last_label_base = None
                self._commit(ref, {'choice': 'C', 'name': name})
                reclassify = f'  (was {old})' if old else ''
                rawval = '' if value is None else f'  "{str(value)[:30]}"'
                self._log('HEADER', f'{ref}{rawval}  →  {name}{reclassify}')
                self._refresh_panel()
                self._advance()

            self.push_screen(
                _FieldsModal(
                    '[bold blue]Header[/bold blue] — section title, no value follows',
                    [('Header name', slug)],
                ),
                _done,
            )

        async def action_act_V(self) -> None:
            value        = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            ref          = _cell_ref(self._ws_row, self._ws_col)
            slug         = _slugify(str(value)) if value is not None else 'field'
            default_name = self._last_label_base or slug
            p            = self._proposal()
            default_type = p[4:] if p.startswith('var:') else 'string'

            def _done(result: list[str] | None) -> None:
                self._clear_highlights()
                if result is None:
                    self._log('VALUE-X', f'{ref}  cancelled')
                    return
                name, ftype, match = result[0], result[1], result[2]
                old  = self._choices.get(ref, {}).get('choice')
                self._push_undo(ref)
                self._last_label_base = None
                self._commit(ref, {'choice': 'V', 'name': name,
                                   'ftype': ftype, 'match': match})
                reclassify = f'  (was {old})' if old else ''
                rawval = '(empty)' if value is None else f'"{str(value)[:30]}"'
                self._log('VALUE',
                          f'{ref}  {rawval}  →  {name}  [{ftype}, {match}]{reclassify}')
                self._refresh_panel()
                self._advance()

            self.push_screen(
                _FieldsModal(
                    '[bold yellow]Value[/bold yellow] — extract this cell\'s content',
                    [('Field name', default_name),
                     ('Type',       default_type),
                     ('Match',      '.*')],
                ),
                _done,
            )

        async def action_act_T(self) -> None:
            """Two-step table definition: (1) name/range/mult → (2) column names."""
            cur_ref  = _cell_ref(self._ws_row, self._ws_col)
            cur_val  = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            slug     = _slugify(str(cur_val)) if cur_val is not None else 'table'
            end_col  = self._detect_table_end_col(self._ws_row, self._ws_col)
            auto_rng = (f'{_cell_ref(self._ws_row, self._ws_col)}'
                        f':{_cell_ref(self._ws_row, end_col)}')

            def _on_setup(setup: dict | None) -> None:
                """Called after Step 1 (setup modal)."""
                self._clear_highlights()
                if setup is None:
                    self._log('TABLE-X', f'{cur_ref}  cancelled at setup')
                    return
                start_row, start_col = setup['start']
                _, end_col2           = setup['end']
                cols  = self._extract_table_columns(start_row, start_col, end_col2)
                name  = setup['name']
                mult  = setup['mult']
                rng   = setup['range_str']

                # Build _FieldsModal fields: label shows cell ref + value, default = slug
                col_fields = []
                for col in cols:
                    val_disp = f'"{_trunc(col["cell_value"], 24)}"' if col['cell_value'] else '(empty)'
                    col_fields.append((f'{col["ref"]}  {val_disp}', col['var_name']))

                def _on_cols(result: list[str] | None) -> None:
                    """Called after Step 2 (column-name modal)."""
                    self._clear_highlights()
                    if result is None:
                        self._log('TABLE-X', f'{cur_ref}  cancelled at column names')
                        return
                    # Merge user-supplied names back into cols
                    for i, col in enumerate(cols):
                        col['var_name']  = result[i] if i < len(result) else col['var_name']
                        col['lbl_name']  = f'col_{col["var_name"]}_label'

                    # Push undo for ALL affected cells as one atomic entry
                    affected = [{'ref': col['ref'],
                                 'prev': self._choices.get(col['ref'])}
                                for col in cols]
                    self._undo_stack.append({
                        'type':            'table',
                        'cells':           affected,
                        'prev_label_base': self._last_label_base,
                    })

                    self._last_label_base = None
                    anchor_ref = cols[0]['ref']
                    self._commit_table(anchor_ref, name, mult, rng, cols)
                    col_names = ', '.join(c['var_name'] for c in cols)
                    self._log('TABLE',
                              f'{rng}  name={name}  mult={mult}  '
                              f'cols=[{col_names}]')
                    self._refresh_panel()
                    self._advance()

                self.push_screen(
                    _FieldsModal(
                        f'[bold magenta]Column variable names[/bold magenta]  '
                        f'table "{name}"  range {rng}',
                        col_fields,
                    ),
                    _on_cols,
                )

            self.push_screen(
                _TableSetupModal(slug, auto_rng),
                _on_setup,
            )

        async def action_act_I(self) -> None:
            ref   = _cell_ref(self._ws_row, self._ws_col)
            old   = self._choices.get(ref, {}).get('choice')
            val   = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            rawval = '(empty)' if val is None else f'"{str(val)[:30]}"'
            self._push_undo(ref)
            self._last_label_base = None
            self._clear_highlights()
            self._commit(ref, {'choice': 'I'})
            reclassify = f'  (was {old})' if old else ''
            self._log('IGNORE', f'{ref}  {rawval}{reclassify}')
            self._refresh_panel()
            self._advance()

        async def action_act_R(self) -> None:
            """Remove classification.  For T/T-HEAD cells removes the whole table."""
            ref  = _cell_ref(self._ws_row, self._ws_col)
            meta = self._choices.get(ref)
            if not meta:
                self.notify('Cell is not classified yet.', timeout=2)
                return
            choice = meta.get('choice', '?')

            if choice in ('T', 'T-HEAD'):
                # Resolve the anchor
                if choice == 'T-HEAD':
                    anchor_ref  = meta.get('anchor', ref)
                    anchor_meta = self._choices.get(anchor_ref, {})
                else:
                    anchor_ref  = ref
                    anchor_meta = meta
                tname = anchor_meta.get('name', '')
                trng  = anchor_meta.get('range', anchor_ref)
                # Snapshot all cells for undo
                affected = [{'ref': anchor_ref,
                             'prev': self._choices.get(anchor_ref)}]
                for col in anchor_meta.get('columns', [])[1:]:
                    cref = col['ref']
                    affected.append({'ref': cref, 'prev': self._choices.get(cref)})
                self._undo_stack.append({
                    'type':            'table',
                    'cells':           affected,
                    'prev_label_base': self._last_label_base,
                })
                self._remove_table(anchor_ref, anchor_meta)
                self._log('REMOVE',
                          f'table "{tname}" ({trng}) — all {len(affected)} cells cleared')
                self._refresh_panel()
                self.notify(f'Table "{tname}" removed ({trng}).', timeout=3)
            else:
                old_name = meta.get('name', '')
                self._push_undo(ref)
                del self._choices[ref]
                self._restyle_cell(self._ws_row, self._ws_col)
                self._log('REMOVE',
                          f'{ref}  cleared  was: {_CHOICE_NAME.get(choice, choice)}'
                          + (f' "{old_name}"' if old_name else ''))
                self._refresh_panel()
                self.notify(f'{ref} cleared — choose a new type.', timeout=2)

        # ── Undo ──────────────────────────────────────────────────────────────

        async def action_undo(self) -> None:
            if not self._undo_stack:
                self.notify('Nothing to undo.', timeout=2)
                return
            entry = self._undo_stack.pop()
            self._last_label_base = entry.get('prev_label_base')

            if entry.get('type') == 'table':
                # Multi-cell table undo
                for cell_entry in entry['cells']:
                    cref = cell_entry['ref']
                    prev = cell_entry['prev']
                    if prev is None:
                        self._choices.pop(cref, None)
                    else:
                        self._choices[cref] = prev
                    p = _parse_cell_ref(cref)
                    if p:
                        self._restyle_cell(p[0], p[1])
                first_ref = entry['cells'][0]['ref'] if entry['cells'] else None
                if first_ref:
                    p = _parse_cell_ref(first_ref)
                    if p:
                        self._ws_row, self._ws_col = p[0], p[1]
                        self._move_cursor(p[0], p[1])
                self._log('UNDO', f'table  {len(entry["cells"])} cells restored')
                self.notify(f'Table undo: {len(entry["cells"])} cells restored', timeout=2)
            else:
                ref     = entry['ref']
                prev    = entry['prev_choice']
                curr_ch = self._choices.get(ref, {}).get('choice', 'unclassified')
                if prev is None:
                    self._choices.pop(ref, None)
                    restored = 'unclassified'
                else:
                    self._choices[ref] = prev
                    restored = _CHOICE_NAME.get(prev.get('choice', ''), prev.get('choice', '?'))
                self._log('UNDO',
                          f'{ref}  {_CHOICE_NAME.get(curr_ch, curr_ch)}  →  {restored}')
                p = _parse_cell_ref(ref)
                if p:
                    self._ws_row, self._ws_col = p[0], p[1]
                    self._move_cursor(p[0], p[1])
                    self._restyle_cell(p[0], p[1])
                self.notify(f'Undone: {ref}', timeout=2)
            self._refresh_panel()

        # ── Other actions ──────────────────────────────────────────────────────

        async def action_zoom(self) -> None:
            ref   = _cell_ref(self._ws_row, self._ws_col)
            value = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            meta  = self._choices.get(ref)
            self.push_screen(_ZoomModal(ref, value, meta), lambda _: None)

        async def action_preview(self) -> None:
            csv_text = _choices_to_csv(
                self._ws, self._choices, self._cells,
                self._state.direction, self._state.sheet_name,
            )
            self.push_screen(_PreviewModal(csv_text), lambda _: None)

        async def action_show_help(self) -> None:
            self.push_screen(_HelpModal(), lambda _: None)

        async def action_show_log(self) -> None:
            """F12: hidden session event log for debugging and support reports."""
            self._log('LOG-OPEN', f'{len(self._event_log)} events so far')
            def _on_close(log_path: str | None) -> None:
                if log_path:
                    self.notify(f'Log written: {log_path}', timeout=6)
            self.push_screen(_LogModal(list(self._event_log), self._data_file), _on_close)

        # ── Exit ──────────────────────────────────────────────────────────────

        async def action_end_save(self) -> None:
            done  = len(self._choices)
            total = self._total_nonempty
            self._log('END-SAVE',
                      f'{done}/{total} classified  →  saving pattern')
            _save_history(self._data_file,
                          {r: m['choice'] for r, m in self._choices.items()})
            state = _build_state_from_choices(
                self._ws, self._choices, self._cells,
                self._state.direction, self._state.sheet_name,
            )
            self.exit(result=state)

        async def action_cancel(self) -> None:
            self._log('CANCEL', 'user cancelled — no pattern saved')
            _save_history(self._data_file,
                          {r: m['choice'] for r, m in self._choices.items()})
            self.exit(result=None)


# ── Public entry point ─────────────────────────────────────────────────────────

def run_wizard_tui(
    data_file: str,
    sheet: str | None = None,
    output: str | None = None,
) -> int:
    """Run the TUI wizard.

    Returns
    -------
    0   pattern saved successfully
    1   user cancelled
    2   textual not installed (caller falls back to sequential wizard)
    """
    if not _TEXTUAL_OK:
        return 2

    try:
        import openpyxl
        wb = openpyxl.load_workbook(data_file, data_only=True)
    except Exception as exc:
        print(f'Error loading {data_file}: {exc}', file=sys.stderr)
        return 1

    if sheet:
        if sheet in wb.sheetnames:
            ws = wb[sheet]
        else:
            try:
                ws = wb.worksheets[int(sheet)]
            except (ValueError, IndexError):
                print(f'Sheet {sheet!r} not found. Available: {wb.sheetnames}',
                      file=sys.stderr)
                return 1
    elif len(wb.sheetnames) > 1:
        print(f'Multiple sheets: {wb.sheetnames}')
        print('Use  --sheet NAME  to choose; defaulting to first sheet.')
        ws = wb.worksheets[0]
    else:
        ws = wb.active

    if not output:
        stem   = os.path.splitext(os.path.basename(data_file))[0]
        output = os.path.join(
            os.path.dirname(os.path.abspath(data_file)),
            f'pattern-{stem}.csv',
        )

    state  = WizardState(sheet_name=ws.title)
    app    = WizardTUIApp(ws, state, data_file)
    result = app.run()

    if result is None:
        print('Wizard cancelled.', file=sys.stderr)
        return 1

    _write_pattern(result, output)
    print(f'\n✓  Pattern written: {output}')
    print(f'   Try:  grepxcel extract -p {output} {data_file}')
    print(f'         grepxcel validate-pattern {output}')
    return 0
