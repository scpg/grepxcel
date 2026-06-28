"""Full-screen TUI wizard for grepxcel.

Requires textual: pip install 'grepxcel[wizard]'

Layout
------
┌─ Header ─────────────────────────────────────────────────────┐
│ grepxcel wizard — file.xlsx (Sheet1)                         │
├─ DataTable (scrollable) ──────────┬─ ClassifyPanel ──────────┤
│   #   A         B       C         │  Cell B4                 │
│   1   INVOICE            DATE     │  'Invoice No:'           │
│ ▶ 2   Invoice No:  [empty]  ...   │  Proposed: label         │
│   3   Due Date:    [empty]  ...   │                          │
│                                   │  [L] Field label         │
│                                   │  [C] Control label       │
│                                   │  [V] Variable            │
│                                   │  [I] Ignore              │
│                                   │                          │
│                                   │  [N] Next  [P] Prev      │
│                                   │  [G] Goto  [E] End       │
│                                   │                          │
│                                   │  ── Pattern ──           │
│                                   │  Labels:    0            │
│                                   │  Variables: 0            │
├─ Footer (key hints) ──────────────┴──────────────────────────┤
"""
from __future__ import annotations

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
    _var_type_from_proposal,
    _CHOICE_LABELS,
)


# ── cell styling ──────────────────────────────────────────────────────────────

_CHOICE_STYLE: dict[str, str] = {
    'L': 'bold green',
    'C': 'bold blue',
    'V': 'bold yellow',
    'I': 'dim',
    'T': 'bold magenta',
}


def _styled(value: Any, choice: str) -> 'RichText':
    display = '' if value is None else str(value)[:20]
    return RichText(display, style=_CHOICE_STYLE.get(choice, ''))


# ── guards ────────────────────────────────────────────────────────────────────

if _TEXTUAL_OK:

    # ── Modals ────────────────────────────────────────────────────────────────

    class _FieldsModal(ModalScreen):
        """Generic input modal with one or more labelled text fields.

        ENTER in any non-last field moves focus to the next; ENTER in the last
        field (or when there is only one) submits. ESC cancels.
        Returns a list of str values (one per field), or None on cancel.
        """

        DEFAULT_CSS = """
        _FieldsModal            { align: center middle; }
        _FieldsModal > #dialog  { background: $surface; border: thick $primary;
                                  width: 64; height: auto; max-height: 22;
                                  padding: 1 2; }
        _FieldsModal .title     { text-style: bold; margin-bottom: 1; }
        _FieldsModal .lbl       { color: $text-muted; margin-top: 1; }
        _FieldsModal .hint      { color: $text-muted; margin-top: 1; }
        """

        def __init__(self, title: str, fields: list[tuple[str, str]]):
            super().__init__()
            self._title  = title
            self._fields = fields   # [(label, default), ...]

        def compose(self) -> ComposeResult:
            with Vertical(id='dialog'):
                yield Label(self._title, classes='title')
                for i, (lbl, default) in enumerate(self._fields):
                    yield Label(lbl, classes='lbl')
                    yield Input(value=default, id=f'f{i}')
                yield Label('ENTER = next / confirm  •  ESC = cancel', classes='hint')

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
                inputs[idx + 1].focus()   # move to next field
            else:
                self._submit()            # last field → commit

        def on_key(self, event) -> None:
            if event.key == 'escape':
                self.dismiss(None)

        def _submit(self) -> None:
            values = [inp.value for inp in self.query(Input)]
            result = [v if v else d for v, (_, d) in zip(values, self._fields)]
            self.dismiss(result)


    class _GotoModal(ModalScreen):
        """Ask for a cell reference to jump to (e.g. 'B5')."""

        DEFAULT_CSS = """
        _GotoModal            { align: center middle; }
        _GotoModal > #dialog  { background: $surface; border: thick $primary;
                                width: 44; height: auto; padding: 1 2; }
        _GotoModal .title     { text-style: bold; margin-bottom: 1; }
        _GotoModal .hint      { color: $text-muted; margin-top: 1; }
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
        """Shown at startup: scan direction + template mode."""

        DEFAULT_CSS = """
        _ConfigModal            { align: center middle; }
        _ConfigModal > #dialog  { background: $surface; border: thick $primary;
                                  width: 60; height: auto; padding: 1 2; }
        _ConfigModal Label      { margin-bottom: 1; }
        _ConfigModal Select     { margin-bottom: 1; }
        """

        def __init__(self, is_template: bool) -> None:
            super().__init__()
            self._is_template = is_template

        def compose(self) -> ComposeResult:
            note = '  [yellow]⚑ template detected[/yellow]' if self._is_template else ''
            with Vertical(id='dialog'):
                yield Label(f'[bold]Wizard configuration[/bold]{note}')
                yield Label('Scan direction')
                yield Select(
                    options=[
                        ('Left → Right  (LR)', 'LR'),
                        ('Top → Down   (TD)', 'TD'),
                    ],
                    value='LR',
                    id='dir',
                )
                yield Label('Template mode  (empty cells after labels shown as slots)')
                yield Select(
                    options=[('Yes', 'yes'), ('No', 'no')],
                    value='yes' if self._is_template else 'no',
                    id='tpl',
                )
                yield Label('[dim]ENTER = start  •  ESC = cancel[/dim]')

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


    # ── Main application ──────────────────────────────────────────────────────

    class _Panel(Static):
        """Scrollable right-side classification panel."""

    class WizardTUIApp(App):
        """Full-screen interactive pattern wizard."""

        CSS = """
        Screen { layers: default modal; }

        #main   { height: 1fr; }

        DataTable {
            width: 1fr;
            height: 1fr;
        }

        _Panel {
            width: 38;
            height: 1fr;
            border-left: solid $primary-darken-1;
            padding: 0 1;
            overflow-y: auto;
        }

        Footer { height: 1; }
        """

        BINDINGS = [
            Binding('l', 'act_L', 'Field label'),
            Binding('c', 'act_C', 'Control'),
            Binding('v', 'act_V', 'Variable'),
            Binding('i', 'act_I', 'Ignore'),
            Binding('n', 'nav_next', 'Next'),
            Binding('p', 'nav_prev', 'Prev'),
            Binding('g', 'nav_goto', 'Goto'),
            Binding('e', 'end_save', 'End & Save'),
            Binding('enter', 'accept', 'Accept', show=False),
            Binding('ctrl+c', 'cancel', 'Cancel', show=False),
        ]

        def __init__(
            self,
            ws,
            state: WizardState,
            data_file: str,
        ) -> None:
            super().__init__()
            self._ws        = ws
            self._state     = state
            self._data_file = data_file

            self._max_row = ws.max_row or 1
            self._max_col = ws.max_column or 1

            self._history = _load_history(data_file)
            self._choices: dict[str, dict] = {}     # ref → {choice, name?, ...}
            self._last_label_base: str | None = None

            # Current worksheet position (1-based); set properly after config
            self._ws_row = 1
            self._ws_col = 1
            self._is_template = False
            self._cells: list[tuple[int, int]] = []
            self._ready = False

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
            cfg    = await self.push_screen_wait(_ConfigModal(is_tpl))

            if cfg is None:
                self.exit(result=None)
                return

            self._state.direction = cfg['direction']
            self._is_template     = cfg['template']
            self._cells           = _build_cell_order(self._ws, self._state.direction)

            self._populate_table()

            # Jump cursor to the first non-empty cell
            first = _find_next_nonempty(self._cells, self._ws, 0)
            if first is not None:
                r, c = self._cells[first]
                self._ws_row, self._ws_col = r, c
                self._move_cursor(r, c)

            self._ready = True
            self._refresh_panel()

        # ── DataTable helpers ─────────────────────────────────────────────────

        def _populate_table(self) -> None:
            table = self.query_one('#sheet', DataTable)
            table.clear(columns=True)

            # First column: row numbers (pinned)
            table.add_column('#', width=4, key='__rn__')
            for col in range(1, self._max_col + 1):
                table.add_column(_col_label(col), key=str(col))

            for row in range(1, self._max_row + 1):
                row_cells: list[Any] = [RichText(str(row), style='dim')]
                for col in range(1, self._max_col + 1):
                    v    = self._ws.cell(row=row, column=col).value
                    ref  = _cell_ref(row, col)
                    meta = self._choices.get(ref)
                    row_cells.append(
                        _styled(v, meta['choice']) if meta
                        else ('' if v is None else str(v)[:20])
                    )
                table.add_row(*row_cells, key=str(row))

            table.fixed_columns = 1  # keep row-number column pinned

        def _move_cursor(self, ws_row: int, ws_col: int) -> None:
            """Move DataTable cursor to the worksheet cell (1-based coords)."""
            table = self.query_one('#sheet', DataTable)
            # DataTable col 0 = row-number pin; DataTable col N = worksheet col N
            table.move_cursor(row=ws_row - 1, column=ws_col, animate=False)

        def _restyle_cell(self, ws_row: int, ws_col: int) -> None:
            """Re-draw one cell in the DataTable after classification."""
            table = self.query_one('#sheet', DataTable)
            ref   = _cell_ref(ws_row, ws_col)
            value = self._ws.cell(row=ws_row, column=ws_col).value
            meta  = self._choices.get(ref)
            new   = _styled(value, meta['choice']) if meta else ('' if value is None else str(value)[:20])
            try:
                table.update_cell(str(ws_row), str(ws_col), new, update_width=False)
            except Exception:
                pass

        def on_data_table_cursor_moved(self, event: DataTable.CursorMoved) -> None:
            if not self._ready:
                return
            dt_col = event.cursor_column
            if dt_col == 0:
                # Cursor landed on the row-number pin — redirect right
                self.call_after_refresh(
                    lambda: self.query_one('#sheet', DataTable)
                    .move_cursor(row=event.cursor_row, column=1, animate=False)
                )
                return
            self._ws_row = event.cursor_row + 1
            self._ws_col = dt_col   # dt_col 1 == worksheet col 1
            self._refresh_panel()

        # ── Right panel ───────────────────────────────────────────────────────

        def _proposal(self) -> str:
            v = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            if v is None:
                if self._is_template and self._last_label_base:
                    return 'var:string'
                return 'skip'
            return _propose_type(v, ws=self._ws, row=self._ws_row, col=self._ws_col)

        def _refresh_panel(self) -> None:
            ref      = _cell_ref(self._ws_row, self._ws_col)
            value    = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            proposal = self._proposal()
            meta     = self._choices.get(ref)
            prior    = self._history.get(ref)

            # Value display
            if value is None:
                if self._is_template and self._last_label_base:
                    vd = '[dim](empty — template slot)[/dim]'
                else:
                    vd = '[dim](empty)[/dim]'
            else:
                vd = f'[yellow]{str(value)[:28]}[/yellow]'

            # Proposal colour
            if proposal == 'label':
                pd = f'[green]{proposal}[/green]'
            elif proposal.startswith('var:'):
                pd = f'[blue]{proposal}[/blue]'
            else:
                pd = f'[dim]{proposal}[/dim]'

            lines = [
                f'[bold cyan]Cell {ref}[/bold cyan]',
                vd,
                '',
                f'Proposed: {pd}',
            ]
            if self._last_label_base:
                lines += [
                    f'[magenta dim]← [{self._last_label_base}_label][/magenta dim]',
                    f'[magenta dim]  hint: {self._last_label_base}[/magenta dim]',
                ]
            if meta:
                cname = _CHOICE_LABELS.get(meta['choice'], meta['choice'])
                lines += ['', f'[green]✓ {cname}[/green]']
                if 'name' in meta:
                    lines.append(f'  [dim]{meta["name"]}[/dim]')
            elif prior:
                prior_name = _CHOICE_LABELS.get(prior, prior)
                lines += ['', f'[magenta]↺ prev [{prior}] {prior_name}[/magenta]']

            lines += [
                '',
                '[bold]─── Keys ───[/bold]',
                ' [green bold]L[/green bold]  Field label',
                ' [blue bold]C[/blue bold]  Control label',
                ' [yellow bold]V[/yellow bold]  Variable',
                ' [dim]I[/dim]  Ignore',
                '',
                ' [cyan]N[/cyan]  Next non-empty',
                ' [cyan]P[/cyan]  Prev non-empty',
                ' [cyan]G[/cyan]  Goto cell ref',
                ' [red]E[/red]  End & save',
                '',
                '[bold]─── Pattern ───[/bold]',
                f'Labels:  {len(self._state.lbl_defs)}',
                f'Vars:    {len(self._state.var_defs)}',
                f'Tables:  '
                f'{sum(1 for r in self._state.body_rows if r and r[0].startswith("table:"))}',
                f'Dir:     {self._state.direction}',
            ]
            if self._is_template:
                lines.append('[yellow]Template: on[/yellow]')

            self.query_one('#panel', _Panel).update('\n'.join(lines))

        # ── Navigation actions ────────────────────────────────────────────────

        def _scan_idx(self) -> int:
            try:
                return self._cells.index((self._ws_row, self._ws_col))
            except ValueError:
                return -1

        async def action_nav_next(self) -> None:
            nxt = _find_next_nonempty(self._cells, self._ws, self._scan_idx() + 1)
            if nxt is not None:
                r, c = self._cells[nxt]
                self._ws_row, self._ws_col = r, c
                self._move_cursor(r, c)
            else:
                self.notify('No more non-empty cells.', timeout=2)

        async def action_nav_prev(self) -> None:
            idx = self._scan_idx()
            prv = _find_prev_nonempty(self._cells, self._ws, idx - 1)
            if prv is not None:
                r, c = self._cells[prv]
                self._ws_row, self._ws_col = r, c
                self._move_cursor(r, c)
            else:
                self.notify('No previous non-empty cell.', timeout=2)

        async def action_nav_goto(self) -> None:
            ref = await self.push_screen_wait(_GotoModal())
            if ref is None:
                return
            parsed = _parse_cell_ref(ref)
            if parsed:
                wr, wc = parsed
                if 1 <= wr <= self._max_row and 1 <= wc <= self._max_col:
                    self._ws_row, self._ws_col = wr, wc
                    self._move_cursor(wr, wc)
                else:
                    self.notify(f'Cell {ref} is out of range.', timeout=2)
            else:
                self.notify(f'Invalid reference: {ref}', timeout=2)

        # ── Classification helpers ────────────────────────────────────────────

        def _commit(self, ref: str, meta: dict) -> None:
            self._choices[ref] = meta
            r, c = _parse_cell_ref(ref)  # type: ignore[misc]
            self._restyle_cell(r, c)

        # ── Classification actions ────────────────────────────────────────────

        async def action_accept(self) -> None:
            p = self._proposal()
            if p == 'label':
                await self.action_act_L()
            elif p.startswith('var:') or (p == 'skip' and self._is_template and self._last_label_base):
                await self.action_act_V()
            else:
                await self.action_nav_next()

        async def action_act_L(self) -> None:
            value = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            ref   = _cell_ref(self._ws_row, self._ws_col)
            slug  = _slugify(str(value)) if value is not None else 'label'
            default = f'{slug}_label'

            result = await self.push_screen_wait(
                _FieldsModal('Field label', [('Label anchor name', default)])
            )
            if result is None:
                return

            name = result[0]
            self._state.lbl_defs.append((name, 'string', str(value) if value is not None else ''))
            self._state.body_rows.append(['cell:1', name])
            base = name[:-6] if name.endswith('_label') else name
            self._last_label_base = base
            self._commit(ref, {'choice': 'L', 'name': name})
            self._refresh_panel()
            await self.action_nav_next()

        async def action_act_C(self) -> None:
            value = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            ref   = _cell_ref(self._ws_row, self._ws_col)
            slug  = _slugify(str(value)) if value is not None else 'ctrl'

            result = await self.push_screen_wait(
                _FieldsModal('Control label', [('Label name', slug)])
            )
            if result is None:
                return

            name = result[0]
            self._state.lbl_defs.append((name, 'string', str(value) if value is not None else ''))
            self._state.body_rows.append(['cell:1', name])
            self._last_label_base = None
            self._commit(ref, {'choice': 'C', 'name': name})
            self._refresh_panel()
            await self.action_nav_next()

        async def action_act_V(self) -> None:
            value   = self._ws.cell(row=self._ws_row, column=self._ws_col).value
            ref     = _cell_ref(self._ws_row, self._ws_col)
            slug    = _slugify(str(value)) if value is not None else 'field'
            default_name = self._last_label_base or slug
            p        = self._proposal()
            default_type = p[4:] if p.startswith('var:') else 'string'

            result = await self.push_screen_wait(
                _FieldsModal('Variable', [
                    ('Field name',    default_name),
                    ('Type',          default_type),
                    ('Match pattern', '.*'),
                ])
            )
            if result is None:
                return

            name, ftype, match = result[0], result[1], result[2]
            self._state.var_defs.append((name, ftype, match))
            self._state.body_rows.append(['cell:1', name])
            self._last_label_base = None
            self._commit(ref, {'choice': 'V', 'name': name})
            self._refresh_panel()
            await self.action_nav_next()

        async def action_act_I(self) -> None:
            ref = _cell_ref(self._ws_row, self._ws_col)
            self._state.body_rows.append(['cell:1', 'IGNORE'])
            self._last_label_base = None
            self._commit(ref, {'choice': 'I'})
            self._refresh_panel()
            await self.action_nav_next()

        # ── End ───────────────────────────────────────────────────────────────

        async def action_end_save(self) -> None:
            _save_history(
                self._data_file,
                {ref: meta['choice'] for ref, meta in self._choices.items()},
            )
            self.exit(result=self._state)

        async def action_cancel(self) -> None:
            _save_history(
                self._data_file,
                {ref: meta['choice'] for ref, meta in self._choices.items()},
            )
            self.exit(result=None)


# ── Public entry point ────────────────────────────────────────────────────────

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
    2   textual not installed (caller should fall back to sequential wizard)
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
        print('Use  --sheet NAME  to choose one; defaulting to first sheet.')
        ws = wb.worksheets[0]
    else:
        ws = wb.active

    if not output:
        stem   = os.path.splitext(os.path.basename(data_file))[0]
        output = os.path.join(
            os.path.dirname(os.path.abspath(data_file)),
            f'pattern-{stem}.csv',
        )

    state = WizardState(sheet_name=ws.title)
    app   = WizardTUIApp(ws, state, data_file)
    result = app.run()

    if result is None:
        print('Wizard cancelled.', file=sys.stderr)
        return 1

    _write_pattern(state, output)
    print(f'\n✓  Pattern written: {output}')
    print(f'   Try:  grepxcel extract -p {output} {data_file}')
    print(f'         grepxcel validate-pattern {output}')
    return 0
