"""Interactive wizard that guides a user through creating a grepxcel pattern CSV.

Three phases:
  1. Config   — sheet, direction, ignore.case, currency symbol
  2. Cell walk — classify each cell as label / variable / ignore / table / goto / end
               Empty/blank cells are silently skipped; the engine handles them natively.
  3. Summary + save — write the CSV pattern and print a try-it hint

Table sub-flow (entered with T):
  - Multiplicity (1 / N / N..M / *)
  - HEADER row: each column → auto lbl: anchor + prompted var: field name + type + regex
  - Optional FOOTER row
"""

from __future__ import annotations

import csv
import datetime
import json
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any


# ── Terminal colours ──────────────────────────────────────────────────────────

_USE_COLOR = sys.stdout.isatty()


class _C:
    RESET   = '\033[0m'
    BOLD    = '\033[1m'
    DIM     = '\033[2m'
    CYAN    = '\033[96m'
    YELLOW  = '\033[93m'
    GREEN   = '\033[92m'
    BLUE    = '\033[94m'
    MAGENTA = '\033[95m'
    RED     = '\033[91m'
    WHITE   = '\033[97m'


def _c(text: str, *codes: str) -> str:
    if not _USE_COLOR:
        return str(text)
    return ''.join(codes) + str(text) + _C.RESET


def _proposal_color(proposal: str) -> str:
    if proposal == 'label':
        return _C.GREEN
    if proposal.startswith('var:'):
        return _C.BLUE
    return _C.DIM


def _key_label(k: str, label: str) -> str:
    return f'[{_c(k, _C.BOLD, _C.CYAN)}] {label}'


# ── Single-keypress input ─────────────────────────────────────────────────────

def _getch() -> str:
    """Read one keypress without requiring ENTER.

    Returns the character (uppercased by the caller).  On Ctrl-C / Ctrl-D /
    EOF returns '' so the caller can exit cleanly.  Falls back to
    ``input().strip()[:1]`` when stdin is not a TTY (e.g. CI, piped input).
    """
    if not sys.stdin.isatty():
        try:
            line = sys.stdin.readline()
            return line.strip()[:1] if line else ''
        except (EOFError, KeyboardInterrupt):
            return ''
    try:
        import tty
        import termios
        fd = sys.stdin.fileno()
        old = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            ch = sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old)
        if ch in ('\x03', '\x04'):  # Ctrl-C, Ctrl-D
            return ''
        return ch
    except (ImportError, AttributeError, OSError):
        try:
            import msvcrt  # type: ignore[import]
            return msvcrt.getwch()
        except ImportError:
            pass
    try:
        return input().strip()[:1]
    except (EOFError, KeyboardInterrupt):
        return ''


# ── Wizard history ────────────────────────────────────────────────────────────

_HISTORY_PATH = os.path.join(os.path.expanduser('~'), '.grepxcel_wizard_history.json')

_CHOICE_LABELS = {'L': 'label', 'V': 'variable', 'I': 'ignore', 'T': 'table'}


def _load_history(data_file: str) -> dict[str, str]:
    """Return {cell_ref: choice_letter} from the most recent wizard run on *data_file*."""
    key = os.path.abspath(data_file)
    try:
        with open(_HISTORY_PATH, encoding='utf-8') as f:
            return json.load(f).get(key, {})
    except (FileNotFoundError, json.JSONDecodeError, PermissionError, OSError):
        return {}


def _save_history(data_file: str, choices: dict[str, str]) -> None:
    """Persist per-cell choices so the next run can suggest them."""
    key = os.path.abspath(data_file)
    try:
        with open(_HISTORY_PATH, encoding='utf-8') as f:
            all_history: dict = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, PermissionError, OSError):
        all_history = {}
    all_history[key] = choices
    try:
        with open(_HISTORY_PATH, 'w', encoding='utf-8') as f:
            json.dump(all_history, f, indent=2, ensure_ascii=False)
    except (PermissionError, OSError):
        pass  # history is a nice-to-have; never crash on it


# ── State ─────────────────────────────────────────────────────────────────────

@dataclass
class WizardState:
    direction: str = 'LR'
    ignore_case: bool = False
    currency_sign: str = '€'
    sheet_name: str | None = None
    lbl_defs: list[tuple[str, str, str]] = field(default_factory=list)
    var_defs: list[tuple[str, str, str]] = field(default_factory=list)
    body_rows: list[list[str]] = field(default_factory=list)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _ask(prompt: str, default: str = '') -> str:
    display = f'{prompt} [{_c(default, _C.DIM)}]: ' if default else f'{prompt}: '
    try:
        answer = input(display).strip()
    except (EOFError, KeyboardInterrupt):
        print('\nAborted.', file=sys.stderr)
        sys.exit(0)
    return answer if answer else default


def _slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r'[^a-z0-9]+', '_', text)
    return text.strip('_') or 'field'


def _col_label(col: int) -> str:
    """1-based column index → Excel letter (1→A, 26→Z, 27→AA)."""
    result = ''
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _parse_cell_ref(ref: str) -> tuple[int, int] | None:
    """Parse 'A5' → (row=5, col=1), or None if invalid."""
    m = re.match(r'^([A-Z]+)(\d+)$', ref.upper().strip())
    if not m:
        return None
    col_str, row_str = m.groups()
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - 64)
    return int(row_str), col


def _cell_ref(row: int, col: int) -> str:
    return f'{_col_label(col)}{row}'


def _classify_string(s: str, ws=None, row: int = None, col: int = None) -> str:
    """Classify a string cell as a wizard proposal using structural signals.

    Rules derived from analysis of 22 real-world fixture files:

    1. Colon suffix  → label  (scalar label convention in all fixtures)
    2. Contains @    → var:string  (email)
    3. Contains ://  → var:string  (URL)
    4. Single word, alpha+digit mixed → var:string  (codes: AB123456, ELC001)
    5. Single word, pure alpha → label by default; but if the left-neighbour cell
       in the same row ends with ':', this cell is a value — override to var:string
       (e.g. "Department:" → "Engineering": left neighbour heuristic fires)
    6. Multi-word with a digit in any word → var:string  (period text: "Q1 2026")
    7. Multi-word, pure alpha → var:string  (proper names, descriptions, titles).
       Multi-word column headers are virtually always handled by the T sub-flow;
       they will not reach this rule in practice.
    """
    if s.endswith(':') or s.endswith('：'):
        return 'label'
    if '@' in s:
        return 'var:string'
    if '://' in s or s.lower().startswith('www.'):
        return 'var:string'

    words = s.split()

    if len(words) == 1:
        has_alpha = any(c.isalpha() for c in s)
        has_digit = any(c.isdigit() for c in s)
        if has_alpha and has_digit:
            return 'var:string'
        if has_alpha:
            if ws is not None and row is not None and col is not None and col > 1:
                left = ws.cell(row=row, column=col - 1).value
                if isinstance(left, str) and left.strip().endswith(':'):
                    return 'var:string'
            return 'label'

    if any(any(c.isdigit() for c in w) for w in words):
        return 'var:string'

    return 'var:string'


def _propose_type(value: Any, ws=None, row: int = None, col: int = None) -> str:
    """Propose a wizard action for a single cell value.

    Pass ``ws``, ``row``, ``col`` from the live worksheet to enable the
    left-neighbour heuristic for single-word pure-alpha strings.
    """
    if value is None:
        return 'skip'
    if isinstance(value, str):
        if not value.strip():
            return 'skip'
        return _classify_string(value.strip(), ws, row, col)
    if isinstance(value, bool):
        return 'var:string'
    if isinstance(value, datetime.datetime):
        return 'var:datetime'
    if isinstance(value, datetime.date):
        return 'var:date'
    if isinstance(value, (datetime.time, datetime.timedelta)):
        return 'var:string'
    if isinstance(value, float):
        return 'var:integer' if value == int(value) else 'var:currency'
    if isinstance(value, int):
        return 'var:integer'
    return 'var:string'


def _var_type_from_proposal(proposal: str) -> str:
    """'var:currency' → 'currency'; 'label'/'skip' → 'string'."""
    if proposal.startswith('var:'):
        return proposal[4:]
    return 'string'


# ── Phase 1: Config ───────────────────────────────────────────────────────────

def _run_config_phase(state: WizardState) -> None:
    print('\n' + _c('── Phase 1: Configuration ──', _C.BOLD, _C.CYAN))
    direction = _ask('Scan direction (LR / TD)', 'LR').upper()
    state.direction = direction if direction in ('LR', 'TD') else 'LR'

    ic = _ask('Ignore case (yes/no)', 'no').lower()
    state.ignore_case = ic in ('yes', 'y')

    currency = _ask('Currency symbol', '€')
    state.currency_sign = currency if currency else '€'


# ── Phase 2: Cell walk ────────────────────────────────────────────────────────

def _build_cell_order(ws, direction: str) -> list[tuple[int, int]]:
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    if direction == 'TD':
        return [(r, c) for c in range(1, max_col + 1) for r in range(1, max_row + 1)]
    return [(r, c) for r in range(1, max_row + 1) for c in range(1, max_col + 1)]


def _handle_label(state: WizardState, value: Any) -> None:
    default_name = _slugify(str(value)) if value is not None else 'label'
    name = _ask(_c('  Label name', _C.CYAN), default_name)
    state.lbl_defs.append((name, 'string', str(value) if value is not None else ''))
    state.body_rows.append(['cell:1', name])


def _handle_variable(state: WizardState, value: Any, proposal: str) -> None:
    name = _ask(_c('  Field name', _C.CYAN))
    if not name:
        name = _slugify(str(value)) if value is not None else 'field'
    type_default = _var_type_from_proposal(proposal)
    ftype = _ask(_c('  Type', _C.CYAN), type_default) or type_default
    match = _ask(_c('  Match pattern', _C.CYAN), '.*') or '.*'
    state.var_defs.append((name, ftype, match))
    state.body_rows.append(['cell:1', name])


def _run_cell_walk(ws, state: WizardState, data_file: str) -> None:
    history = _load_history(data_file)
    choices: dict[str, str] = dict(history)

    if history:
        count = len(history)
        print(_c(f'\n  ↺  {count} cell choice(s) from a previous session — shown as suggestions.',
                 _C.MAGENTA))

    print('\n' + _c('── Phase 2: Cell Walk ──', _C.BOLD, _C.CYAN))
    print(_c('   Empty cells are skipped. Press a key — no ENTER needed.\n', _C.DIM))

    cells = _build_cell_order(ws, state.direction)
    idx = 0
    sep = _c('─' * 52, _C.DIM)

    while idx < len(cells):
        row, col = cells[idx]
        cell = ws.cell(row=row, column=col)
        value = cell.value
        ref = _cell_ref(row, col)
        heuristic = _propose_type(value, ws=ws, row=row, col=col)

        if heuristic == 'skip':
            idx += 1
            continue

        prior = history.get(ref)

        print(sep)
        print(f'Cell {_c(ref, _C.BOLD, _C.WHITE)}  │  {_c(repr(value), _C.YELLOW)}')

        prop_text = _c(heuristic, _C.BOLD, _proposal_color(heuristic))
        if prior:
            prior_label = _CHOICE_LABELS.get(prior, prior)
            prior_hint = _c(f'  ↺ prev [{prior}] {prior_label}', _C.MAGENTA, _C.DIM)
            print(f'Proposed: {prop_text}{prior_hint}')
        else:
            print(f'Proposed: {prop_text}')

        print(sep)
        print(f'  {_key_label("L", "Label")}  {_key_label("V", "Variable")}  '
              f'{_key_label("I", "Ignore")}  {_key_label("T", "Table")}  '
              f'{_key_label("G", "Goto")}  {_key_label("E", "End")}')
        enter_hint = ('prior' if prior else 'proposal')
        print(_c(f'  ENTER = accept {enter_hint}', _C.DIM))

        print(_c('> ', _C.BOLD, _C.WHITE), end='', flush=True)
        ch = _getch()

        # Echo the pressed key and move to new line
        if ch in ('\r', '\n'):
            print()
        elif ch:
            print(_c(ch.upper(), _C.BOLD))
        else:
            # Ctrl-C / EOF
            print()
            print(_c('\nAborted.', _C.RED), file=sys.stderr)
            _save_history(data_file, choices)
            sys.exit(0)

        # Map ENTER to prior choice or heuristic acceptance
        if ch in ('\r', '\n'):
            answer = prior if prior else 'ACCEPT'
        else:
            answer = ch.upper()

        if answer == 'E':
            choices[ref] = 'E'
            break

        elif answer in ('', 'ACCEPT'):
            if heuristic == 'label':
                _handle_label(state, value)
                choices[ref] = 'L'
            else:
                _handle_variable(state, value, heuristic)
                choices[ref] = 'V'
            idx += 1

        elif answer == 'L':
            _handle_label(state, value)
            choices[ref] = 'L'
            idx += 1

        elif answer == 'V':
            _handle_variable(state, value, heuristic)
            choices[ref] = 'V'
            idx += 1

        elif answer in ('I', 'S'):
            state.body_rows.append(['cell:1', 'IGNORE'])
            choices[ref] = 'I'
            idx += 1

        elif answer == 'T':
            idx = _run_table_subflow(ws, state, row, col, cells, idx)
            choices[ref] = 'T'

        elif answer == 'G':
            target_raw = _ask(_c('  Go to cell', _C.CYAN)).strip().upper()
            parsed = _parse_cell_ref(target_raw)
            if parsed:
                target_row, target_col = parsed
                try:
                    idx = cells.index((target_row, target_col))
                    print(_c(f'  → Jumped to {target_raw}.', _C.CYAN))
                except ValueError:
                    print(_c(f'  Cell {target_raw} is not in the scan order.', _C.RED))
            else:
                print(_c('  Invalid cell reference. Use e.g. B5', _C.RED))

        else:
            print(_c('  Unknown — L / V / I / T / G / E', _C.RED))

    _save_history(data_file, choices)


# ── Mini-table sub-flow ────────────────────────────────────────────────────────

def _run_table_subflow(ws, state: WizardState,
                       start_row: int, start_col: int,
                       cells: list[tuple[int, int]], cell_idx: int) -> int:
    """Mini-table sub-flow. Returns the new scan index after the table."""
    print('\n' + _c('── Mini-table ──', _C.BOLD, _C.CYAN))

    # Step 1: Multiplicity
    print('  How many instances?  1 / N / N..M / * (any)')
    mult_raw = _ask(_c('  >', _C.BOLD), '*')
    if mult_raw == '1':
        mult = '1'
    elif mult_raw == '*':
        mult = '*'
    elif '..' in mult_raw:
        parts = mult_raw.split('..', 1)
        mult = f'{{{parts[0].strip()},{parts[1].strip()}}}'
    elif mult_raw.isdigit():
        mult = mult_raw
    else:
        mult = '*'

    # Step 2: Define columns (HEADER + DATA combined)
    print(f'\n  {_c(f"--- Table columns (row {start_row}) ---", _C.DIM)}')
    header_lbl_names: list[str] = []
    header_var_names: list[str] = []
    new_lbl_defs: list[tuple[str, str, str]] = []
    new_var_defs: list[tuple[str, str, str]] = []
    max_col = ws.max_column or 1

    for col in range(start_col, max_col + 1):
        hcell = ws.cell(row=start_row, column=col)
        value = hcell.value
        ref = _cell_ref(start_row, col)

        if value is None:
            action = _ask(
                f'  Cell {_c(ref, _C.WHITE)}  '
                f'{_c("(empty)", _C.DIM)}  [I]gnore / [done]', 'done'
            ).upper()
            if action in ('DONE', ''):
                break
            header_lbl_names.append('IGNORE')
            header_var_names.append('IGNORE')
            continue

        col_lbl = f'col_{_slugify(str(value))}'

        print(f'  Cell {_c(ref, _C.WHITE)}  {_c(repr(value), _C.YELLOW)}')
        default_var = _slugify(str(value))
        var_name = _ask(_c('    Data field name', _C.CYAN), default_var)
        if var_name.upper() in ('DONE', ''):
            break
        if var_name.upper() in ('I', 'IGNORE'):
            header_lbl_names.append('IGNORE')
            header_var_names.append('IGNORE')
            continue

        data_val = ws.cell(row=start_row + 1, column=col).value
        type_proposal = _var_type_from_proposal(
            _propose_type(data_val, ws=ws, row=start_row + 1, col=col)
        )
        var_type = _ask(_c('    Type', _C.CYAN), type_proposal) or type_proposal
        var_match = _ask(_c('    Match pattern', _C.CYAN), '.*') or '.*'

        new_lbl_defs.append((col_lbl, 'string', str(value)))
        new_var_defs.append((var_name, var_type, var_match))
        header_lbl_names.append(col_lbl)
        header_var_names.append(var_name)

    if not header_var_names:
        print(_c('  No columns defined — table cancelled.', _C.RED))
        return cell_idx + 1

    # Step 3: FOOTER (optional)
    footer_row_vals: list[str] | None = None
    has_footer = _ask(_c('  FOOTER row?', _C.CYAN) + ' (yes/no)', 'no').lower()
    if has_footer in ('yes', 'y'):
        footer_row_num = start_row + 2
        print(f'\n  {_c(f"--- FOOTER row (approx. row {footer_row_num}) ---", _C.DIM)}')
        footer_fields: list[str] = []
        for i in range(len(header_var_names)):
            col = start_col + i
            hname = header_var_names[i]
            if hname == 'IGNORE':
                footer_fields.append('IGNORE')
                continue
            fcell = ws.cell(row=footer_row_num, column=col)
            value = fcell.value
            ref = _cell_ref(footer_row_num, col)
            if value is None:
                footer_fields.append('IGNORE')
                print(f'  Cell {_c(ref, _C.WHITE)} {_c("(empty)", _C.DIM)} → IGNORE')
            else:
                default_name = _slugify(str(value))
                name = _ask(
                    f'  Cell {_c(ref, _C.WHITE)} {_c(repr(value), _C.YELLOW)}  '
                    + _c('field name', _C.CYAN),
                    default_name,
                )
                footer_fields.append(name)
                new_var_defs.append((name, 'string', '.*'))
        footer_row_vals = footer_fields

    # Register all new defs
    state.lbl_defs.extend(new_lbl_defs)
    state.var_defs.extend(new_var_defs)

    # Emit body rows
    state.body_rows.append([f'table:{mult}'])
    state.body_rows.append(['', 'HEADER:1'] + header_lbl_names)
    state.body_rows.append(['', 'DATA:*'] + header_var_names)
    if footer_row_vals is not None:
        state.body_rows.append(['', 'FOOTER:1'] + footer_row_vals)

    last_table_row = start_row + 1
    if has_footer in ('yes', 'y'):
        last_table_row = start_row + 2

    n_data_cols = len([n for n in header_var_names if n != 'IGNORE'])
    print(_c(f'\n  Table added ({n_data_cols} columns, multiplicity={mult}).', _C.GREEN))
    print(_c('  The engine will capture all matching rows automatically.', _C.DIM))

    for i, (r, c) in enumerate(cells):
        if r > last_table_row:
            print(_c(f'  Resuming at cell {_cell_ref(r, c)} — press E to finish.', _C.CYAN))
            return i
    return len(cells)


# ── Phase 3: Write pattern ────────────────────────────────────────────────────

def _write_pattern(state: WizardState, output_path: str) -> None:
    """Write the accumulated state as a CSV pattern file."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['config:', 'read.direction', state.direction])
        if state.ignore_case:
            writer.writerow(['config:', 'ignore.case', 'yes'])
        writer.writerow(['config:', 'currency.sign', state.currency_sign])
        for name, ltype, value in state.lbl_defs:
            writer.writerow(['lbl:', name, ltype, value])
        for name, vtype, match in state.var_defs:
            writer.writerow(['var:', name, vtype, match])
        writer.writerow(['START:'])
        for row in state.body_rows:
            writer.writerow(row)
        writer.writerow(['END:'])


# ── Entry point ───────────────────────────────────────────────────────────────

def run_wizard(data_file: str, sheet: str | None = None, output: str | None = None) -> int:
    """Interactive wizard: loads *data_file*, walks cells, writes a CSV pattern."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(data_file, data_only=True)
    except Exception as exc:
        print(f'Error loading {data_file}: {exc}', file=sys.stderr)
        return 1

    # Sheet selection
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
        print(f'Sheets: {wb.sheetnames}')
        chosen = _ask('Sheet', wb.sheetnames[0])
        ws = wb[chosen] if chosen in wb.sheetnames else wb.active
    else:
        ws = wb.active

    state = WizardState(sheet_name=ws.title)

    if not output:
        stem = os.path.splitext(os.path.basename(data_file))[0]
        output = os.path.join(os.path.dirname(os.path.abspath(data_file)),
                              f'pattern-{stem}.csv')

    print(_c(f'\ngrepxcel wizard', _C.BOLD, _C.CYAN)
          + f' — {_c(os.path.basename(data_file), _C.WHITE)} ({_c(ws.title, _C.DIM)})')
    print(f'Output → {_c(output, _C.DIM)}\n')

    _run_config_phase(state)
    _run_cell_walk(ws, state, data_file)

    # Summary
    n_tables = sum(1 for r in state.body_rows if r and r[0].startswith('table:'))
    mult_strs = [r[0].split(':', 1)[1] for r in state.body_rows
                 if r and r[0].startswith('table:')]
    table_desc = ', '.join(f'mult={m}' for m in mult_strs) if mult_strs else ''

    print('\n' + _c('─' * 49, _C.DIM))
    print(_c('Pattern summary', _C.BOLD))
    print(f'  Direction : {_c(state.direction, _C.CYAN)}')
    print(f'  Currency  : {_c(state.currency_sign, _C.CYAN)}')
    print(f'  Labels    : {_c(str(len(state.lbl_defs)), _C.GREEN)}')
    print(f'  Variables : {_c(str(len(state.var_defs)), _C.BLUE)}')
    print(f'  Tables    : {_c(str(n_tables), _C.CYAN)}'
          + (f'  {_c(f"({table_desc})", _C.DIM)}' if table_desc else ''))
    print(_c('─' * 49, _C.DIM))

    confirmed = _ask(f'\nSave pattern to', output)
    if confirmed:
        output = confirmed

    _write_pattern(state, output)

    print(_c(f'\n✓ Pattern written to: {output}', _C.BOLD, _C.GREEN))
    print('\nTry it:')
    print(_c(f'  grepxcel extract -p {output} {data_file}', _C.CYAN))
    print(_c(f'  grepxcel validate-pattern {output}', _C.CYAN))
    print(_c('─' * 49, _C.DIM))

    return 0
