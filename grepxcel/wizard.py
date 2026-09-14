"""Interactive wizard that guides a user through creating a grepxcel pattern CSV.

Three phases:
  1. Config   — sheet, direction, ignore.case, currency symbol; template detection
  2. Cell walk — classify each cell: field label / control label / variable / ignore /
               table / goto / end.  Empty cells are silently skipped unless template
               mode is active and the cell follows a field label.
  3. Summary + save — write the CSV pattern and print a try-it hint

Label naming convention
  Field label  [L]: anchor name is ``<base>_label``; base is auto-suggested as the
               variable name for the immediately next cell.
  Control label [C]: navigation anchor only (section header, separator); no variable
               lookahead; named ``<slug>`` without a suffix.

Table sub-flow (entered with T):
  - Multiplicity (1 / N / N..M / *)
  - HEADER row: each column → auto lbl: anchor + prompted var: field name + type + regex
  - Optional FOOTER row

Template mode
  When the file looks like a blank form (filename keywords or [placeholder] cells),
  the user is asked to confirm template mode.  In template mode an empty cell that
  immediately follows a field label is shown as a variable slot instead of being
  skipped — because the empty slot is exactly where the end-user will fill in data.
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


def _kl(k: str, label: str) -> str:
    """Render a single key+label option."""
    return f'[{_c(k, _C.BOLD, _C.CYAN)}] {label}'


# ── Single-keypress input ─────────────────────────────────────────────────────

def _getch() -> str:
    """Read one keypress without requiring ENTER.

    Returns the raw character.  On Ctrl-C / Ctrl-D / EOF returns '' so the
    caller can exit cleanly.  Falls back to ``input().strip()[:1]`` when stdin
    is not a TTY (CI, piped input).
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
        if ch in ('\x03', '\x04'):
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

_CHOICE_LABELS = {
    'L': 'field label',
    'C': 'control label',
    'V': 'variable',
    'I': 'ignore',
    'T': 'table',
}


def _load_history(data_file: str) -> dict[str, str]:
    key = os.path.abspath(data_file)
    try:
        with open(_HISTORY_PATH, encoding='utf-8') as f:
            return json.load(f).get(key, {})
    except (FileNotFoundError, json.JSONDecodeError, PermissionError, OSError):
        return {}


def _save_history(data_file: str, choices: dict[str, str]) -> None:
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
        pass


# ── Template detection ────────────────────────────────────────────────────────

_TEMPLATE_KEYWORDS = ('template', 'tmpl', 'blank', '-form', '_form', 'formulaire')

_PLACEHOLDER_RE = re.compile(r'^\[.+\]')


def _detect_template(ws, data_file: str) -> bool:
    """Return True if heuristics suggest this is a blank template file."""
    basename = os.path.basename(data_file).lower()
    if any(kw in basename for kw in _TEMPLATE_KEYWORDS):
        return True
    # Any cell containing a [placeholder] value is a strong template signal
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str) and _PLACEHOLDER_RE.match(v.strip()):
                return True
    return False


# ── State ─────────────────────────────────────────────────────────────────────

@dataclass
class WizardState:
    direction: str = 'LR'
    ignore_case_labels: bool = False
    ignore_case_values: bool = False
    currency_sign: str = '€'
    sheet_name: str | None = None
    # Each lbl_def is a 4-tuple: (name, type, match, lbl_mode)
    # lbl_mode: '' = use global lbl.match; 'glob' or 'regexp' = per-field override
    lbl_defs: list[tuple] = field(default_factory=list)
    # Each var_def is a 4-tuple: (name, type, match, col_a_extra)
    # col_a_extra: '' = bare var:; otherwise colon-joined modifier tokens,
    # e.g. 'nullable', 'not-null:trim-whitespace', 'literal', 'glob:nullable'
    var_defs: list[tuple] = field(default_factory=list)
    body_rows: list[list[str]] = field(default_factory=list)
    # ── New global config fields ───────────────────────────────────────────
    template: bool = False                  # config: allow empty cells (blank-form spreadsheets)
    trim_whitespace_labels: bool = False    # config: trim.whitespace.labels yes
    trim_whitespace_values: bool = False    # config: trim.whitespace.values yes
    lbl_match: str = ''               # config: lbl.match ('' = omit / use default)
    var_match: str = ''               # config: var.match ('' = omit / use default)
    empty_aliases: list[str] = field(default_factory=list)  # config: empty.aliases

    def to_dict(self) -> dict:
        """Serialise to a plain JSON-safe dict.  Tuples become lists."""
        import dataclasses
        return dataclasses.asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> 'WizardState':
        """Reconstruct from a plain dict (e.g. loaded from JSON)."""
        return cls(
            direction=d.get('direction', 'LR'),
            ignore_case_labels=d.get('ignore_case_labels', d.get('ignore_case', False)),
            ignore_case_values=d.get('ignore_case_values', d.get('ignore_case', False)),
            currency_sign=d.get('currency_sign', '€'),
            sheet_name=d.get('sheet_name'),
            lbl_defs=[tuple(t) for t in d.get('lbl_defs', [])],
            var_defs=[tuple(t) for t in d.get('var_defs', [])],
            body_rows=[list(r) for r in d.get('body_rows', [])],
            template=d.get('template', False),
            trim_whitespace_labels=d.get('trim_whitespace_labels', d.get('trim_whitespace', False)),
            trim_whitespace_values=d.get('trim_whitespace_values', d.get('trim_whitespace', False)),
            lbl_match=d.get('lbl_match', ''),
            var_match=d.get('var_match', ''),
            empty_aliases=list(d.get('empty_aliases', [])),
        )



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
    """Classify a string cell value as a wizard proposal.

    Rules derived from analysis of 22 real-world fixture files:
    1. Colon suffix          → label
    2. Contains @            → var:string  (email)
    3. Contains ://          → var:string  (URL)
    4. Single word alpha+digit → var:string  (codes: AB123, ELC001)
    5. Single word pure alpha  → label by default;
       override to var:string if left neighbour ends with ':'
    6. Multi-word with digit   → var:string  (Q1 2026)
    7. Multi-word pure alpha   → var:string  (names, descriptions)
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
    """Propose a wizard action for a single cell value."""
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
    if proposal.startswith('var:'):
        return proposal[4:]
    return 'string'


# ── Phase 1: Config ───────────────────────────────────────────────────────────

def _run_config_phase(state: WizardState) -> None:
    print('\n' + _c('── Phase 1: Configuration ──', _C.BOLD, _C.CYAN))
    direction = _ask('Scan direction (LR / TD)', 'LR').upper()
    state.direction = direction if direction in ('LR', 'TD') else 'LR'

    ic = _ask('Ignore case for labels (yes/no)', 'no').lower()
    state.ignore_case_labels = ic in ('yes', 'y')
    ic_v = _ask('Ignore case for values (yes/no)', 'no').lower()
    state.ignore_case_values = ic_v in ('yes', 'y')

    currency = _ask('Currency symbol', '€')
    state.currency_sign = currency if currency else '€'


# ── Phase 2: Cell walk ────────────────────────────────────────────────────────

def _find_next_nonempty(cells: list, ws, from_idx: int) -> int | None:
    """Index of the first non-empty cell at or after *from_idx*, or None."""
    for i in range(from_idx, len(cells)):
        r, c = cells[i]
        v = ws.cell(row=r, column=c).value
        if v is not None and (not isinstance(v, str) or v.strip()):
            return i
    return None


def _find_prev_nonempty(cells: list, ws, from_idx: int) -> int | None:
    """Index of the last non-empty cell at or before *from_idx*, or None."""
    for i in range(from_idx, -1, -1):
        r, c = cells[i]
        v = ws.cell(row=r, column=c).value
        if v is not None and (not isinstance(v, str) or v.strip()):
            return i
    return None


def _build_cell_order(ws, direction: str) -> list[tuple[int, int]]:
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    if direction == 'TD':
        return [(r, c) for c in range(1, max_col + 1) for r in range(1, max_row + 1)]
    return [(r, c) for r in range(1, max_row + 1) for c in range(1, max_col + 1)]


def _handle_field_label(state: WizardState, value: Any) -> str:
    """Ask for a field-label anchor name; return the variable base name for lookahead.

    Default anchor name is ``<base>_label``.  The base (without suffix) is
    returned so the next cell can auto-suggest it as the variable field name.
    """
    slug = _slugify(str(value)) if value is not None else 'label'
    default_name = f'{slug}_label'
    name = _ask(_c('  Label anchor name', _C.CYAN), default_name)
    if not name:
        name = default_name
    ltype = _var_type_from_proposal(_propose_type(value))
    state.lbl_defs.append((name, ltype, str(value) if value is not None else '', ''))
    state.body_rows.append(['cell:1', name])
    # Derive base for lookahead: strip _label suffix if present
    base = name[:-6] if name.endswith('_label') else name
    return base


def _handle_control_label(state: WizardState, value: Any) -> None:
    """Ask for a control-label anchor name (section header / navigation anchor).

    No variable lookahead is set after a control label.
    """
    slug = _slugify(str(value)) if value is not None else 'ctrl'
    name = _ask(_c('  Control label name', _C.CYAN), slug)
    if not name:
        name = slug
    ltype = _var_type_from_proposal(_propose_type(value))
    state.lbl_defs.append((name, ltype, str(value) if value is not None else '', ''))
    state.body_rows.append(['cell:1', name])


def _handle_variable(state: WizardState, value: Any, proposal: str,
                     name_hint: str = '') -> None:
    """Ask for variable name, type, and match pattern.

    *name_hint* pre-fills the field-name prompt when the caller knows a
    likely name (e.g. the base of the preceding field label).
    """
    slug = _slugify(str(value)) if value is not None else 'field'
    default_name = name_hint or slug
    name = _ask(_c('  Field name', _C.CYAN), default_name)
    if not name:
        name = default_name
    type_default = _var_type_from_proposal(proposal)
    ftype = _ask(_c('  Type', _C.CYAN), type_default) or type_default
    match = _ask(_c('  Match pattern', _C.CYAN), '.*') or '.*'
    state.var_defs.append((name, ftype, match, ''))
    state.body_rows.append(['cell:1', name])


def _run_cell_walk(ws, state: WizardState, data_file: str,
                   is_template: bool = False) -> None:
    history = _load_history(data_file)
    choices: dict[str, str] = dict(history)

    if history:
        print(_c(f'\n  ↺  {len(history)} cell choice(s) from a previous session'
                 ' — shown as suggestions.', _C.MAGENTA))

    print('\n' + _c('── Phase 2: Cell Walk ──', _C.BOLD, _C.CYAN))
    lines = ['   Empty cells are skipped.  Press a key — no ENTER needed.']
    if is_template:
        lines.append('   Template mode: empty cells after field labels shown as slots.')
    print(_c('\n'.join(lines), _C.DIM) + '\n')

    cells = _build_cell_order(ws, state.direction)
    idx = 0
    sep = _c('─' * 54, _C.DIM)
    last_label_base: str | None = None   # set after [L]; cleared after [V]/[C]/[I]/[T]
    goto_target = False  # True for exactly one iteration after a successful G jump
    seen: set[str] = set()              # cell refs classified in this session

    while idx < len(cells):
        row, col = cells[idx]
        value = ws.cell(row=row, column=col).value
        ref = _cell_ref(row, col)
        heuristic = _propose_type(value, ws=ws, row=row, col=col)

        # Consume the goto flag before any skip logic so it applies once only
        was_goto_target = goto_target
        goto_target = False

        # Empty cell handling
        if heuristic == 'skip':
            if was_goto_target:
                # User explicitly jumped here via G — always show, even if empty
                heuristic = 'var:string'
            elif is_template and last_label_base is not None:
                # Show as template slot — the end-user will fill it in
                heuristic = 'var:string'
            else:
                # Not in template mode or no pending label → silent skip
                last_label_base = None
                idx += 1
                continue

        prior = history.get(ref)

        # ── Cell display ──────────────────────────────────────────────────────
        print(sep)

        if value is None:
            slot_reason = 'goto target' if was_goto_target else 'template slot'
            val_display = _c(f'(empty — {slot_reason})', _C.DIM)
        else:
            val_display = _c(repr(value), _C.YELLOW)
        print(f'Cell {_c(ref, _C.BOLD, _C.WHITE)}  │  {val_display}')

        # Proposal line
        prop_text = _c(heuristic, _C.BOLD, _proposal_color(heuristic))
        line2 = f'Proposed: {prop_text}'
        if last_label_base is not None:
            line2 += _c(f'  ← follows [{last_label_base}_label]'
                        f'  →  name hint: {last_label_base}', _C.MAGENTA, _C.DIM)
        if prior:
            prior_label = _CHOICE_LABELS.get(prior, prior)
            line2 += _c(f'   ↺ prev [{prior}] {prior_label}', _C.MAGENTA, _C.DIM)
        print(line2)

        if ref in seen:
            print(_c('  🟡  Already classified this session'
                     ' — reclassifying adds a duplicate instruction.', _C.YELLOW))

        print(sep)
        print(f'  {_kl("L", "Field label")}  {_kl("C", "Control")}  '
              f'{_kl("V", "Variable")}  {_kl("I", "Ignore")}')
        print(f'  {_kl("T", "Table")}  {_kl("G", "Goto ref")}  '
              f'{_kl("N", "Next")}  {_kl("P", "Prev")}  {_kl("E", "End")}')
        enter_hint = 'prior' if prior else 'proposal'
        print(_c(f'  ENTER = accept {enter_hint}', _C.DIM))

        print(_c('> ', _C.BOLD, _C.WHITE), end='', flush=True)
        ch = _getch()

        if ch in ('\r', '\n'):
            print()
        elif ch:
            print(_c(ch.upper(), _C.BOLD))
        else:
            print()
            print(_c('\nAborted.', _C.RED), file=sys.stderr)
            _save_history(data_file, choices)
            sys.exit(0)

        answer = prior if ch in ('\r', '\n') and prior else \
                 ('ACCEPT' if ch in ('\r', '\n') else ch.upper())

        # ── Dispatch ──────────────────────────────────────────────────────────

        if answer == 'E':
            choices[ref] = 'E'
            break

        elif answer in ('', 'ACCEPT'):
            if heuristic == 'label':
                base = _handle_field_label(state, value)
                last_label_base = base
                choices[ref] = 'L'
            else:
                _handle_variable(state, value, heuristic,
                                 name_hint=last_label_base or '')
                last_label_base = None
                choices[ref] = 'V'
            seen.add(ref)
            idx += 1

        elif answer == 'L':
            base = _handle_field_label(state, value)
            last_label_base = base
            choices[ref] = 'L'
            seen.add(ref)
            idx += 1

        elif answer == 'C':
            _handle_control_label(state, value)
            last_label_base = None
            choices[ref] = 'C'
            seen.add(ref)
            idx += 1

        elif answer == 'V':
            _handle_variable(state, value, heuristic,
                             name_hint=last_label_base or '')
            last_label_base = None
            choices[ref] = 'V'
            seen.add(ref)
            idx += 1

        elif answer in ('I', 'S'):
            state.body_rows.append(['cell:1', 'IGNORE'])
            last_label_base = None
            choices[ref] = 'I'
            seen.add(ref)
            idx += 1

        elif answer == 'T':
            idx = _run_table_subflow(ws, state, row, col, cells, idx)
            last_label_base = None
            choices[ref] = 'T'
            seen.add(ref)

        elif answer == 'G':
            target_raw = _ask(_c('  Go to cell', _C.CYAN)).strip().upper()
            parsed = _parse_cell_ref(target_raw)
            if parsed:
                target_row, target_col = parsed
                try:
                    idx = cells.index((target_row, target_col))
                    goto_target = True  # force-show target even if empty
                    print(_c(f'  → Jumped to {_cell_ref(target_row, target_col)}.', _C.CYAN))
                except ValueError:
                    print(_c(f'  Cell {target_raw} not in scan order.', _C.RED))
            else:
                print(_c('  Invalid cell reference. Use e.g. B5', _C.RED))

        elif answer == 'N':
            nxt = _find_next_nonempty(cells, ws, idx + 1)
            if nxt is not None:
                idx = nxt
                print(_c(f'  → Next: {_cell_ref(*cells[nxt])}.', _C.CYAN))
            else:
                print(_c('  No more non-empty cells ahead.', _C.DIM))

        elif answer == 'P':
            prv = _find_prev_nonempty(cells, ws, idx - 1)
            if prv is not None:
                idx = prv
                print(_c(f'  → Prev: {_cell_ref(*cells[prv])}.', _C.CYAN))
            else:
                print(_c('  No non-empty cells before this one.', _C.DIM))

        else:
            print(_c('  Unknown — L / C / V / I / T / G / N / P / E', _C.RED))

    _save_history(data_file, choices)


# ── Mini-table sub-flow ────────────────────────────────────────────────────────

def _run_table_subflow(ws, state: WizardState,
                       start_row: int, start_col: int,
                       cells: list[tuple[int, int]], cell_idx: int) -> int:
    """Mini-table sub-flow. Returns the new scan index after the table."""
    print('\n' + _c('── Mini-table ──', _C.BOLD, _C.CYAN))

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

    print(f'\n  {_c(f"--- Table columns (header row {start_row}) ---", _C.DIM)}')
    header_lbl_names: list[str] = []
    header_var_names: list[str] = []
    new_lbl_defs: list[tuple] = []
    new_var_defs: list[tuple] = []
    max_col = ws.max_column or 1

    for col in range(start_col, max_col + 1):
        hcell = ws.cell(row=start_row, column=col)
        value = hcell.value
        ref = _cell_ref(start_row, col)

        if value is None:
            action = _ask(
                f'  Cell {_c(ref, _C.WHITE)} {_c("(empty)", _C.DIM)}'
                '  [I]gnore / [done]', 'done'
            ).upper()
            if action in ('DONE', ''):
                break
            header_lbl_names.append('IGNORE')
            header_var_names.append('IGNORE')
            continue

        col_lbl = f'col_{_slugify(str(value))}_label'

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

        new_lbl_defs.append((col_lbl, 'string', str(value), ''))
        new_var_defs.append((var_name, var_type, var_match, ''))
        header_lbl_names.append(col_lbl)
        header_var_names.append(var_name)

    if not header_var_names:
        print(_c('  No columns defined — table cancelled.', _C.RED))
        return cell_idx + 1

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
                    f'  Cell {_c(ref, _C.WHITE)} {_c(repr(value), _C.YELLOW)}'
                    f'  {_c("field name", _C.CYAN)}',
                    default_name,
                )
                footer_fields.append(name)
                new_var_defs.append((name, 'string', '.*', ''))
        footer_row_vals = footer_fields

    state.lbl_defs.extend(new_lbl_defs)
    state.var_defs.extend(new_var_defs)

    state.body_rows.append([f'table:{mult}'])
    state.body_rows.append(['', 'HEADER:1'] + header_lbl_names)
    state.body_rows.append(['', 'DATA:*'] + header_var_names)
    if footer_row_vals is not None:
        state.body_rows.append(['', 'FOOTER:1'] + footer_row_vals)

    last_table_row = start_row + (2 if has_footer in ('yes', 'y') else 1)

    n_data_cols = len([n for n in header_var_names if n != 'IGNORE'])
    print(_c(f'\n  Table added ({n_data_cols} columns, multiplicity={mult}).', _C.GREEN))
    print(_c('  The engine captures all matching rows automatically.', _C.DIM))

    for i, (r, c) in enumerate(cells):
        if r > last_table_row:
            print(_c(f'  Resuming at cell {_cell_ref(r, c)} — press E to finish.', _C.CYAN))
            return i
    return len(cells)


# ── Phase 3: Write pattern ────────────────────────────────────────────────────

def _pattern_rows(state: WizardState) -> list[list]:
    """Flat list of rows for the pattern — shared by CSV and xlsx writers."""
    rows: list[list] = []
    rows.append(['config:', 'read.direction', state.direction])
    if state.ignore_case_labels:
        rows.append(['config:', 'ignore.case.labels', 'yes'])
    if state.ignore_case_values:
        rows.append(['config:', 'ignore.case.values', 'yes'])
    if state.trim_whitespace_labels:
        rows.append(['config:', 'trim.whitespace.labels', 'yes'])
    if state.trim_whitespace_values:
        rows.append(['config:', 'trim.whitespace.values', 'yes'])
    rows.append(['config:', 'currency.sign', state.currency_sign])
    if state.lbl_match:
        rows.append(['config:', 'lbl.match', state.lbl_match])
    if state.var_match:
        rows.append(['config:', 'var.match', state.var_match])
    for alias in state.empty_aliases:
        rows.append(['config:', 'empty.aliases', alias])
    for t in state.lbl_defs:
        name, ltype, value = t[0], t[1], t[2]
        lbl_mode = t[3] if len(t) > 3 else ''
        # '(default)' is the UI sentinel for "use global default" — never write it
        col_a = f'lbl:{lbl_mode}' if lbl_mode and lbl_mode != '(default)' else 'lbl:'
        rows.append([col_a, name, ltype, value])
    for t in state.var_defs:
        name, vtype, match = t[0], t[1], t[2]
        col_a_extra = t[3] if len(t) > 3 else ''
        col_a = f'var:{col_a_extra}' if col_a_extra else 'var:'
        rows.append([col_a, name, vtype, match])
    rows.append(['START:'])
    for row in state.body_rows:
        rows.append(row)
    rows.append(['END:'])
    return rows


def _write_pattern(state: WizardState, output_path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    rows = _pattern_rows(state)
    if os.path.splitext(output_path)[1].lower() == '.xlsx':
        import openpyxl as _openpyxl  # already a core dependency
        from .pattern_colors import colorize_pattern_file
        wb = _openpyxl.Workbook()
        ws = wb.active
        ws.title = 'pattern'
        for row in rows:
            ws.append([str(c) if c is not None else '' for c in row])
        wb.save(output_path)
        colorize_pattern_file(output_path)
    else:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in rows:
                writer.writerow(row)


# ── Entry point ───────────────────────────────────────────────────────────────

def _save_state_json(state: WizardState, path: str) -> None:
    """Write *state* as JSON to *path*."""
    with open(path, 'w', encoding='utf-8') as fh:
        json.dump(state.to_dict(), fh, indent=2, ensure_ascii=False)
    print(_c(f'   State saved → {path}', _C.DIM))


def run_wizard(
    data_file: str | None,
    sheet: str | None = None,
    output: str | None = None,
    fmt: str = 'xlsx',
    no_tui: bool = False,
    load_state: str | None = None,
    save_state: str | None = None,
    load_pattern: str | None = None,
) -> int:
    """Interactive wizard: loads *data_file*, walks cells, writes a pattern file.

    *fmt* controls the default output format when *output* is not given —
    ``'xlsx'`` (default) or ``'csv'``.  When *output* is specified its
    extension takes precedence.

    When *textual* is installed and stdout is a TTY the full-screen TUI is used
    by default.  Pass ``no_tui=True`` (or set ``GREPXCEL_NO_TUI=1``) to fall
    back to the sequential terminal wizard.

    ``load_state``
        Path to a JSON state file previously written by ``--save-state``.
        Skips the interactive walk and writes the pattern directly.
    ``save_state``
        Path to write the wizard session state as JSON after the pattern is saved.
        Enables replay and scripted testing.
    ``load_pattern``
        Path to an existing pattern file (.xlsx or .csv).  Pre-populates the
        TUI with field classifications from that pattern so the user can review
        and adjust rather than start from scratch.  Ignored in ``--no-tui``
        / sequential mode.
    """
    # ── load-state mode: JSON → pattern, no interactive walk ─────────────────
    if load_state:
        try:
            with open(load_state, encoding='utf-8') as fh:
                state = WizardState.from_dict(json.load(fh))
        except (FileNotFoundError, json.JSONDecodeError, OSError) as exc:
            print(f'Error loading state file {load_state!r}: {exc}', file=sys.stderr)
            return 1
        if not output:
            stem = os.path.splitext(os.path.basename(load_state))[0]
            if data_file:
                stem = os.path.splitext(os.path.basename(data_file))[0]
            ts     = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            output = f'{stem}-wizard-{ts}.{fmt}'
        _write_pattern(state, output)
        print(_c(f'\n🟢  Pattern written to: {output}', _C.BOLD, _C.GREEN))
        if save_state:
            _save_state_json(state, save_state)
        return 0

    # TUI wizard is deprecated in favour of the browser-based web-wizard.
    # It is disabled here; the sequential terminal wizard below is the fallback.
    # NOTE: wizard_tui.py is NOT removed — wizard_api.py imports pure-Python
    #       utilities from it (_infer_cell_type, _preload_from_pattern, etc.).
    _c_dim = '\033[2m'
    _c_rst = '\033[0m'
    if sys.stdout.isatty():
        print(
            f'{_c_dim}ℹ  The TUI wizard is deprecated.  '
            f'For the best experience use:\n'
            f'   grepxcel web-wizard {data_file}\n'
            f'Continuing with the sequential terminal wizard…{_c_rst}',
            file=sys.stderr,
        )
    # _force_seq kept for the env-var / --no-tui path; TUI branch is intentionally skipped
    _force_seq = True  # noqa: F841 (kept for readability / future re-enable)
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
        print(f'Sheets: {wb.sheetnames}')
        chosen = _ask('Sheet', wb.sheetnames[0])
        ws = wb[chosen] if chosen in wb.sheetnames else wb.active
    else:
        ws = wb.active

    state = WizardState(sheet_name=ws.title)

    if not output:
        stem   = os.path.splitext(os.path.basename(data_file))[0]
        ts     = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        output = os.path.join(os.path.dirname(os.path.abspath(data_file)),
                              f'{stem}-wizard-{ts}.{fmt}')

    print(_c('\ngrepxcel wizard', _C.BOLD, _C.CYAN)
          + f' — {_c(os.path.basename(data_file), _C.WHITE)}'
          + f' ({_c(ws.title, _C.DIM)})')
    print(f'Output → {_c(output, _C.DIM)}\n')

    # ── Template detection ────────────────────────────────────────────────────
    is_template = _detect_template(ws, data_file)
    if is_template:
        print(_c('  ⚑  Template detected', _C.YELLOW, _C.BOLD)
              + _c(' — filename or [placeholder] cells found.', _C.DIM))
        confirm = _ask(
            _c('  Activate template mode?', _C.CYAN)
            + _c(' (empty cells after labels shown as variable slots)', _C.DIM)
            + '\n  (yes/no)', 'yes'
        ).lower()
        is_template = confirm in ('yes', 'y', '')
    else:
        tpl = _ask(
            _c('  Template mode?', _C.DIM)
            + ' (yes/no)', 'no'
        ).lower()
        is_template = tpl in ('yes', 'y')

    _run_config_phase(state)
    _run_cell_walk(ws, state, data_file, is_template=is_template)

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

    # B4: warn on duplicate field names (lbl: and var: names share the same namespace)
    all_names = [t[0] for t in state.lbl_defs] + [t[0] for t in state.var_defs]
    seen: set[str] = set()
    dups: list[str] = []
    for n in all_names:
        if n in seen:
            dups.append(n)
        seen.add(n)
    if dups:
        dup_list = ', '.join(sorted(set(dups)))
        print(_c(f'\n🟡  Duplicate field name(s): {dup_list}', _C.YELLOW))
        print(_c('   Fields with the same name will overwrite each other in the pattern.',
                 _C.DIM))
        answer = _ask(_c('   Save anyway?', _C.CYAN), 'no')
        if answer.lower() not in ('yes', 'y'):
            print(_c('   Save cancelled.', _C.DIM))
            return 1

    confirmed = _ask(f'\nSave pattern to', output)
    if confirmed:
        output = confirmed

    _write_pattern(state, output)

    print(_c(f'\n✓  Pattern written to: {output}', _C.BOLD, _C.GREEN))
    if save_state:
        _save_state_json(state, save_state)
    print('\nTry it:')
    print(_c(f'  grepxcel extract -p {output} {data_file}', _C.CYAN))
    print(_c(f'  grepxcel validate-pattern {output}', _C.CYAN))
    print(_c('─' * 49, _C.DIM))

    return 0
