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
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any


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
    display = f'{prompt} [{default}]: ' if default else f'{prompt}: '
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
            # Left-neighbour check: if the cell immediately to the left ends with ':'
            # the current cell is its value, not a column header.
            if ws is not None and row is not None and col is not None and col > 1:
                left = ws.cell(row=row, column=col - 1).value
                if isinstance(left, str) and left.strip().endswith(':'):
                    return 'var:string'
            return 'label'

    if any(any(c.isdigit() for c in w) for w in words):
        return 'var:string'

    # Multi-word, pure alpha → data value (name, description, section title).
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
    print('\n── Phase 1: Configuration ──')
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
    name = _ask('  Label name', default_name)
    state.lbl_defs.append((name, 'string', str(value) if value is not None else ''))
    state.body_rows.append(['cell:1', name])


def _handle_variable(state: WizardState, value: Any, proposal: str) -> None:
    name = _ask('  Field name')
    if not name:
        name = _slugify(str(value)) if value is not None else 'field'
    type_default = _var_type_from_proposal(proposal)
    ftype = _ask('  Type', type_default) or type_default
    match = _ask('  Match pattern', '.*') or '.*'
    state.var_defs.append((name, ftype, match))
    state.body_rows.append(['cell:1', name])


def _run_cell_walk(ws, state: WizardState) -> None:
    print('\n── Phase 2: Cell Walk  (E to finish) ──')
    print('   Empty cells are skipped automatically.\n')
    cells = _build_cell_order(ws, state.direction)
    idx = 0

    while idx < len(cells):
        row, col = cells[idx]
        cell = ws.cell(row=row, column=col)
        value = cell.value
        ref = _cell_ref(row, col)
        proposal = _propose_type(value, ws=ws, row=row, col=col)

        # Silently advance past empty/blank cells — the engine handles them natively.
        # Emitting a cell:1 instruction for an empty cell would consume the next
        # non-empty cell instead, shifting all subsequent fields.
        if proposal == 'skip':
            idx += 1
            continue

        print('─' * 52)
        val_str = repr(value)
        print(f'Cell {ref}  │  {val_str}')
        print(f'Proposed: {proposal}')
        print('─' * 52)
        print('  [L] Label    [V] Variable  [I] Ignore')
        print('  [T] Table    [G] Goto <ref>           [E] End')

        raw = _ask('>').strip()
        answer = raw.upper()

        if answer in ('E', 'END'):
            break

        elif answer == 'L':
            _handle_label(state, value)
            idx += 1

        elif answer == 'V':
            _handle_variable(state, value, proposal)
            idx += 1

        elif answer in ('I', 'S'):
            # Both S and I mean "consume this cell, don't extract it".
            # The engine keyword is IGNORE (SKIP is not a valid engine keyword).
            state.body_rows.append(['cell:1', 'IGNORE'])
            idx += 1

        elif answer == 'T':
            idx = _run_table_subflow(ws, state, row, col, cells, idx)

        elif answer.startswith('G ') or (re.match(r'^[A-Z]+\d+$', answer)):
            target_ref = answer[2:].strip() if answer.startswith('G ') else answer
            parsed = _parse_cell_ref(target_ref)
            if parsed:
                target_row, target_col = parsed
                try:
                    idx = cells.index((target_row, target_col))
                    print(f'  Jumped to {target_ref}.')
                except ValueError:
                    print(f'  Cell {target_ref} is not in the scan order.')
            else:
                print('  Invalid cell reference. Use e.g. A5')

        elif answer == '':
            # Accept proposal
            if proposal == 'label':
                _handle_label(state, value)
            else:
                _handle_variable(state, value, proposal)
            idx += 1

        else:
            print('  Unknown command. L/V/I/T/G/E')


# ── Mini-table sub-flow ────────────────────────────────────────────────────────

def _run_table_subflow(ws, state: WizardState,
                       start_row: int, start_col: int,
                       cells: list[tuple[int, int]], cell_idx: int) -> int:
    """Mini-table sub-flow. Returns the new scan index after the table."""
    print('\n── Mini-table ──')

    # Step 1: Multiplicity
    print('  How many instances?  1 / N / N..M / * (any)')
    mult_raw = _ask('  >', '*')
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
    # For each column the wizard:
    #   - Creates a lbl: anchor from the header cell value → used in HEADER:1
    #   - Asks for a var: field name, type, and regex    → used in DATA:*
    # This split is required: HEADER:1 references lbl names so the engine can
    # re-anchor on repeated header rows; DATA:* references var names to extract values.
    print(f'\n  --- Table columns (row {start_row}) ---')
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
            action = _ask(f'  Cell {ref} (empty)  [I]gnore / [done]', 'done').upper()
            if action in ('DONE', ''):
                break
            header_lbl_names.append('IGNORE')
            header_var_names.append('IGNORE')
            continue

        # Auto-generate lbl anchor from the header cell text (transparent to user)
        col_lbl = f'col_{_slugify(str(value))}'

        print(f'  Cell {ref}  {repr(value)}')
        default_var = _slugify(str(value))
        var_name = _ask('    Data field name', default_var)
        if var_name.upper() in ('DONE', ''):
            break
        if var_name.upper() in ('I', 'IGNORE'):
            header_lbl_names.append('IGNORE')
            header_var_names.append('IGNORE')
            continue

        # Propose type from the first data row
        data_val = ws.cell(row=start_row + 1, column=col).value
        type_proposal = _var_type_from_proposal(
            _propose_type(data_val, ws=ws, row=start_row + 1, col=col)
        )
        var_type = _ask('    Type', type_proposal) or type_proposal
        var_match = _ask('    Match pattern', '.*') or '.*'

        new_lbl_defs.append((col_lbl, 'string', str(value)))
        new_var_defs.append((var_name, var_type, var_match))
        header_lbl_names.append(col_lbl)
        header_var_names.append(var_name)

    if not header_var_names:
        print('  No columns defined — table cancelled.')
        return cell_idx + 1

    # Step 3: FOOTER (optional)
    footer_row_vals: list[str] | None = None
    has_footer = _ask('  FOOTER row? (yes/no)', 'no').lower()
    if has_footer in ('yes', 'y'):
        footer_row_num = start_row + 2
        print(f'\n  --- FOOTER row (approx. row {footer_row_num}) ---')
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
                print(f'  Cell {ref} (empty) → IGNORE')
            else:
                default_name = _slugify(str(value))
                name = _ask(f'  Cell {ref} {repr(value)}  field name', default_name)
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

    # Advance past the rows the wizard explicitly inspected
    last_table_row = start_row + 1
    if has_footer in ('yes', 'y'):
        last_table_row = start_row + 2

    print(f'\n  Table added ({len([n for n in header_var_names if n != "IGNORE"])} columns, '
          f'multiplicity={mult}).')
    print(f'  The engine will capture all matching rows automatically.')

    for i, (r, c) in enumerate(cells):
        if r > last_table_row:
            print(f'  Resuming at cell {_cell_ref(r, c)} — press E to finish.')
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

    print(f'\ngrepxcel wizard — {os.path.basename(data_file)} ({ws.title})')
    print(f'Output will be saved to: {output}\n')

    _run_config_phase(state)
    _run_cell_walk(ws, state)

    # Summary
    n_tables = sum(1 for r in state.body_rows if r and r[0].startswith('table:'))
    mult_strs = [r[0].split(':', 1)[1] for r in state.body_rows
                 if r and r[0].startswith('table:')]
    table_desc = ', '.join(f'multiplicity={m}' for m in mult_strs) if mult_strs else ''

    print('\n─────────────────────────────────────────────────')
    print('Pattern summary')
    print(f'  Direction : {state.direction}')
    print(f'  Currency  : {state.currency_sign}')
    print(f'  Labels    : {len(state.lbl_defs)}')
    print(f'  Variables : {len(state.var_defs)}')
    print(f'  Tables    : {n_tables}' + (f'  ({table_desc})' if table_desc else ''))
    print('─────────────────────────────────────────────────')

    confirmed = _ask('\nSave pattern to', output)
    if confirmed:
        output = confirmed

    _write_pattern(state, output)

    print(f'\nPattern written to: {output}')
    print('\nTry it:')
    print(f'  grepxcel extract -p {output} {data_file}')
    print(f'  grepxcel validate-pattern {output}')
    print('─────────────────────────────────────────────────')

    return 0
