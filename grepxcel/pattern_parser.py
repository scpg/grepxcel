import csv
import os
import re

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from .models import Config, FieldDef, TemplateColumn, TemplateRow, CellInstruction, TableInstruction
from .security import check_regex_safety, SecurityError

_MAX_PATTERN_CELL_LEN = 1_000  # max characters in any pattern file cell value

_A1_RE = re.compile(r'^[A-Z]{1,3}[1-9][0-9]*$', re.IGNORECASE)
_BOUNDED_DATA_RE = re.compile(r'^\{(\d+),(\d+)\}$')
_POSINT_RE = re.compile(r'^[1-9][0-9]*$')

# Recognised template-row keywords inside a table: block.
_VALID_TABLE_ROW_TYPES = frozenset({'HEADER', 'DATA', 'FOOTER', 'SPLITTER', 'SKIP_IF'})


_TRUTHY = frozenset({'1', 'true', 'yes', 'on', 'y'})
_FALSY  = frozenset({'0', 'false', 'no', 'off', 'n', ''})

# Field types the engine knows how to validate (see utils.validate_type).
# Keep this in lockstep with that function — a name here that it can't handle
# would parse cleanly but make every value fail validation.
_VALID_FIELD_TYPES = frozenset({
    'string', 'text',
    'integer',
    'number', 'float', 'decimal',
    'currency', 'percentage',
    'boolean', 'bool',
    'date', 'datetime', 'timestamp',
})


def _truthy(val) -> bool:
    """Interpret a config cell as a boolean.

    Accepts yes/no, true/false, on/off, 1/0 (case-insensitive). Excel may store
    the cell as a real bool, so handle that too. Unknown text → False.
    """
    if isinstance(val, bool):
        return val
    return str(val).strip().lower() in _TRUTHY


class PatternError(Exception):
    """Raised when a pattern file contains a structural or ordering error."""


def _coord_after(prev: tuple, new: tuple, direction: str) -> bool:
    """Return True if `new` comes strictly after `prev` in `direction` reading order."""
    pr, pc = prev
    nr, nc = new
    if direction == 'LR':
        return (nr, nc) > (pr, pc)
    return (nc, nr) > (pc, pr)  # TD: column primary


def _coord_str(pos: tuple) -> str:
    return f'{get_column_letter(pos[1])}{pos[0]}'


class PatternParser:
    def parse(self, filepath: str) -> tuple:
        """
        Parse a pattern file (.xlsx or .csv) into the same structures regardless
        of source format. Returns (global_config, defs, start_sequence).

        Both formats are read into a common 2D-grid IR (list[list[str|None]]);
        all semantics below operate on that grid, so the two source formats share
        one parser. See _read_grid for the format dispatch.
        """
        rows = self._read_grid(filepath)

        global_config = Config()
        defs = {}
        start_sequence = []

        i = 0
        in_start = False
        last_abs_pos: tuple | None = None  # (row, col) of last cell:XY seen

        while i < len(rows):
            row = self._pad(rows[i])
            col_a = row[0]

            if col_a == 'END:':
                break

            if col_a == 'START:':
                self._check_comment_zone(row, 1, i + 1, 'START:')   # B+ may be a # comment
                in_start = True
                i += 1
                continue

            if not in_start:
                if col_a == 'config:':
                    self._apply_global_config(row, global_config)
                    self._check_comment_zone(row, 3, i + 1, 'config:')    # D+ comment
                elif col_a in ('def:', 'var:'):
                    fd = self._parse_field(row, role='var', row_num=i + 1)
                    defs[fd.name] = fd
                    self._check_comment_zone(row, 4, i + 1, col_a)        # E+ comment
                elif col_a == 'lbl:':
                    fd = self._parse_field(row, role='lbl', row_num=i + 1)
                    defs[fd.name] = fd
                    self._check_comment_zone(row, 4, i + 1, 'lbl:')       # E+ comment
                elif col_a in ('doc:', 'info:'):
                    pass  # inline documentation — ignored by engine
                i += 1
                continue

            # Inside START: section
            if col_a and col_a.startswith('cell:'):
                raw = col_a.split(':', 1)[1]
                field = str(row[1] or 'IGNORE')
                self._check_comment_zone(row, 2, i + 1, 'cell:')      # C+ may be a # comment
                if raw in ('1', 'next'):
                    start_sequence.append(CellInstruction(multiplicity=raw, field=field))
                elif _A1_RE.match(raw):
                    ref = raw.upper()
                    new_pos = coordinate_to_tuple(ref)
                    if last_abs_pos is not None and not _coord_after(last_abs_pos, new_pos, global_config.read_direction):
                        raise PatternError(
                            f"cell:{ref} at pattern row {i + 1} is before or equal to the "
                            f"previous absolute reference {_coord_str(last_abs_pos)} in "
                            f"{global_config.read_direction} reading order — unreachable"
                        )
                    last_abs_pos = new_pos
                    start_sequence.append(CellInstruction(multiplicity='abs', field=field, target=ref))
                else:
                    raise PatternError(
                        f"Unknown cell instruction 'cell:{raw}' at pattern row {i + 1}. "
                        f"Use 'cell:next', 'cell:1', or a cell coordinate like 'cell:B5'."
                    )
                i += 1

            elif col_a and col_a.startswith('table:'):
                mult = col_a.split(':', 1)[1]
                # table:<mult> must be '*' or a positive instance count.
                if mult != '*' and not _POSINT_RE.match(mult):
                    raise PatternError(
                        f"Invalid table multiplicity 'table:{mult}' at pattern row "
                        f"{i + 1}. Use 'table:*' (all instances) or a positive "
                        f"count like 'table:1'."
                    )
                table_config = Config(
                    read_direction=global_config.read_direction,
                    currency_sign=global_config.currency_sign,
                    empty_aliases=list(global_config.empty_aliases),
                    ignore_case=global_config.ignore_case,
                )
                template_rows = []
                i += 1

                while i < len(rows):
                    sub = self._pad(rows[i])
                    sub_a = sub[0]
                    sub_b = sub[1]

                    # Non-empty col A means we're back at the top level
                    if sub_a is not None:
                        break

                    # Blank row inside table block
                    if sub_b is None:
                        i += 1
                        continue

                    if sub_b == 'config:':
                        key, val = sub[2], sub[3]
                        if key == 'read.direction' and val:
                            table_config.read_direction = str(val)
                        elif key == 'ignore.case' and val is not None:
                            table_config.ignore_case = _truthy(val)
                        i += 1
                        continue

                    # HEADER / SPLITTER / DATA / FOOTER / SKIP_IF row
                    sub_b_str = str(sub_b) if sub_b else ''
                    is_skip_if = sub_b_str.upper() == 'SKIP_IF'
                    has_colon  = ':' in sub_b_str

                    if is_skip_if or has_colon:
                        if is_skip_if:
                            row_type = 'SKIP_IF'
                            row_mult = ''
                        else:
                            row_type, row_mult = sub_b_str.rsplit(':', 1)

                        # Reject typo'd row keywords (e.g. 'HEDER:1') — otherwise
                        # they would be silently dropped by the engine.
                        if row_type not in _VALID_TABLE_ROW_TYPES:
                            raise PatternError(
                                f"Unknown table row type {sub_b_str!r} at pattern "
                                f"row {i + 1}. Valid row types: "
                                f"{', '.join(sorted(_VALID_TABLE_ROW_TYPES))}."
                            )

                        # Validate the multiplicity for each row type.
                        #   DATA            → '*', positive int, or '{n,m}'
                        #   HEADER/FOOTER/SPLITTER → a positive int (e.g. ':1')
                        if row_type == 'DATA':
                            if not (row_mult == '*' or _POSINT_RE.match(row_mult)
                                    or _BOUNDED_DATA_RE.match(row_mult)):
                                raise PatternError(
                                    f"Invalid DATA multiplicity 'DATA:{row_mult}' at "
                                    f"pattern row {i + 1}. Use 'DATA:*', 'DATA:1', or "
                                    f"a bounded 'DATA:{{n,m}}'."
                                )
                        elif row_type in ('HEADER', 'FOOTER', 'SPLITTER'):
                            if not _POSINT_RE.match(row_mult):
                                raise PatternError(
                                    f"Invalid {row_type} multiplicity "
                                    f"'{row_type}:{row_mult}' at pattern row {i + 1}. "
                                    f"Use a positive count like '{row_type}:1'."
                                )

                        cols_raw = list(sub[2:])
                        while cols_raw and cols_raw[-1] is None:
                            cols_raw.pop()
                        columns = [
                            TemplateColumn(field=str(v) if v is not None else 'EMPTY')
                            for v in cols_raw
                        ]

                        min_rows, max_rows = 0, None
                        if row_type == 'DATA':
                            m = _BOUNDED_DATA_RE.match(row_mult)
                            if m:
                                min_rows = int(m.group(1))
                                max_rows = int(m.group(2))
                                if min_rows > max_rows:
                                    raise PatternError(
                                        f'DATA:{{{min_rows},{max_rows}}}: '
                                        f'min ({min_rows}) must be ≤ max ({max_rows})'
                                    )

                        template_rows.append(TemplateRow(
                            row_type=row_type,
                            multiplicity=row_mult,
                            columns=columns,
                            min_rows=min_rows,
                            max_rows=max_rows,
                        ))

                    i += 1

                # ── DATA type mutual-exclusivity validation ────────────────────
                data_rows = [r for r in template_rows if r.row_type == 'DATA']
                if data_rows:
                    has_bounded = any(r.max_rows is not None for r in data_rows)
                    has_star    = any(r.multiplicity == '*' for r in data_rows)
                    has_one     = any(r.multiplicity == '1' for r in data_rows)
                    type_count  = sum([has_bounded, has_star, has_one])
                    if type_count > 1:
                        raise PatternError(
                            'Table block mixes DATA types. '
                            'Use only one of: DATA:1 (repeatable), DATA:*, or DATA:{n,m}.'
                        )
                    if (has_bounded or has_star) and len(data_rows) > 1:
                        kind = 'DATA:{n,m}' if has_bounded else 'DATA:*'
                        raise PatternError(
                            f'Only one {kind} row is allowed per table block.'
                        )
                    # SKIP_IF without a bounded DATA row is meaningless
                    skip_if_rows = [r for r in template_rows if r.row_type == 'SKIP_IF']
                    if skip_if_rows and not has_bounded:
                        raise PatternError(
                            'SKIP_IF requires DATA:{n,m}. '
                            'SKIP_IF has no effect with DATA:* or DATA:1.'
                        )

                # Structural rules for a table block:
                #   • at least one DATA row is required (HEADER and FOOTER are
                #     optional). A block with no DATA extracts nothing and would
                #     otherwise crash the engine.
                #   • any HEADER row must come before the first DATA row.
                #   • any FOOTER row must come after the last DATA row.
                row_types = [r.row_type for r in template_rows]
                if 'DATA' not in row_types:
                    raise PatternError(
                        f"table:{mult} block has no DATA row — a table must define "
                        f"at least one DATA row (HEADER and FOOTER are optional). "
                        f"Add e.g. a 'DATA:*' row."
                    )
                _data_idxs  = [k for k, t in enumerate(row_types) if t == 'DATA']
                _first_data, _last_data = _data_idxs[0], _data_idxs[-1]
                if any(k > _first_data for k, t in enumerate(row_types) if t == 'HEADER'):
                    raise PatternError(
                        f"table:{mult} block has a HEADER row after a DATA row — "
                        f"HEADER rows must come before DATA."
                    )
                if any(k < _last_data for k, t in enumerate(row_types) if t == 'FOOTER'):
                    raise PatternError(
                        f"table:{mult} block has a FOOTER row before a DATA row — "
                        f"FOOTER rows must come after DATA."
                    )

                start_sequence.append(TableInstruction(
                    multiplicity=mult,
                    config=table_config,
                    rows=template_rows,
                ))

            else:
                i += 1

        return global_config, defs, start_sequence

    # ── Grid readers (front-ends over the common 2D-grid IR) ────────────────────

    def _read_grid(self, filepath: str) -> list:
        """
        Read a pattern file into the common grid IR, dispatching on extension.

          .csv  → _read_csv_and_validate  (plain text, comma-separated)
          .xlsx → openpyxl + _read_and_validate  (default)

        Both return list[list[str|None]] with empty cells as None, so the
        downstream semantic parser is identical for either source format.
        """
        _, ext = os.path.splitext(filepath)
        if ext.lower() == '.csv':
            return self._read_csv_and_validate(filepath)
        wb = openpyxl.load_workbook(filepath)  # data_only=False: we want to see formulas
        ws = wb.active
        return self._read_and_validate(ws)

    def _read_csv_and_validate(self, filepath: str) -> list:
        """
        Read a CSV pattern file into the grid IR with the same content rules as
        the xlsx reader.

        - Empty fields → None, so a blank column A (which marks table-template
          rows) behaves identically to a blank xlsx cell.
        - Leading '=' is rejected for parity with the xlsx formula guard.
        - Per-cell length is capped at _MAX_PATTERN_CELL_LEN.

        Authoring note: a regex value containing a comma (e.g. \\d{1,3}) must be
        quoted in the CSV ("\\d{1,3}") — the csv module unquotes it correctly.
        utf-8-sig transparently strips a BOM written by Excel's "Save as CSV".
        """
        rows: list = []
        try:
            with open(filepath, newline='', encoding='utf-8-sig') as fh:
                for line_no, raw in enumerate(csv.reader(fh), start=1):
                    row_values = []
                    for col_idx, val in enumerate(raw):
                        if val == '':
                            row_values.append(None)
                            continue
                        if val.startswith('='):
                            coord = f'{get_column_letter(col_idx + 1)}{line_no}'
                            raise SecurityError(
                                f'Formulas are not allowed in pattern files. '
                                f'Cell {coord} contains: {val!r}  '
                                f'Replace it with a plain text value.'
                            )
                        if len(val) > _MAX_PATTERN_CELL_LEN:
                            coord = f'{get_column_letter(col_idx + 1)}{line_no}'
                            raise SecurityError(
                                f'Pattern file cell {coord} value is too long '
                                f'({len(val)} chars, limit is {_MAX_PATTERN_CELL_LEN}). '
                                f'Pattern values should be short identifiers or regex patterns.'
                            )
                        row_values.append(val)
                    rows.append(row_values)
        except UnicodeDecodeError as exc:
            raise SecurityError(
                f'Pattern CSV {filepath!r} is not valid UTF-8 text: {exc}'
            )
        return rows

    def _read_and_validate(self, ws) -> list:
        """
        Read all rows from the pattern worksheet, enforcing that every cell is
        either empty or a plain string.  Formulas, numbers, dates, and booleans
        are all rejected — the pattern file is a configuration document, not a
        spreadsheet.
        """
        rows = []
        for row in ws.iter_rows():
            row_values = []
            for cell in row:
                val = cell.value
                if val is None:
                    row_values.append(None)
                    continue

                # Formulas are never allowed in pattern files
                if cell.data_type == 'f' or (isinstance(val, str) and val.startswith('=')):
                    raise SecurityError(
                        f'Formulas are not allowed in pattern files. '
                        f'Cell {cell.coordinate} contains: {val!r}  '
                        f'Replace it with a plain text value.'
                    )

                # Only plain strings are accepted
                if not isinstance(val, str):
                    raise SecurityError(
                        f'Pattern file cells must contain plain text only. '
                        f'Cell {cell.coordinate} contains a {type(val).__name__} value: {val!r}  '
                        f'All values in a pattern file must be strings.'
                    )

                # Guard against excessively long values
                if len(val) > _MAX_PATTERN_CELL_LEN:
                    raise SecurityError(
                        f'Pattern file cell {cell.coordinate} value is too long '
                        f'({len(val)} chars, limit is {_MAX_PATTERN_CELL_LEN}). '
                        f'Pattern values should be short identifiers or regex patterns.'
                    )

                row_values.append(val)
            rows.append(row_values)
        return rows

    def _pad(self, row, length=10) -> list:
        """Ensure a row list has at least `length` elements."""
        row = list(row)
        while len(row) < length:
            row.append(None)
        return row

    def _check_comment_zone(self, row, start_idx: int, row_num: int, context: str) -> None:
        """Validate the trailing cells (from start_idx) of a non-table row.

        The trailing cells may be empty, or a comment: a cell whose text starts
        with '#' begins a comment that runs to the end of the row. Any non-empty
        cell that is not a comment (and not already inside one) is a mistake —
        fail fast so a value typed into the wrong column is never silently lost.
        Comments are NOT processed for table rows (there '#' is a literal value).
        """
        in_comment = False
        for j in range(start_idx, len(row)):
            val = row[j]
            if val is None or in_comment:
                continue
            if str(val).lstrip().startswith('#'):
                in_comment = True
                continue
            raise PatternError(
                f"Unexpected content {val!r} in column {get_column_letter(j + 1)} "
                f"at pattern row {row_num} ({context}). Columns after the {context} "
                f"fields may only contain a comment starting with '#'."
            )

    def _apply_global_config(self, row, config: Config):
        key, val = row[1], row[2]
        if key == 'read.direction' and val:
            config.read_direction = str(val)
        elif key == 'currency.sign' and val:
            config.currency_sign = str(val)
        elif key == 'empty.aliases' and val:
            config.empty_aliases.append(str(val))
        elif key == 'ignore.case' and val is not None:
            config.ignore_case = _truthy(val)

    def _parse_field(self, row, role: str = 'var', row_num: int | None = None) -> FieldDef:
        where = f' at pattern row {row_num}' if row_num is not None else ''
        name  = str(row[1]) if row[1] else ''
        if not name:
            raise PatternError(
                f"A {role}: definition{where} has no field name. "
                f"Expected:  {role}: | <name> | <type> | <regex>"
            )
        type_ = str(row[2]) if row[2] else 'string'
        if type_ not in _VALID_FIELD_TYPES:
            raise PatternError(
                f"Unknown field type {type_!r} for {role}: {name!r}{where}. "
                f"Valid types: {', '.join(sorted(_VALID_FIELD_TYPES))}."
            )
        regex = str(row[3]) if row[3] else '.*'
        check_regex_safety(regex, field_name=name)
        return FieldDef(name=name, type=type_, regex=regex, role=role)

    # backward-compat alias
    def _parse_def(self, row) -> FieldDef:
        return self._parse_field(row, role='var')
