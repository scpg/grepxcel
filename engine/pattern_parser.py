import openpyxl
from .models import Config, FieldDef, TemplateColumn, TemplateRow, CellInstruction, TableInstruction


class PatternParser:
    def parse(self, filepath: str) -> tuple:
        """
        Parse a pattern Excel file.
        Returns (global_config, defs, start_sequence).
        """
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active

        rows = [[cell.value for cell in row] for row in ws.iter_rows()]

        global_config = Config()
        defs = {}
        start_sequence = []

        i = 0
        in_start = False

        while i < len(rows):
            row = self._pad(rows[i])
            col_a = row[0]

            if col_a == 'END:':
                break

            if col_a == 'START:':
                in_start = True
                i += 1
                continue

            if not in_start:
                if col_a == 'config:':
                    self._apply_global_config(row, global_config)
                elif col_a == 'def:':
                    fd = self._parse_def(row)
                    defs[fd.name] = fd
                i += 1
                continue

            # Inside START: section
            if col_a and col_a.startswith('cell:'):
                mult = col_a.split(':', 1)[1]
                field = row[1] or 'IGNORE'
                start_sequence.append(CellInstruction(multiplicity=mult, field=str(field)))
                i += 1

            elif col_a and col_a.startswith('table:'):
                mult = col_a.split(':', 1)[1]
                table_config = Config(
                    read_direction=global_config.read_direction,
                    currency_sign=global_config.currency_sign,
                    empty_aliases=list(global_config.empty_aliases),
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
                        i += 1
                        continue

                    # HEADER / SPLITTER / DATA / FOOTER row
                    if sub_b and ':' in str(sub_b):
                        row_type, row_mult = str(sub_b).rsplit(':', 1)
                        cols_raw = list(sub[2:])
                        while cols_raw and cols_raw[-1] is None:
                            cols_raw.pop()
                        columns = [
                            TemplateColumn(field=str(v) if v is not None else 'EMPTY')
                            for v in cols_raw
                        ]
                        template_rows.append(TemplateRow(
                            row_type=row_type,
                            multiplicity=row_mult,
                            columns=columns,
                        ))

                    i += 1

                start_sequence.append(TableInstruction(
                    multiplicity=mult,
                    config=table_config,
                    rows=template_rows,
                ))

            else:
                i += 1

        return global_config, defs, start_sequence

    def _pad(self, row, length=10) -> list:
        """Ensure a row list has at least `length` elements."""
        row = list(row)
        while len(row) < length:
            row.append(None)
        return row

    def _apply_global_config(self, row, config: Config):
        key, val = row[1], row[2]
        if key == 'read.direction' and val:
            config.read_direction = str(val)
        elif key == 'currency.sign' and val:
            config.currency_sign = str(val)
        elif key == 'empty.aliases' and val:
            config.empty_aliases.append(str(val))

    def _parse_def(self, row) -> FieldDef:
        name = str(row[1]) if row[1] else ''
        type_ = str(row[2]) if row[2] else 'string'
        regex = str(row[3]) if row[3] else '.*'
        return FieldDef(name=name, type=type_, regex=regex)
