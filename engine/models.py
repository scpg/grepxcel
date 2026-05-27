from dataclasses import dataclass, field


@dataclass
class Config:
    read_direction: str = 'LR'
    currency_sign: str = '€'
    empty_aliases: list = field(default_factory=list)


@dataclass
class FieldDef:
    name: str
    type: str        # string | integer | currency | date | datetime | timestamp
    regex: str
    role: str = 'var'  # 'var' (extract → output) or 'lbl' (anchor only, never output)


@dataclass
class TemplateColumn:
    field: str  # field name, 'EMPTY', or 'IGNORE'


@dataclass
class TemplateRow:
    row_type: str    # HEADER | SPLITTER | DATA | FOOTER
    multiplicity: str  # '1', '*', or integer string
    columns: list    # list[TemplateColumn]


@dataclass
class CellInstruction:
    multiplicity: str        # '1', 'next', or 'abs'
    field: str               # field name or 'IGNORE'
    target: str | None = None  # A1-notation ref for multiplicity='abs', e.g. 'B5'


@dataclass
class TableInstruction:
    multiplicity: str  # '1' or '*'
    config: Config     # table-level config (may override global)
    rows: list         # list[TemplateRow]
