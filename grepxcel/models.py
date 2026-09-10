from dataclasses import dataclass, field

# Valid values for lbl: / var: match mode (global config and per-field override).
# 're' is a short alias for 'regexp' — normalised to 'regexp' by the parser.
LBL_MATCH_MODES = frozenset({'literal', 'glob', 'regexp'})
VAR_MATCH_MODES = frozenset({'literal', 'glob', 'regexp'})


@dataclass
class Config:
    read_direction: str = 'LR'
    currency_sign: str = '€'
    empty_aliases: list = field(default_factory=list)
    ignore_case: bool = False  # when True, all regex/literal/glob matching is case-insensitive
    trim_whitespace: bool = False  # when True, strip leading/trailing spaces before matching/extraction
    pattern_version: int = 1   # pattern-format/semantics generation (absent ⇒ 1)
    pattern_version_explicit: bool = False  # True if the pattern declared it
    lbl_match: str = 'literal'  # how lbl: patterns are matched: literal | glob | regexp
    var_match: str = 'regexp'  # default var: column-D mode: regexp | glob | literal
    unknown_config_keys: list = field(default_factory=list)  # unrecognised config: keys


@dataclass
class FieldDef:
    name: str
    type: str        # string | integer | currency | date | datetime | timestamp
    regex: str
    role: str = 'var'              # 'var' (extract → output) or 'lbl' (anchor only, never output)
    lbl_match: str | None = None   # per-field lbl mode override; None = use Config.lbl_match
    var_mode: str | None = None    # per-field var mode; None/'regexp' = regex (default);
                                   # 'literal' | 'glob' = non-regex pattern in column D
    required: bool = False         # True → not-null/not-empty: fatal if value is empty/null
    trim_whitespace: bool = False  # True → strip leading/trailing whitespace before match/extract


@dataclass
class TemplateColumn:
    field: str  # field name, 'EMPTY', or 'IGNORE'


@dataclass
class TemplateRow:
    row_type: str      # HEADER | SPLITTER | DATA | FOOTER | SKIP_IF
    multiplicity: str  # '1', '*', '{n,m}', or '' for SKIP_IF/HEADER/FOOTER
    columns: list      # list[TemplateColumn]
    min_rows: int = 0            # DATA:{n,m} only — minimum total physical rows
    max_rows: int | None = None  # DATA:{n,m} only — maximum total physical rows


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


@dataclass
class SeekInstruction:
    target: str  # A1-notation cell address, e.g. 'G5' — cursor repositions here without reading


@dataclass
class DirectionInstruction:
    direction: str  # 'LR' or 'TD' — switches the scalar scan direction from this point on
