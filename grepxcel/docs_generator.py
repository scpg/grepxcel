"""
DocsGenerator — writes pattern-reference.xlsx and grepxcel-guide.docx.

pattern-reference.xlsx  — two sheets:
  guide             Quick-start guide for new users (active sheet)
  pattern-reference Self-documenting pattern syntax reference

grepxcel-guide.docx — Word version of the guide (for non-Excel users).

Colour coding in the reference sheet:
  doc:    — yellow   (inline comment, ignored by engine)
  lbl:    — blue     (anchor label, never in output JSON)
  var:    — green    (variable, extracted to output JSON)
  config: — orange   (global setting)
  cell:   — lavender (cell extraction instruction)
  table:  — lavender (table extraction instruction)
  START:/END: — grey (section markers)
  HEADER/DATA/FOOTER — light lavender (table template rows)
"""

import datetime
import os
import sys
import zipfile

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment


def _pin_core_timestamps(path: str, ts: datetime.datetime) -> None:
    """Make a ZIP-based Office file byte-reproducible.

    openpyxl stamps both core.xml's <modified> AND every zip member's
    mod-time with now() on save, so we patch core.xml and rewrite every
    member with a fixed mod-time.  Works for both .xlsx and .docx.
    """
    import re

    ts_iso = ts.strftime('%Y-%m-%dT%H:%M:%SZ')
    fixed_date = (ts.year, ts.month, ts.day, ts.hour, ts.minute, ts.second)

    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
    core = data.get('docProps/core.xml')
    if core is not None:
        text = core.decode('utf-8')
        for tag in ('created', 'modified'):
            text = re.sub(rf'(<dcterms:{tag}[^>]*>)[^<]*(</dcterms:{tag}>)',
                          rf'\g<1>{ts_iso}\g<2>', text)
        data['docProps/core.xml'] = text.encode('utf-8')

    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zi = zipfile.ZipInfo(info.filename, date_time=fixed_date)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.internal_attr = info.internal_attr
            zi.create_system = info.create_system
            zout.writestr(zi, data[info.filename])
    os.replace(tmp, path)


# Row colours (ARGB hex, no leading '#')
_FILL = {
    'doc':     PatternFill('solid', fgColor='FFFFD0'),  # yellow
    'lbl':     PatternFill('solid', fgColor='D0E8FF'),  # blue
    'var':     PatternFill('solid', fgColor='D0FFD0'),  # green
    'config':  PatternFill('solid', fgColor='FFE8D0'),  # orange
    'cell':    PatternFill('solid', fgColor='F0E0FF'),  # lavender
    'table':   PatternFill('solid', fgColor='F0E0FF'),  # lavender
    'marker':  PatternFill('solid', fgColor='E0E0E0'),  # grey  (START/END)
    'tmpl':    PatternFill('solid', fgColor='F8F0FF'),  # light lavender
    'section': PatternFill('solid', fgColor='E8E8E8'),  # light grey (guide sections)
}

_BOLD = Font(bold=True)
_TITLE = Font(bold=True, size=18, color='1F3864')
_H2 = Font(bold=True, size=12, color='2F5597')
_CODE = Font(name='Courier New', size=9)


def _row(ws, row_num: int, cells: list, fill_key: str) -> None:
    fill = _FILL.get(fill_key)
    for col, val in enumerate(cells, start=1):
        c = ws.cell(row=row_num, column=col, value=val)
        if fill:
            c.fill = fill


class DocsGenerator:
    def write(self, output_dir: str) -> list[str]:
        """Generate pattern-reference.xlsx and grepxcel-guide.docx in output_dir."""
        from . import __version__
        os.makedirs(output_dir, exist_ok=True)
        xlsx_path = os.path.join(output_dir, 'pattern-reference.xlsx')
        docx_path = os.path.join(output_dir, 'grepxcel-guide.docx')

        epoch = os.environ.get('SOURCE_DATE_EPOCH')
        if epoch:
            ts = datetime.datetime.fromtimestamp(int(epoch), datetime.timezone.utc)
        else:
            ts = datetime.datetime(2020, 1, 1, tzinfo=datetime.timezone.utc)

        meta = {
            'version': __version__,
            'generated': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M UTC'),
            'executable': os.path.abspath(sys.argv[0]),
        }

        self._write_xlsx(xlsx_path, ts, meta)
        self._write_docx(docx_path, ts, meta)
        return [xlsx_path, docx_path]

    # ── xlsx ──────────────────────────────────────────────────────────────────

    def _write_xlsx(self, output_path: str, ts: datetime.datetime, meta: dict) -> None:
        wb = openpyxl.Workbook()

        ws_guide = wb.active
        ws_guide.title = 'guide'
        self._fill_guide_sheet(ws_guide, meta)

        ws_ref = wb.create_sheet('pattern-reference')
        self._fill_reference_sheet(ws_ref, meta)

        wb.save(output_path)
        _pin_core_timestamps(output_path, ts)

    def _fill_guide_sheet(self, ws, meta: dict) -> None:
        ws.column_dimensions['A'].width = 24
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 44

        r = 1

        def row(cells, fill_key='', font=None):
            nonlocal r
            _row(ws, r, cells, fill_key)
            if font:
                ws.cell(r, 1).font = font
            r += 1

        def section(label):
            nonlocal r
            for col in range(1, 4):
                ws.cell(r, col).fill = _FILL['section']
            ws.cell(r, 1, label).font = _H2
            r += 1

        def blank():
            nonlocal r
            r += 1

        # Title
        ws.cell(r, 1, 'grepxcel guide').font = _TITLE
        r += 1
        ws.cell(r, 1, 'Pattern-based data extraction from Excel files')
        r += 1
        blank()

        # What is grepxcel?
        section('WHAT IS GREPXCEL?')
        row(['', 'grepxcel reads structured data from Excel files using a pattern file.'])
        row(['', 'You describe what to look for — cell values, table rows, field names —'])
        row(['', 'and grepxcel attempts to extract them across any number of files.'])
        row(['', 'Output is JSON: ready for databases, APIs, or data pipelines.'])
        row(['', 'When data does not match the pattern, grepxcel reports it and stops — no silent errors.'])
        blank()

        # Quick start
        section('QUICK START')
        row(['1  Install grepxcel',
             'pip install grepxcel'],
            font=_BOLD)
        ws.cell(r - 1, 2).font = _CODE
        row(['2  Generate ready-to-run examples',
             'grepxcel generate-examples -o examples/'])
        ws.cell(r - 1, 2).font = _CODE
        row(['3  Run your first extraction',
             'grepxcel extract -p examples/01_simple_invoice/pattern.xlsx'
             '  examples/01_simple_invoice/data.xlsx'])
        ws.cell(r - 1, 2).font = _CODE
        row(['4  Write output to files',
             'grepxcel extract -p pattern.xlsx data.xlsx -o output/'])
        ws.cell(r - 1, 2).font = _CODE
        row(['5  Build your own pattern',
             'Use the web wizard (see below) or open the pattern-reference sheet in this file'])
        blank()

        # Wizard
        section('WIZARD — VISUAL PATTERN BUILDER')
        row(['Browser wizard (mouse)',
             "grepxcel web-wizard data.xlsx",
             "requires: pip install 'grepxcel[web]'"])
        ws.cell(r - 1, 2).font = _CODE
        row(['Pre-load an existing pattern', 'grepxcel web-wizard data.xlsx -p pattern.xlsx'])
        ws.cell(r - 1, 2).font = _CODE
        row(['Custom port / no auto-open', 'grepxcel web-wizard data.xlsx --port 9000 --no-browser'])
        ws.cell(r - 1, 2).font = _CODE
        blank()

        # Commands
        section('AVAILABLE COMMANDS')
        cmds = [
            ('extract',           'Extract data from Excel files using a pattern',       'grepxcel extract -p pattern.xlsx data.xlsx'),
            ('validate-pattern',  'Check a pattern file without running extraction',     'grepxcel validate-pattern pattern.xlsx'),
            ('draft',             'AI-powered pattern drafter (optional)',               "pip install 'grepxcel[suggest]'  then  grepxcel draft data.xlsx"),
            ('web-wizard',        'Browser-based visual pattern builder',               "grepxcel web-wizard data.xlsx  [pip install 'grepxcel[web]']"),
            ('test',              'Run pattern against a folder — pass/warn/fail report','grepxcel test -p pattern.xlsx samples/'),
            ('lint',              'Inspect an Excel file before writing a pattern',      'grepxcel lint data.xlsx'),
            ('schema',            'Generate JSON Schema from a pattern file',            'grepxcel schema pattern.xlsx'),
            ('docs',              'Regenerate this guide + grepxcel-guide.docx',         'grepxcel docs -o /output/directory/'),
            ('generate-examples', 'Write 4 ready-to-run example files to a directory',  'grepxcel generate-examples -o examples/'),
            ('doctor',            'Check your environment is ready',                    'grepxcel doctor'),
            ('quickstart',        'Guided tutorial in your terminal',                    'grepxcel quickstart'),
        ]
        for cmd, desc, example in cmds:
            row([cmd, desc, example])
            ws.cell(r - 1, 1).font = _BOLD
            ws.cell(r - 1, 3).font = _CODE
        blank()

        # Colour key
        section('PATTERN FILE COLOUR KEY')
        row(['doc:',           'yellow',    'Comment rows — ignored by the engine'],               'doc')
        row(['lbl:',           'blue',      'Anchor labels — matched for position, never extracted'], 'lbl')
        row(['var:',           'green',     'Variables — extracted and written to output JSON'],    'var')
        row(['config:',        'orange',    'Global settings (read.direction, currency.sign, ...)'],'config')
        row(['cell: / table:', 'lavender',  'Instructions — cursor movement and table scanning'],   'cell')
        blank()

        # TAB completion
        section('SHELL TAB COMPLETION')
        row(['', 'Add once to ~/.bashrc or ~/.zshrc to enable TAB completion for all commands and flags:'])
        row(['', 'eval "$(register-python-argcomplete grepxcel)"'])
        ws.cell(r - 1, 2).font = _CODE
        blank()

        # Regenerate
        section('REGENERATING THIS GUIDE')
        row(['', 'grepxcel docs'])
        ws.cell(r - 1, 2).font = _CODE
        row(['', 'grepxcel docs -o /your/output/directory/'])
        ws.cell(r - 1, 2).font = _CODE
        row(['', 'Both pattern-reference.xlsx and grepxcel-guide.docx are regenerated together.'])
        blank()

        # Generated by
        section('GENERATED BY')
        row(['', f'grepxcel {meta["version"]}'])
        row(['', f'Generated:  {meta["generated"]}'])
        row(['', f'Executable: {meta["executable"]}'])

    def _fill_reference_sheet(self, ws, meta: dict) -> None:
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 28
        ws.column_dimensions['E'].width = 50

        r = 1

        def row(cells, fill_key=''):
            nonlocal r
            _row(ws, r, cells, fill_key)
            r += 1

        def blank():
            nonlocal r
            r += 1

        # ── Header ───────────────────────────────────────────────────────────
        row(['doc:', 'AUTO-GENERATED — do not edit this file directly.'], 'doc')
        row(['doc:', 'To regenerate: grepxcel docs'], 'doc')
        row(['doc:', 'To write to a specific directory: grepxcel docs -o /your/path/'], 'doc')
        row(['doc:', 'A current copy is kept at docs/pattern-reference.xlsx in the project repository.'], 'doc')
        blank()

        # ── Title ─────────────────────────────────────────────────────────────
        ws.cell(r, 1, 'grepxcel pattern reference').font = _BOLD
        r += 1
        blank()

        # ── Colour key (proper coloured section) ──────────────────────────────
        ws.cell(r, 1, 'Column colour key:').font = _BOLD
        r += 1
        row(['doc:',           'yellow',   '', 'Comment rows — ignored by the engine'],               'doc')
        row(['lbl:',           'blue',     '', 'Anchor labels — matched for position, never extracted'], 'lbl')
        row(['var:',           'green',    '', 'Variables — extracted and written to output JSON'],    'var')
        row(['config:',        'orange',   '', 'Global settings (read.direction, currency.sign, ...)'],'config')
        row(['cell: / table:', 'lavender', '', 'Instructions — cursor movement and table scanning'],   'cell')
        blank()

        # ── CONFIG section ────────────────────────────────────────────────────
        row(['doc:', '', '', '', 'Config rows set global options. Must appear before lbl:/var: rows.'], 'doc')
        row(['config:', 'pattern.version', '1', '', 'Pattern-format version (optional; absent = 1). Engine errors if newer than it understands.'], 'config')
        row(['config:', 'read.direction', 'LR', '', 'LR = left-to-right scan (default). TD = top-to-bottom.'], 'config')
        row(['config:', 'currency.sign', '€', '', 'Symbol used when parsing currency values.'], 'config')
        row(['config:', 'ignore.case', 'no', '', 'yes = match all regexes case-insensitively. Default no.'], 'config')
        blank()

        # ── LBL section ───────────────────────────────────────────────────────
        row(['doc:', '', '', '', 'lbl: defines an anchor label. Matched for position only — NEVER written to output JSON.'], 'doc')
        row(['doc:', '', '', '', 'Use lbl: for literal text like "Invoice No:" or table column headers like "Product".'], 'doc')
        row(['lbl:', 'po_label',     'string',   'PO Number:',  'Matches the literal text "PO Number:" in the sheet.'], 'lbl')
        row(['lbl:', 'date_label',   'string',   'Date:',       'Matches "Date:".'], 'lbl')
        row(['lbl:', 'vendor_label', 'string',   'Vendor:',     'Matches "Vendor:".'], 'lbl')
        row(['lbl:', 'col_item',     'string',   'Item',        'Table column header anchor.'], 'lbl')
        row(['lbl:', 'col_qty',      'string',   'Qty',         'Table column header anchor.'], 'lbl')
        row(['lbl:', 'col_price',    'string',   'Unit Price',  'Table column header anchor.'], 'lbl')
        row(['lbl:', 'col_total',    'string',   'Total',       'Table column header anchor.'], 'lbl')
        blank()

        # ── VAR section ───────────────────────────────────────────────────────
        row(['doc:', '', '', '', 'var: defines a data field. Extracted and written to output JSON.'], 'doc')
        row(['doc:', '', '', '', 'Dot notation creates nested JSON: po.number → {"po": {"number": ...}}'], 'doc')
        row(['doc:', '', '', '', 'All var: fields in one table DATA row must share the same group prefix.'], 'doc')
        row(['doc:', '', '', '', 'Types: string text integer number/float/decimal currency percentage boolean time duration date datetime'], 'doc')
        row(['doc:', '', '', '', 'Regex: Python re.fullmatch pattern. Leave blank (or .*) to accept any value of the declared type.'], 'doc')
        row(['doc:', '', '', '', 'Comments: a doc:/info: row is a whole-line comment. On config/var/lbl/cell/START rows,'], 'doc')
        row(['doc:', '', '', '', 'a cell starting with # AFTER the real columns is a trailing comment to end of row'], 'doc')
        row(['doc:', '', '', '', '(var:/lbl: → col E+, cell: → col C+, config: → col D+). NOT in table: rows (# is literal there).'], 'doc')
        row(['var:', 'po.number',    'string',   r'PO-[0-9]{4}',    'Matches e.g. "PO-2026". Output key: {"po": {"number": ...}}'], 'var')
        row(['var:', 'po.date',      'date',     '',                 'Any date value. Regex left empty → defaults to .*'], 'var')
        row(['var:', 'vendor.name',  'string',   '.+',              'Any non-empty string.'], 'var')
        row(['var:', 'line.item',    'string',   '.+',              'Table DATA field. Group prefix "line" → {"line": [...]}'], 'var')
        row(['var:', 'line.qty',     'integer',  r'[1-9][0-9]*',    'Positive integer.'], 'var')
        row(['var:', 'line.price',   'currency', '.*',              'Any currency value.'], 'var')
        row(['var:', 'line.total',   'currency', '.*',              'Any currency value.'], 'var')
        row(['var:', 'footer.label', 'string',   'Grand Total',     'FOOTER field — matches the literal "Grand Total".'], 'var')
        row(['var:', 'footer.value', 'currency', '.*',              'FOOTER value field.'], 'var')
        blank()

        # ── Column A modifier examples ────────────────────────────────────────
        row(['doc:', '', '', '', 'Column A modifiers — order-independent, colon-separated. Add after var: or lbl:'], 'doc')
        row(['doc:', '', '', '', 'not-null / not-empty (synonyms): fatal error if value is empty/null (always, not just with --strict)'], 'doc')
        row(['doc:', '', '', '', 'var:glob → column D is a shell glob (PROD-* matches PROD-42); type check still runs'], 'doc')
        row(['doc:', '', '', '', 'var:literal → column D is exact string (special chars are literal, not regex)'], 'doc')
        row(['doc:', '', '', '', 'var:re / var:regexp → explicit alias for default regex mode; lbl:re → alias for lbl:regexp'], 'doc')
        row(['doc:', '', '', '', 'trim-whitespace → strip leading/trailing spaces before match and in extracted JSON value'], 'doc')
        row(['doc:', '', '', '', 'Modifiers combine freely: var:not-null:glob, var:trim-whitespace, lbl:trim-whitespace:not-null'], 'doc')
        row(['var:not-null',          'invoice.number', 'string', r'INV-\d+', 'Required regex field — fatal if empty.'], 'var')
        row(['var:glob',              'sku',            'string', 'PROD-*',   'Glob match: PROD-42, PROD-XYZ, ...'], 'var')
        row(['var:literal',           'status',         'string', 'Active',   'Exact string match — "Active" only.'], 'var')
        row(['var:trim-whitespace',   'company',        'string', '.*',       'Strip spaces: " Acme Corp  " → "Acme Corp".'], 'var')
        row(['var:not-null:re',       'code',           'string', '[A-Z]{3}', 'Explicit regex, required.'], 'var')
        row(['lbl:not-null',          'inv_label',      'string', 'Invoice:', 'Anchor must be present or fatal.'], 'lbl')
        row(['lbl:trim-whitespace',   'header',         'string', 'Date',     'Strip spaces before matching the anchor.'], 'lbl')
        blank()

        # ── START section ─────────────────────────────────────────────────────
        row(['doc:', '', '', '', 'Everything between START: and END: defines the extraction order.'], 'doc')
        row(['doc:', '', '', '', 'cell:next  reads the next non-empty cell (alias: cell:1). Scans in read.direction order.'], 'doc')
        row(['doc:', '', '', '', 'cell:B5    jumps directly to cell B5 (absolute A1-notation reference).'], 'doc')
        row(['doc:', '', '', '', 'seek:G5    repositions the cursor to G5 WITHOUT reading it. Next cell:next starts from there.'], 'doc')
        row(['doc:', '', '', '', 'dir:LR     switches scan direction to left-to-right from here on (dir:TD = top-down).'], 'doc')
        row(['doc:', '', '', '', 'table:*    finds all instances of a repeating table block.'], 'doc')
        row(['START:'], 'marker')

        row(['cell:A1', 'po_label',     '', '', 'Jump to A1 and read "PO Number:" anchor (lbl: field — not in output).'], 'cell')
        row(['cell:next', 'po.number',  '', '', 'Read next non-empty cell after A1 → the PO number.'], 'cell')
        row(['cell:C1', 'date_label',   '', '', 'Jump to C1 and read "Date:" anchor.'], 'cell')
        row(['cell:next', 'po.date',    '', '', 'Read next non-empty cell after C1 → the PO date.'], 'cell')
        row(['cell:next', 'vendor_label', '', '', 'Read next non-empty cell → "Vendor:" anchor.'], 'cell')
        row(['cell:next', 'vendor.name',  '', '', 'Extract the vendor name.'], 'cell')
        row(['doc:', '', '', '', 'Use IGNORE to skip a non-empty cell without defining a field for it.'], 'doc')
        row(['cell:next', 'IGNORE',     '', '', 'Skip one non-empty cell without capturing it.'], 'cell')
        blank()

        row(['doc:', '', '', '', 'table:* matches 0-or-more mini-table instances in greedy order.'], 'doc')
        row(['doc:', '', '', '', 'Use table:1 when exactly one instance is expected.'], 'doc')
        row(['table:*'], 'table')
        row([None, 'HEADER:1', 'col_item', 'col_qty', 'col_price', 'col_total'], 'tmpl')
        row([None, 'DATA:*',   'line.item', 'line.qty', 'line.price', 'line.total'], 'tmpl')
        row([None, 'FOOTER:1', 'footer.label', 'IGNORE', 'IGNORE', 'footer.value'], 'tmpl')
        row(['doc:', '', '', '', 'HEADER/FOOTER are strict: wrong or missing value = no match.'], 'doc')
        row(['doc:', '', '', '', 'DATA rows are lenient: wrong value emits a warning but extraction continues.'], 'doc')
        row(['doc:', '', '', '', 'EMPTY column keyword: cell must be blank. IGNORE: consume without matching.'], 'doc')
        row(['doc:', '', '', '', '── BOUNDED DATA + SKIP_IF (fixed-slot templates) ───────────────────────'], 'doc')
        row(['doc:', '', '', '', 'Use DATA:{n,m} when the template pre-allocates a fixed number of rows.'], 'doc')
        row(['doc:', '', '', '', 'n = min total physical rows expected, m = max total rows to scan.'], 'doc')
        row(['doc:', '', '', '', 'SKIP_IF: rows matching the condition are silently excluded from output'], 'doc')
        row(['doc:', '', '', '', 'but still count toward the {n,m} bounds.'], 'doc')
        row(['doc:', '', '', '', 'SKIP_IF is valid with DATA:{n,m} (rows count toward bounds) and DATA:* (rows filtered, scan continues).'], 'doc')
        row(['table:1'], 'table')
        row([None, 'HEADER:1', 'col_item', 'col_qty'], 'tmpl')
        row([None, 'SKIP_IF',  'EMPTY',    'IGNORE'], 'tmpl')
        row([None, 'DATA:{0,15}', 'line.item', 'line.qty'], 'tmpl')
        row([None, 'FOOTER:1', 'lbl_total', 'inv.total'], 'tmpl')
        blank()

        row(['END:'], 'marker')
        blank()

        # ── Quick reference table ──────────────────────────────────────────────
        row(['doc:', '', '', '', '── QUICK REFERENCE ──────────────────────────────────────────────'], 'doc')
        ref = [
            ('doc:',     'doc: | free text',           'Inline comment. Ignored by engine.'),
            ('config:',  'config: | key | value',      'Global setting. Keys: read.direction, currency.sign, empty.aliases, ignore.case'),
            ('lbl:',     'lbl: | name | type | regex', 'Anchor label. Matched but never in output.'),
            ('var:',     'var: | name | type | regex', 'Extracted variable. Use dot notation for nesting.'),
            ('cell:next','cell:next | FieldName',      'Read next non-empty cell (alias: cell:1).'),
            ('cell:next','cell:next | IGNORE',         'Skip next non-empty cell without capturing.'),
            ('cell:next','cell:B5 | FieldName',        'Jump directly to B5 (absolute reference). Ordering must be forward.'),
            ('seek:',    'seek:G5',                    'Reposition cursor to G5 without reading it. Resets abs-ref ordering constraint.'),
            ('table:*',  'table:*  (or table:1)',      'Begin a repeating table block.'),
            ('HEADER:N', ' | HEADER:1 | F1 | F2',     'Strict header row template (col A must be blank).'),
            ('DATA:*',      ' | DATA:* | F1 | F2',           'Data row template (greedy, lenient validation).'),
            ('DATA:{n,m}',  ' | DATA:{0,15} | F1 | F2',      'Bounded data: scan at most m physical rows, warn if < n.'),
            ('SKIP_IF',     ' | SKIP_IF | EMPTY | IGNORE',    'Skip row if non-IGNORE columns match. Valid with DATA:{n,m} and DATA:*.'),
            ('FOOTER:N',    ' | FOOTER:1 | F1 | F2',          'Strict footer row template.'),
            ('SPLITTER',    ' | SPLITTER:1',                   'All columns must be blank (separator row).'),
        ]
        for keyword, syntax, description in ref:
            fill_key = keyword.rstrip(':*1').lower()
            if fill_key.startswith('header') or fill_key.startswith('data') or \
               fill_key.startswith('footer') or fill_key.startswith('splitter'):
                fill_key = 'tmpl'
            row([keyword, syntax, '', description], fill_key)

        blank()
        row(['doc:', '', '', '── GENERATED BY ──────────────────────────────────────────────────'], 'doc')
        row(['doc:', '', '', f'grepxcel {meta["version"]}'], 'doc')
        row(['doc:', '', '', f'Generated:  {meta["generated"]}'], 'doc')
        row(['doc:', '', '', f'Executable: {meta["executable"]}'], 'doc')

    # ── docx ──────────────────────────────────────────────────────────────────

    def _write_docx(self, output_path: str, ts: datetime.datetime, meta: dict) -> None:
        ts_iso = ts.strftime('%Y-%m-%dT%H:%M:%SZ')

        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml"'
            ' ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '</Types>'
        )

        rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
            ' Target="word/document.xml"/>'
            '<Relationship Id="rId2"'
            ' Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"'
            ' Target="docProps/core.xml"/>'
            '</Relationships>'
        )

        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
            ' Target="styles.xml"/>'
            '</Relationships>'
        )

        core_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties'
            ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
            ' xmlns:dcterms="http://purl.org/dc/terms/"'
            ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>grepxcel</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{ts_iso}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{ts_iso}</dcterms:modified>'
            '</cp:coreProperties>'
        )

        styles_xml = self._docx_styles()
        document_xml = self._docx_document(meta)

        fixed_date = (ts.year, ts.month, ts.day, ts.hour, ts.minute, ts.second)
        members = {
            '[Content_Types].xml': content_types,
            '_rels/.rels': rels,
            'docProps/core.xml': core_xml,
            'word/_rels/document.xml.rels': doc_rels,
            'word/styles.xml': styles_xml,
            'word/document.xml': document_xml,
        }

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, content in members.items():
                zi = zipfile.ZipInfo(name, date_time=fixed_date)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(zi, content.encode('utf-8'))

    @staticmethod
    def _docx_styles() -> str:
        W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:styles {W}>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/>'
            '<w:pPr><w:spacing w:after="120"/></w:pPr>'
            '<w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
            '<w:sz w:val="22"/>'
            '</w:rPr>'
            '</w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1">'
            '<w:name w:val="heading 1"/>'
            '<w:basedOn w:val="Normal"/>'
            '<w:pPr>'
            '<w:outlineLvl w:val="0"/>'
            '<w:spacing w:before="360" w:after="120"/>'
            '</w:pPr>'
            '<w:rPr>'
            '<w:b/>'
            '<w:sz w:val="48"/>'
            '<w:color w:val="1F3864"/>'
            '</w:rPr>'
            '</w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading2">'
            '<w:name w:val="heading 2"/>'
            '<w:basedOn w:val="Normal"/>'
            '<w:pPr>'
            '<w:outlineLvl w:val="1"/>'
            '<w:spacing w:before="280" w:after="80"/>'
            '</w:pPr>'
            '<w:rPr>'
            '<w:b/>'
            '<w:sz w:val="28"/>'
            '<w:color w:val="2F5597"/>'
            '</w:rPr>'
            '</w:style>'
            '<w:style w:type="paragraph" w:styleId="Code">'
            '<w:name w:val="Code"/>'
            '<w:basedOn w:val="Normal"/>'
            '<w:pPr>'
            '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
            '<w:ind w:left="360"/>'
            '<w:spacing w:before="0" w:after="60"/>'
            '</w:pPr>'
            '<w:rPr>'
            '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>'
            '<w:sz w:val="18"/>'
            '</w:rPr>'
            '</w:style>'
            '</w:styles>'
        )

    def _docx_document(self, meta: dict) -> str:
        W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

        def esc(s: str) -> str:
            return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        def p(text: str, style: str = 'Normal') -> str:
            return (
                f'<w:p>'
                f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
                f'</w:p>'
            )

        def pbold(label: str, text: str, style: str = 'Normal') -> str:
            return (
                f'<w:p>'
                f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
                f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{esc(label)}</w:t></w:r>'
                f'<w:r><w:t xml:space="preserve">  {esc(text)}</w:t></w:r>'
                f'</w:p>'
            )

        def h1(text: str) -> str:
            return p(text, 'Heading1')

        def h2(text: str) -> str:
            return p(text, 'Heading2')

        def code(text: str) -> str:
            return p(text, 'Code')

        def blank() -> str:
            return '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>'

        parts = [
            h1('grepxcel guide'),
            p('Pattern-based data extraction from Excel files.'),
            blank(),

            h2('What Is grepxcel?'),
            p('grepxcel reads structured data from Excel files using a pattern file.'),
            p('You describe what to look for — cell values, table rows, field names — '
              'and grepxcel attempts to extract them across any number of files.'),
            p('Output is JSON, ready for databases, APIs, or data pipelines.'),
            p('When data does not match the pattern, grepxcel reports it and stops — '
              'it does not silently return partial results.'),
            blank(),

            h2('Quick Start'),
            pbold('1  Install grepxcel', ''),
            code('pip install grepxcel'),
            pbold('2  Generate ready-to-run examples', ''),
            code('grepxcel generate-examples -o examples/'),
            pbold('3  Run your first extraction', ''),
            code('grepxcel extract -p examples/01_simple_invoice/pattern.xlsx'
                 '  examples/01_simple_invoice/data.xlsx'),
            pbold('4  Write output to files', ''),
            code('grepxcel extract -p pattern.xlsx data.xlsx -o output/'),
            pbold('5  Build your own pattern', ''),
            p('    Use the web wizard (below) or open the pattern-reference sheet in pattern-reference.xlsx.'),
            blank(),

            h2('Web Wizard — Visual Pattern Builder'),
            p('The web wizard helps you build a pattern by clicking directly on your Excel file in a browser.'),
            pbold('Start the wizard:', "requires: pip install 'grepxcel[web]'"),
            code('grepxcel web-wizard data.xlsx'),
            pbold('Pre-load an existing pattern:', ''),
            code('grepxcel web-wizard data.xlsx -p pattern.xlsx'),
            pbold('Custom port / no auto-open:', ''),
            code('grepxcel web-wizard data.xlsx --port 9000 --no-browser'),
            blank(),

            h2('Available Commands'),
            pbold('extract',           'Extract data from Excel files using a pattern file.'),
            code('grepxcel extract -p pattern.xlsx data.xlsx'),
            pbold('validate-pattern',  'Check a pattern file without running extraction.'),
            code('grepxcel validate-pattern pattern.xlsx'),
            pbold('draft',             "AI-powered pattern drafter (pip install 'grepxcel[suggest]')."),
            code('grepxcel draft data.xlsx'),
            pbold('web-wizard',        "Browser-based visual pattern builder (pip install 'grepxcel[web]')."),
            code('grepxcel web-wizard data.xlsx'),
            pbold('test',              'Run pattern against a folder — pass/warn/fail report per file.'),
            code('grepxcel test -p pattern.xlsx samples/'),
            pbold('lint',              'Inspect an Excel file before writing a pattern.'),
            code('grepxcel lint data.xlsx'),
            pbold('schema',            'Generate JSON Schema from a pattern file.'),
            code('grepxcel schema pattern.xlsx'),
            pbold('docs',              'Regenerate this guide and pattern-reference.xlsx.'),
            code('grepxcel docs -o /output/directory/'),
            pbold('generate-examples', 'Write 4 ready-to-run example files to a directory.'),
            code('grepxcel generate-examples -o examples/'),
            pbold('doctor',            'Check your environment is ready.'),
            code('grepxcel doctor'),
            pbold('quickstart',        'Guided tutorial in your terminal.'),
            code('grepxcel quickstart'),
            blank(),

            h2('Shell TAB Completion'),
            p('grepxcel includes TAB completion for all subcommands and flags (bash, zsh, fish). '
              'Add this once to your shell config file to activate it:'),
            code('eval "$(register-python-argcomplete grepxcel)"'),
            p('Add the line above to ~/.bashrc or ~/.zshrc, then open a new terminal. '
              'After that, pressing TAB after grepxcel completes subcommands and flags automatically.'),
            blank(),

            h2('Pattern File Format'),
            p('The pattern file is an Excel or CSV file with up to five columns per row:'),
            p('    Column A  row type (doc:, lbl:, var:, config:, cell:, table:, START:, END:)'),
            p('    Column B  field name (dot notation for nesting, e.g. invoice.number)'),
            p('    Column C  type (string, integer, currency, date, boolean, …)'),
            p('    Column D  pattern/value (Python regex, glob, or literal)'),
            p('    Column E  optional comment'),
            p('See the pattern-reference sheet in pattern-reference.xlsx for a '
              'fully worked example with colour coding and detailed annotations.'),
            blank(),

            h2('Regenerating This Guide'),
            p('To regenerate both files in the current directory:'),
            code('grepxcel docs'),
            p('To write to a specific directory:'),
            code('grepxcel docs -o /your/output/directory/'),
            p('Both pattern-reference.xlsx and grepxcel-guide.docx are regenerated together.'),
            blank(),

            h2('Generated By'),
            p(f'grepxcel {meta["version"]}'),
            p(f'Generated:  {meta["generated"]}'),
            p(f'Executable: {meta["executable"]}'),
        ]

        body = ''.join(parts)
        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document {W}>'
            f'<w:body>{body}</w:body>'
            f'</w:document>'
        )
