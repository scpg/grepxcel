import datetime as _datetime
import fnmatch
import re
import warnings

import regex as _re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from .models import Config, CellInstruction, TableInstruction, TemplateRow, SeekInstruction, DirectionInstruction
from .utils import is_empty, validate_type, _MAX_REGEX_INPUT_LEN, _regex_timeout
from .pattern_parser import PatternParser, PatternError
from .logger import Logger, LogRecord, EngineError, cell_ref
from .security import validate_file, validate_pattern_file, SecurityError, DEFAULT_MAX_UNCOMPRESSED_MB


# ── assert: rule evaluation ───────────────────────────────────────────────────

def _run_assert_rules(rules, cells: dict, logger: Logger) -> None:
    """Evaluate all assert: rules against the extracted cells dict.

    A failing assertion logs a WARNING.  A rule referencing unknown or None
    fields is silently skipped (the engine cannot assert about fields that
    weren't extracted).

    Args:
        rules:  list[AssertRule] from the pattern parser.
        cells:  flat {field_name: value} dict from _raw['cells'].
        logger: Logger instance for recording warnings.
    """
    from .assert_eval import evaluate_assert, parse_assert, AssertParseError

    for rule in rules:
        try:
            tree = parse_assert(rule.expression)
            result = evaluate_assert(tree, cells)
        except AssertParseError as exc:
            # Should not happen — parse-time validation already caught this.
            logger.warn_assert(
                f'assert: rule could not be evaluated: {exc}',
                hint=f'Expression: {rule.expression}',
            )
            continue

        if result is None:
            # A field was missing/None — skip silently.
            continue
        if not result:
            msg = rule.message or rule.expression
            logger.warn_assert(
                f'Assertion failed: {msg}',
                hint=f'Expression: {rule.expression!r} evaluated to False',
            )


# ── lbl: matching ─────────────────────────────────────────────────────────────

def _match_lbl(cell_value, pattern: str, mode: str, ignore_case: bool) -> bool:
    """Match a cell value against a label pattern using the configured mode.

    Empty pattern (blank column D) always matches — preserves the "relaxed
    default" regardless of mode.  Modes:
      literal — exact string equality (honours ignore_case)
      glob    — shell wildcards (* = any text incl. newlines, ? = one char)
      regexp  — full Python re.search (current / pre-1.0 behaviour)
    """
    if not pattern:
        return True
    text = str(cell_value) if cell_value is not None else ''
    if mode == 'literal':
        return (text.lower() == pattern.lower()) if ignore_case else (text == pattern)
    if mode == 'glob':
        flags = re.DOTALL | (re.IGNORECASE if ignore_case else 0)
        return bool(re.match(fnmatch.translate(pattern), text, flags))
    # regexp — same ReDoS defenses as var: fields (cell-length cap + timeout)
    flags = _re.IGNORECASE if ignore_case else 0
    text_capped = text[:_MAX_REGEX_INPUT_LEN]
    try:
        return bool(_re.search(pattern, text_capped, flags, timeout=_regex_timeout()))
    except TimeoutError:
        return False


def _resolve_lbl_mode(fd, config) -> str:
    """Return the effective lbl match mode for a FieldDef (per-field wins over global)."""
    return fd.lbl_match if fd.lbl_match is not None else config.lbl_match


def _apply_trim(value, fd, config):
    """Strip leading/trailing whitespace from a string cell value when trim-whitespace
    is active — either per-field (``fd.trim_whitespace``) or globally via
    ``config.trim_whitespace_values``.  Non-string values are returned unchanged."""
    if isinstance(value, str) and (fd.trim_whitespace or config.trim_whitespace_values):
        return value.strip()
    return value


def _validate_field(fd, value, config, max_cell_len: int) -> bool:
    """Validate a cell value against a FieldDef.

    Dispatches on ``fd.role`` and ``fd.var_mode``:

    * ``lbl:``           — pattern match (literal/glob/regexp per resolved mode).
    * ``var:`` (default) — type check + regex in column D via :func:`validate_type`.
    * ``var:literal``    — type check + exact-string (or case-insensitive) match.
    * ``var:glob``       — type check + shell-glob match on the string representation.

    Returns True/False; never raises.
    """
    if fd.role == 'lbl':
        # Trim label cell text before matching when trim_whitespace_labels is active.
        lbl_text = value
        if isinstance(lbl_text, str) and config.trim_whitespace_labels:
            lbl_text = lbl_text.strip()
        return _match_lbl(lbl_text, fd.regex, fd.lbl_match or config.lbl_match,
                          config.ignore_case_labels)
    # var: field — per-field var_mode wins; fall back to config.var_match global default.
    effective_var_mode = fd.var_mode if fd.var_mode is not None else config.var_match
    if effective_var_mode in ('literal', 'glob'):
        # Type check (use '.*' so it always passes the regex part).
        type_ok, _ = validate_type(value, fd.type, '.*', config.currency_sign,
                                   max_cell_len, config.ignore_case_values)
        if not type_ok:
            return False
        # Convert date/datetime to ISO string so var:literal|glob date fields work.
        # str(datetime(2024,1,1)) gives '2024-01-01 00:00:00', not '2024-01-01'.
        if isinstance(value, _datetime.datetime):
            _str_val = value.date().isoformat()
        elif isinstance(value, _datetime.date):
            _str_val = value.isoformat()
        else:
            _str_val = str(value) if value is not None else ''
        return _match_lbl(_str_val, fd.regex, effective_var_mode, config.ignore_case_values)
    # regexp mode (default) — validate_type handles both type and regex.
    ok, _ = validate_type(value, fd.type, fd.regex, config.currency_sign,
                          max_cell_len, config.ignore_case_values)
    return ok


# ── Data-sheet size limits ──────────────────────────────────────────────────────
#
# grepxcel targets normal-sized spreadsheets. Very large sheets are both a
# performance risk and untested territory, so the data sheet's dimensions are
# bounded by conservative defaults. The user can raise them (up to Excel's hard
# maximum) but is warned the tool is not validated at that scale.
DEFAULT_MAX_DATA_ROWS = 2048
DEFAULT_MAX_DATA_COLS = 1024
EXCEL_MAX_ROWS = 1_048_576   # Excel's hard row ceiling
EXCEL_MAX_COLS = 16_384      # Excel's hard column ceiling (XFD)


def _used_extent(ws) -> tuple[int, int]:
    """Return (last_row, last_col) with actual non-empty cell values.

    openpyxl's max_row/max_column includes styled-but-empty cells, which
    inflates the declared dimensions. This scans for the real data boundary.
    """
    last_row = 0
    last_col = 0
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for col_idx, val in enumerate(row, start=1):
            if val is not None:
                last_row = row_idx
                if col_idx > last_col:
                    last_col = col_idx
    return last_row, last_col


def _check_sheet_dimensions(ws, max_rows: int, max_cols: int, logger) -> None:
    """Fatal-error if the worksheet exceeds the configured row/column limits.

    Uses the real used extent (non-empty cells) rather than openpyxl's
    declared max_row/max_column, which can be inflated by formatting on
    empty cells. Warns when declared > used but used is within limits.
    """
    declared_rows = ws.max_row or 0
    declared_cols = ws.max_column or 0

    if declared_rows <= max_rows and declared_cols <= max_cols:
        return

    used_rows, used_cols = _used_extent(ws)

    if used_rows > max_rows or used_cols > max_cols:
        logger.fatal(
            f"Sheet {ws.title!r} is too large for grepxcel's limits: "
            f"{used_rows} row(s) × {used_cols} column(s) of data. "
            f"grepxcel targets normal-sized "
            f"spreadsheets. You can raise the limits with --max-rows / "
            f"--max-columns (up to Excel's maximum of {EXCEL_MAX_ROWS:,} rows and "
            f"{EXCEL_MAX_COLS:,} columns), but the tool has not been tested at that "
            f"scale, so correct behavior is not guaranteed.",
            found=f"{used_rows} rows × {used_cols} columns",
            expected=f"≤ {max_rows} rows and ≤ {max_cols} columns",
        )

    logger._emit(
        severity='WARNING', category='STRUCTURAL',
        message=(
            f"Sheet {ws.title!r} declares {declared_rows} row(s) × "
            f"{declared_cols} column(s) but only {used_rows} × {used_cols} "
            f"contain data (likely empty formatted cells). "
            f"Proceeding with real data extent."
        ),
    )


# ── Output helpers ─────────────────────────────────────────────────────────────

def _range_ref(r1: int, c1: int, r2: int, c2: int) -> str:
    """Row/col bounds → Excel A1-notation range string, e.g. 'B3:E10'."""
    return f'{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}'


def _set_nested(d: dict, dotted_key: str, value) -> None:
    """d['a']['b']['c'] = value  for  dotted_key='a.b.c'."""
    parts = dotted_key.split('.')
    for part in parts[:-1]:
        d = d.setdefault(part, {})
    d[parts[-1]] = value


def _field_local(field_name: str) -> str:
    """Everything after the first dot, or the full name if there is no dot."""
    _, _, rest = field_name.partition('.')
    return rest if rest else field_name


def _field_group(field_name: str) -> str:
    """Everything before the first dot, or the full name if there is no dot."""
    return field_name.split('.', 1)[0]


def _table_group(raw_tables_entry: dict, defs: dict) -> str:
    """
    Derive the top-level output key for a table match from its DATA var: fields.
    Falls back to 'table_<index>' when no var: field has a dot.
    """
    counts: dict[str, int] = {}
    for raw_row in raw_tables_entry.get('data', []):
        for field in raw_row:
            fd = defs.get(field)
            if fd and fd.role == 'var' and '.' in field:
                g = _field_group(field)
                counts[g] = counts.get(g, 0) + 1
    if counts:
        return max(counts, key=lambda k: counts[k])
    return f"table_{raw_tables_entry.get('table_index', 0)}"


def _build_row_obj(raw_row: dict, defs: dict) -> dict:
    """Convert a flat field→value dict into a nested object, skipping lbl: fields."""
    obj: dict = {}
    for field, value in raw_row.items():
        fd = defs.get(field)
        if fd and fd.role == 'lbl':
            continue
        _set_nested(obj, _field_local(field), value)
    return obj


def _build_nested_output(raw: dict, defs: dict) -> dict:
    """
    Convert the internal raw result (cells/tables) into the public nested JSON:
      - lbl: fields are stripped
      - var: cell fields: dot-notation → nested dicts
      - var: table fields: per-instance {data, header, footer} objects in an array
    """
    out: dict = {}

    # ── scalar cells ──────────────────────────────────────────────────────────
    for field, value in raw.get('cells', {}).items():
        fd = defs.get(field)
        if fd and fd.role == 'lbl':
            continue
        _set_nested(out, field, value)

    # ── table groups ──────────────────────────────────────────────────────────
    for match in raw.get('tables', []):
        group = _table_group(match, defs)
        instance: dict = {}

        if '_source' in match:
            instance['_source'] = match['_source']

        if match.get('headers'):
            header_obj: dict = {}
            for raw_row in match['headers']:
                header_obj.update(_build_row_obj(raw_row, defs))
            if header_obj:
                instance['header'] = header_obj

        if match.get('data'):
            data_rows = [_build_row_obj(r, defs) for r in match['data']]
            data_rows = [r for r in data_rows if r]  # drop fully-empty rows
            if data_rows:
                instance['data'] = data_rows

        if match.get('footers'):
            footer_obj: dict = {}
            for raw_row in match['footers']:
                footer_obj.update(_build_row_obj(raw_row, defs))
            if footer_obj:
                instance['footer'] = footer_obj

        out.setdefault(group, []).append(instance)

    return out


# ── Sheet preparation helpers ──────────────────────────────────────────────────

def _expand_merged_cells(ws) -> dict:
    """
    Unmerge all merged ranges, fill every non-anchor cell with the anchor value,
    and return a merge-map for the scanner.

    **Fill rationale**: Excel displays merged ranges as a single visual cell.
    Repeated reads of any position in the range (e.g., DATA rows in a vertical
    merge used as a category column) must all return the same value, which
    requires filling.

    **Merge-map rationale**: the merge-map maps every (row, col) to the *full
    set* of (row, col) cells in that range.  When the scanner consumes any cell
    in a range it consumes all of them simultaneously, so the entire merged area
    is treated as one logical cell for sequential scanning.  Without this, a
    horizontal merge after a ``seek:`` + ``dir:`` change would expose duplicated
    values to the scan; with it, consuming the anchor cell silently retires
    every shadow cell too.

    The original file on disk is never modified — this operates on the
    in-memory workbook object only.
    """
    merge_map: dict[tuple, set] = {}

    # Snapshot everything before any structural changes, because unmerging
    # modifies ws.merged_cells in place.
    snapshots = []
    for merged_range in list(ws.merged_cells.ranges):
        cells_in_range: set[tuple] = set()
        for row_num in range(merged_range.min_row, merged_range.max_row + 1):
            for col_num in range(merged_range.min_col, merged_range.max_col + 1):
                cells_in_range.add((row_num, col_num))
        anchor_value = ws.cell(merged_range.min_row, merged_range.min_col).value
        snapshots.append((
            merged_range.min_row, merged_range.min_col,
            merged_range.max_row, merged_range.max_col,
            cells_in_range,
            anchor_value,
        ))

    for min_row, min_col, max_row, max_col, cells_in_range, anchor_value in snapshots:
        ws.unmerge_cells(
            start_row=min_row, start_column=min_col,
            end_row=max_row, end_column=max_col,
        )
        # Fill all cells (including the anchor) with the anchor value.
        for row_num in range(min_row, max_row + 1):
            for col_num in range(min_col, max_col + 1):
                ws.cell(row=row_num, column=col_num).value = anchor_value
        # Register the entire range in the merge-map (frozenset: immutable,
        # no aliasing risk if future code iterates and modifies).
        frozen = frozenset(cells_in_range)
        for pos in cells_in_range:
            merge_map[pos] = frozen

    return merge_map


def _warn_uncached_formulas(ws, logger: Logger) -> None:
    """
    Warn once if any cell still has data_type 'f' (formula) after loading
    with data_only=True, which means the cached value was never written.
    """
    for row in ws.iter_rows():
        for cell in row:
            if getattr(cell, 'data_type', None) == 'f':
                logger.warn_uncached_formulas()
                return


class SheetScanner:
    """
    Scans an openpyxl worksheet in a given direction.
    Tracks consumed cells. The main cursor only advances during cell:1 processing.
    Table scanning uses a separate local cursor so it never moves the main cursor.
    """

    def __init__(self, ws, config: Config, merge_map: dict | None = None):
        self.ws = ws
        self.config = config
        self.direction = config.read_direction
        self.consumed: set = set()
        # merge_map: (row, col) → frozenset of all (row, col) in the same merged range.
        # Consuming any cell in a range also consumes every other cell in that range,
        # so the entire merged area is treated as a single logical cell.
        self.merge_map: dict = merge_map or {}
        self.scan_order = self._build_scan_order()
        self.scan_order_index: dict[tuple, int] = {pos: i for i, pos in enumerate(self.scan_order)}
        self.cursor = 0
        # Set True by seek: instructions; controls whether the *next* dir: change
        # preserves the cursor position (seek-then-dir idiom) or resets it to 0.
        self._cursor_from_seek: bool = False

    def _build_scan_order(self) -> list:
        cells = []
        if self.direction == 'LR':
            for r in range(1, self.ws.max_row + 1):
                for c in range(1, self.ws.max_column + 1):
                    cells.append((r, c))
        elif self.direction == 'TD':
            for c in range(1, self.ws.max_column + 1):
                for r in range(1, self.ws.max_row + 1):
                    cells.append((r, c))
        return cells

    def set_direction(self, direction: str) -> None:
        """Switch the scan direction and rebuild the scan order.

        **Cursor behaviour** depends on what set the cursor most recently:

        * If a ``seek:`` instruction was the last cursor-moving action
          (``_cursor_from_seek`` is True), the cursor's current (row, col) is
          preserved — it is repositioned to the same cell in the new order.
          This is the ``seek:G4`` → ``dir:TD`` idiom: the direction change
          should start scanning *from* the seeked cell, not from the top-left.

        * Otherwise the cursor resets to 0 (beginning of the new scan order).
          Already-consumed cells are skipped naturally by ``advance_to_next``.
          This is the expected behaviour when a plain ``dir:`` mid-sequence
          switches direction without a preceding seek.

        The ``_cursor_from_seek`` flag is cleared unconditionally so that a
        second ``dir:`` without an intervening ``seek:`` always resets.
        """
        preserve = self._cursor_from_seek
        self._cursor_from_seek = False  # always clear after any direction change

        if preserve and self.cursor < len(self.scan_order):
            current_pos = self.scan_order[self.cursor]
        else:
            current_pos = None

        self.direction = direction
        self.scan_order = self._build_scan_order()
        self.scan_order_index = {pos: i for i, pos in enumerate(self.scan_order)}

        if current_pos is not None:
            # Reposition to the seeked cell in the new order (fall back to 0 if
            # somehow not found, which shouldn't happen for valid coordinates).
            self.cursor = self.scan_order_index.get(current_pos, 0)
        else:
            # No seek preceded this direction change — start from the beginning.
            self.cursor = 0

    def cell_value(self, row: int, col: int):
        return self.ws.cell(row=row, column=col).value

    def cell_empty(self, row: int, col: int) -> bool:
        return is_empty(self.cell_value(row, col), self.config.empty_aliases, self.config.ignore_case_values)

    def is_consumed(self, row: int, col: int) -> bool:
        return (row, col) in self.consumed

    def consume(self, row: int, col: int):
        """Mark (row, col) as consumed.  If it belongs to a merged range,
        all cells in that range are consumed at the same time — the entire
        merged region is one logical cell."""
        self.consumed.add((row, col))
        for pos in self.merge_map.get((row, col), ()):
            self.consumed.add(pos)

    def advance_to_next(self) -> tuple:
        """Advance main cursor to the next non-empty, non-consumed cell and return it."""
        while self.cursor < len(self.scan_order):
            r, c = self.scan_order[self.cursor]
            if not self.is_consumed(r, c) and not self.cell_empty(r, c):
                return r, c
            self.cursor += 1
        return None, None

    def advance_one(self):
        self.cursor += 1

    def find_next_from(self, search_cursor: int) -> tuple:
        """
        Find the next non-empty, non-consumed cell starting from search_cursor.
        Does NOT modify self.cursor. Returns (row, col, new_search_cursor).
        """
        i = search_cursor
        while i < len(self.scan_order):
            r, c = self.scan_order[i]
            if not self.is_consumed(r, c) and not self.cell_empty(r, c):
                return r, c, i
            i += 1
        return None, None, i


def extract_images(data_file: str, images_dir: str, stem: str) -> dict:
    """Extract embedded images from an xlsx file without requiring Pillow.

    xlsx files are ZIP archives.  Images live in ``xl/media/``; drawing XMLs in
    ``xl/drawings/drawing*.xml``; anchor relationships in
    ``xl/drawings/_rels/drawing*.xml.rels``.

    Returns a dict mapping cell reference (e.g. ``"B3"``) to the saved image
    path (relative to *images_dir*).  If no images are found, returns ``{}``.
    The *stem* is used to build image filenames:
    ``{stem}_{col}{row}_{idx}.{ext}``.
    """
    import zipfile
    import xml.etree.ElementTree as _ET
    import os

    _NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    _NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    result: dict = {}
    os.makedirs(images_dir, exist_ok=True)

    try:
        zf = zipfile.ZipFile(data_file, 'r')
    except (zipfile.BadZipFile, OSError):
        return result

    with zf:
        namelist = zf.namelist()

        # Find all drawing files
        drawing_names = [n for n in namelist if n.startswith('xl/drawings/drawing')
                         and n.endswith('.xml') and '/_rels/' not in n]

        idx = 0
        for drawing_path in sorted(drawing_names):
            # Read relationship file for this drawing
            rels_path = drawing_path.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels'
            if rels_path not in namelist:
                continue

            # Parse rels to map rId → media filename
            try:
                rels_tree = _ET.fromstring(zf.read(rels_path))
            except _ET.ParseError:
                continue
            rId_to_media: dict = {}
            for rel in rels_tree:
                rid = rel.get('Id', '')
                target = rel.get('Target', '')
                if '../media/' in target:
                    rId_to_media[rid] = target.replace('../media/', 'xl/media/')

            # Parse drawing XML to find anchors and their rIds
            try:
                draw_tree = _ET.fromstring(zf.read(drawing_path))
            except _ET.ParseError:
                continue

            for anchor in draw_tree:
                # <xdr:twoCellAnchor> or <xdr:oneCellAnchor>
                from_el = anchor.find(f'{{{_NS_XDR}}}from')
                if from_el is None:
                    continue
                col_el = from_el.find(f'{{{_NS_XDR}}}col')
                row_el = from_el.find(f'{{{_NS_XDR}}}row')
                if col_el is None or row_el is None:
                    continue
                try:
                    col_0 = int(col_el.text)
                    row_0 = int(row_el.text)
                except (ValueError, TypeError):
                    continue
                # Convert 0-based to 1-based cell ref
                from openpyxl.utils import get_column_letter as _gcl
                cell_ref_str = f'{_gcl(col_0 + 1)}{row_0 + 1}'

                # Find picture blipFill rId
                pic = anchor.find('.//' + f'{{{_NS_XDR}}}pic')
                if pic is None:
                    continue
                blip_fill = pic.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                if blip_fill is None:
                    continue
                r_embed = blip_fill.get(f'{{{_NS_R}}}embed')
                if not r_embed or r_embed not in rId_to_media:
                    continue

                media_path = rId_to_media[r_embed]
                if media_path not in namelist:
                    continue

                ext = os.path.splitext(media_path)[1] or '.bin'
                idx += 1
                out_name = f'{stem}_{cell_ref_str}_{idx}{ext}'
                out_path = os.path.join(images_dir, out_name)
                with open(out_path, 'wb') as fh:
                    fh.write(zf.read(media_path))
                result[cell_ref_str] = out_path

    return result


class Engine:
    def process(self, pattern_file: str, data_file: str,
                logger: Logger = None,
                max_file_mb: float = 5,
                max_uncompressed_mb: float = DEFAULT_MAX_UNCOMPRESSED_MB,
                max_cell_len: int = _MAX_REGEX_INPUT_LEN,
                max_rows: int = DEFAULT_MAX_DATA_ROWS,
                max_cols: int = DEFAULT_MAX_DATA_COLS,
                sheet: str | int | None = None,
                output_format: str = 'nested') -> dict:
        if logger is None:
            logger = Logger()

        self._max_cell_len = max_cell_len
        defs = {}
        _raw = {'cells': {}, 'tables': []}

        try:
            # Pattern file may be .xlsx or .csv; data file is .xlsx only.
            try:
                validate_pattern_file(pattern_file,
                                      max_file_mb=max_file_mb,
                                      max_uncompressed_mb=max_uncompressed_mb)
            except SecurityError as exc:
                logger.fatal(str(exc), found=pattern_file)
            try:
                validate_file(data_file,
                              max_file_mb=max_file_mb,
                              max_uncompressed_mb=max_uncompressed_mb)
            except SecurityError as exc:
                logger.fatal(str(exc), found=data_file)

            _pp = PatternParser()
            try:
                global_config, defs, start_sequence = _pp.parse(pattern_file)
            except (SecurityError, PatternError) as exc:
                logger.fatal(str(exc), found=pattern_file)
            assert_rules = getattr(_pp, 'assert_rules', [])

            if not start_sequence:
                logger.fatal(
                    'Pattern file defines no extraction steps — it has no START: '
                    'section, or the START: … END: block is empty.',
                    found=pattern_file,
                    expected='a START: … END: block with at least one cell: or table: instruction',
                )

            try:
                wb = openpyxl.load_workbook(data_file, data_only=True)
            except Exception as exc:
                logger.fatal(
                    f'Failed to open the data file: {exc}',
                    found=data_file,
                    expected='a valid, uncorrupted .xlsx workbook',
                )

            if sheet is None:
                ws = wb.active
            elif isinstance(sheet, int):
                if sheet < 0 or sheet >= len(wb.worksheets):
                    logger.fatal(
                        f'Sheet index {sheet} is out of range '
                        f'(workbook has {len(wb.worksheets)} sheet(s))',
                        found=str(sheet),
                        expected=f'an index between 0 and {len(wb.worksheets) - 1}',
                    )
                ws = wb.worksheets[sheet]
            elif sheet in wb.sheetnames:
                # Exact name match wins — including numeric names like "2025".
                ws = wb[sheet]
            elif str(sheet).lstrip('-').isdigit():
                # Numeric string with no matching name → treat as a 0-based index.
                idx = int(sheet)
                if idx < 0 or idx >= len(wb.worksheets):
                    logger.fatal(
                        f'Sheet index {idx} is out of range '
                        f'(workbook has {len(wb.worksheets)} sheet(s))',
                        found=str(sheet),
                        expected=f'an index between 0 and {len(wb.worksheets) - 1}',
                    )
                ws = wb.worksheets[idx]
            else:
                logger.fatal(
                    f'Sheet {sheet!r} not found in workbook',
                    found=sheet,
                    expected=f'one of: {", ".join(wb.sheetnames)}',
                )
                ws = wb.active  # unreachable (logger.fatal raises); keeps ws bound

            _check_sheet_dimensions(ws, max_rows, max_cols, logger)

            logger.engine_start(pattern_file, data_file)
            _raw = self._process_sheet(ws, global_config, defs, start_sequence, logger)

            # Run cross-field assert: rules against the extracted result.
            if assert_rules:
                _run_assert_rules(assert_rules, _raw['cells'], logger)

        except EngineError:
            pass  # setup-phase fatal; already logged, return partial result

        logger.summary(_raw)

        if output_format == 'legacy':
            warnings.warn(
                "--format legacy is deprecated and will be removed in a future version. "
                "Use --format nested (the default). The legacy format leaks internal "
                "lbl: keys and _source/_anchor metadata.",
                DeprecationWarning, stacklevel=3,
            )
            return _raw
        return _build_nested_output(_raw, defs)

    def process_all(self, pattern_file: str, data_file: str,
                    logger: Logger = None,
                    max_file_mb: float = 5,
                    max_uncompressed_mb: float = DEFAULT_MAX_UNCOMPRESSED_MB,
                    max_cell_len: int = _MAX_REGEX_INPUT_LEN,
                    max_rows: int = DEFAULT_MAX_DATA_ROWS,
                    max_cols: int = DEFAULT_MAX_DATA_COLS,
                    output_format: str = 'nested') -> dict:
        """
        Process every sheet in data_file using the same pattern.
        Returns a dict keyed by sheet name: {sheet_name: result, ...}.
        """
        if logger is None:
            logger = Logger()

        self._max_cell_len = max_cell_len
        defs = {}
        out: dict = {}

        try:
            # Pattern file may be .xlsx or .csv; data file is .xlsx only.
            try:
                validate_pattern_file(pattern_file,
                                      max_file_mb=max_file_mb,
                                      max_uncompressed_mb=max_uncompressed_mb)
            except SecurityError as exc:
                logger.fatal(str(exc), found=pattern_file)
            try:
                validate_file(data_file,
                              max_file_mb=max_file_mb,
                              max_uncompressed_mb=max_uncompressed_mb)
            except SecurityError as exc:
                logger.fatal(str(exc), found=data_file)

            _pp = PatternParser()
            try:
                global_config, defs, start_sequence = _pp.parse(pattern_file)
            except (SecurityError, PatternError) as exc:
                logger.fatal(str(exc), found=pattern_file)
            assert_rules = getattr(_pp, 'assert_rules', [])

            if not start_sequence:
                logger.fatal(
                    'Pattern file defines no extraction steps — it has no START: '
                    'section, or the START: … END: block is empty.',
                    found=pattern_file,
                    expected='a START: … END: block with at least one cell: or table: instruction',
                )

            try:
                wb = openpyxl.load_workbook(data_file, data_only=True)
            except Exception as exc:
                logger.fatal(
                    f'Failed to open the data file: {exc}',
                    found=data_file,
                    expected='a valid, uncorrupted .xlsx workbook',
                )

            logger.engine_start(pattern_file, data_file)

            for ws in wb.worksheets:
                _raw = {'cells': {}, 'tables': []}
                try:
                    _check_sheet_dimensions(ws, max_rows, max_cols, logger)
                    _raw = self._process_sheet(ws, global_config, defs, start_sequence, logger)
                    if assert_rules:
                        _run_assert_rules(assert_rules, _raw['cells'], logger)
                except EngineError:
                    pass  # per-sheet fatal; log and continue
                logger.summary(_raw)
                if output_format == 'legacy':
                    out[ws.title] = _raw
                else:
                    out[ws.title] = _build_nested_output(_raw, defs)

        except EngineError:
            pass  # setup-phase fatal

        return out

    def _process_sheet(self, ws, global_config, defs: dict,
                       start_sequence: list, logger: Logger) -> dict:
        """Run extraction on a single worksheet. Returns raw flat result dict."""
        _raw = {'cells': {}, 'tables': []}
        logger.begin_summary_scope()  # scope summary/ISSUES to THIS sheet
        logger.sheet_name = ws.title
        merge_map = _expand_merged_cells(ws)
        _warn_uncached_formulas(ws, logger)
        logger.sheet_info(ws.title, ws.max_row, ws.max_column, global_config.read_direction)
        logger.config_verbose(global_config)

        scanner = SheetScanner(ws, global_config, merge_map)
        table_index = 0

        try:
            for instruction in start_sequence:
                if isinstance(instruction, CellInstruction):
                    self._process_cell(instruction, scanner, defs, global_config, _raw, logger)
                elif isinstance(instruction, TableInstruction):
                    self._process_table(instruction, scanner, defs, _raw, table_index, logger)
                    table_index += 1
                elif isinstance(instruction, SeekInstruction):
                    self._process_seek(instruction, scanner, logger)
                elif isinstance(instruction, DirectionInstruction):
                    scanner.set_direction(instruction.direction)
                    logger.direction_changed(instruction.direction)
        except EngineError:
            pass  # already logged; return partial result

        return _raw

    # -------------------------------------------------------------------------
    # cell:1 processing
    # -------------------------------------------------------------------------

    def _resolve_abs_ref(self, instr: CellInstruction, scanner: SheetScanner,
                         defs: dict, config: Config, logger: Logger) -> tuple:
        """
        Resolve an absolute cell reference (e.g. 'B5') for cell:B5 instructions.
        Validates the cursor hasn't already passed the target.
        Returns (row, col) on success; calls logger.fatal (raises EngineError) on failure.
        Empty-cell handling is role-aware: lbl: → fatal, var: → allow None.
        """
        row, col = coordinate_to_tuple(instr.target)

        # Explicit cell addresses (cell:B4) are intentional jumps — they may land
        # before OR after the current cursor, e.g. after a dir: change that
        # preserved the cursor position from a preceding seek:.  No fatal check here;
        # the cursor is repositioned to idx+1 after the read (see _process_cell).

        # Warn when the target was already consumed: subsequent sequential cell:
        # instructions will scan from the wrong position.
        if (row, col) in scanner.consumed:
            logger.warn_abs_backward_ref(
                instr.target, cell_ref(row, col, logger.sheet_name)
            )

        value = scanner.cell_value(row, col)
        if is_empty(value, config.empty_aliases, config.ignore_case_values) and instr.field != 'IGNORE':
            fd = defs.get(instr.field)
            fd_role = fd.role if fd else 'var'
            if fd_role == 'lbl':
                logger.fatal(
                    f"Expected label '{instr.field}' at {instr.target} but cell is empty",
                    location=cell_ref(row, col, logger.sheet_name),
                    expected=(f"string matching {fd.regex!r}" if fd else 'a label string'),
                    found='empty cell',
                )

        return row, col

    def _process_cell(self, instr: CellInstruction, scanner: SheetScanner,
                      defs: dict, config: Config, result: dict, logger: Logger):
        if instr.multiplicity == 'abs':
            row, col = self._resolve_abs_ref(instr, scanner, defs, config, logger)
            value = scanner.cell_value(row, col)
            scanner.consume(row, col)
            idx = scanner.scan_order_index.get((row, col))
            if idx is not None:
                scanner.cursor = idx + 1
            # An abs cell: is a definitive cursor position; clear the seek flag so
            # a subsequent dir: resets to scan-order start rather than preserving
            # the stale seek-derived position.
            scanner._cursor_from_seek = False
        else:
            row, col = scanner.advance_to_next()
            if row is None:
                # Give an extra whitespace hint when a lbl: anchor can't be found —
                # invisible leading/trailing spaces in the source are a common cause.
                fd_check = defs.get(instr.field)
                ws_tip = (
                    " Tip: if the label cell has invisible leading/trailing whitespace "
                    "in the source file, add 'lbl:trim-whitespace' to this field."
                    if fd_check and fd_check.role == 'lbl' else ''
                )
                logger.fatal(
                    f'Expected cell:{instr.multiplicity} ({instr.field!r}) but sheet is exhausted',
                    expected=f'a cell containing field {instr.field!r}',
                    found=f'no more non-empty cells on the sheet.{ws_tip}',
                )
            value = scanner.cell_value(row, col)
            scanner.consume(row, col)
            scanner.advance_one()

        if instr.field == 'IGNORE':
            logger.cell_ignored(row, col, value)
            return

        fd = defs.get(instr.field)
        if fd is None:
            logger.fatal(
                f'Field {instr.field!r} referenced in START: but not found in def: section',
                location=cell_ref(row, col, logger.sheet_name),
                expected=f'a def: entry named {instr.field!r}',
                found='no matching def: row in pattern file',
            )

        # Apply trim-whitespace before required check, validation, logging, and storage.
        value = _apply_trim(value, fd, config)

        # Required (not-null/not-empty) check — fatal before any other validation.
        if fd.required and is_empty(value, config.empty_aliases, config.ignore_case_values):
            logger.fatal(
                f"Required field {fd.name!r} has an empty/null value",
                location=cell_ref(row, col, logger.sheet_name),
                expected=f'a non-empty value for {fd.name!r} (type: {fd.type})',
                found='empty cell',
            )

        # nullable: empty (or empty-after-trim) is silently accepted — normalise to
        # None so that validation is skipped and no warning is generated.
        if fd.nullable and is_empty(value, config.empty_aliases, config.ignore_case_values):
            value = None

        # Validate before tracing so the -v trace can show 🟢/🔴 per field.
        ok = None
        if value is not None:
            ok = _validate_field(fd, value, config, self._max_cell_len)

        logger.cell_processed(row, col, instr.field, value, ok=ok, regex=fd.regex)

        if ok is False:
            rec = logger.warn_validation(row, col, instr.field, fd.type, fd.regex, value)
            logger.commit_warnings([rec])

        result['cells'][instr.field] = value

    # -------------------------------------------------------------------------
    # seek: processing
    # -------------------------------------------------------------------------

    def _process_seek(self, instr: SeekInstruction, scanner: SheetScanner, logger: Logger):
        """Reposition the scanner cursor to the given cell without reading it."""
        row, col = coordinate_to_tuple(instr.target)
        idx = scanner.scan_order_index.get((row, col))
        if idx is None:
            logger.fatal(
                f"seek:{instr.target} targets a cell outside the sheet's used range",
                location=cell_ref(row, col, logger.sheet_name),
                expected='a valid cell within the used range of the sheet',
                found=f'cell {instr.target} is not in the scan order '
                      f'(sheet used range: {scanner.ws.max_row} rows × {scanner.ws.max_column} cols)',
            )
        scanner.cursor = idx
        scanner._cursor_from_seek = True

    # -------------------------------------------------------------------------
    # table:* processing
    # -------------------------------------------------------------------------

    def _process_table(self, instr: TableInstruction, scanner: SheetScanner,
                       defs: dict, result: dict, table_index: int, logger: Logger):
        """
        Greedy search for all mini-table instances.
        Uses a local search cursor — never modifies scanner.cursor.
        """
        logger.table_group_start(table_index, instr.config.read_direction)
        search_cursor = scanner.cursor
        instance_index = 0

        while True:
            match, search_cursor = self._try_match_mini_table(
                instr, scanner, defs, search_cursor, logger
            )
            if match is None:
                break

            anchor_row, anchor_col = match['anchor']
            end_row = match.pop('_end_row', anchor_row)
            end_col = match.pop('_end_col', anchor_col)

            match['_source'] = {
                'sheet': logger.sheet_name,
                'ref': _range_ref(anchor_row, anchor_col, end_row, end_col),
            }
            match['table_index'] = table_index
            match['instance_index'] = instance_index
            result['tables'].append(match)

            logger.mini_table_matched(
                table_index, instance_index,
                anchor_row, anchor_col, end_row, end_col,
            )
            instance_index += 1

        logger.table_group_done(table_index, instance_index)

    def _try_match_mini_table(self, instr: TableInstruction, scanner: SheetScanner,
                               defs: dict, search_cursor: int,
                               logger: Logger) -> tuple:
        """
        Search from search_cursor for the next valid mini-table.
        Returns (match_dict, new_search_cursor) on success or (None, cursor) when exhausted.
        Warnings are only committed for successful matches — failed attempts are silent.
        """
        config = instr.config

        while True:
            anchor_row, anchor_col, search_cursor = scanner.find_next_from(search_cursor)
            if anchor_row is None:
                return None, search_cursor

            logger.anchor_probe(anchor_row, anchor_col, scanner.cell_value(anchor_row, anchor_col))

            local_warnings: list[LogRecord] = []
            local_traces: list[str] = []
            match = self._attempt_match(
                anchor_row, anchor_col, instr, scanner, defs, config,
                local_warnings, local_traces, logger,
            )

            if match is not None:
                logger.commit_warnings(local_warnings)
                logger.commit_traces(local_traces)
                return match, search_cursor + 1

            logger.anchor_rejected(anchor_row, anchor_col, 'mini-table pattern did not match')
            search_cursor += 1

    def _attempt_match(self, anchor_row: int, anchor_col: int,
                       instr: TableInstruction, scanner: SheetScanner,
                       defs: dict, config: Config,
                       local_warnings: list, local_traces: list,
                       logger: Logger) -> dict | None:
        """
        Tentatively match the full mini-table at (anchor_row, anchor_col).
        On success: consumes all cells, returns extracted data.
        On failure: returns None, no cells consumed.
        Warnings go into local_warnings (caller decides whether to keep them).
        """
        tentative_consumed: set = set()
        extracted = {
            'headers': [], 'data': [], 'footers': [],
            'anchor': (anchor_row, anchor_col),
        }

        header_rows  = [r for r in instr.rows if r.row_type == 'HEADER']
        splitter_rows = [r for r in instr.rows if r.row_type == 'SPLITTER']
        data_rows    = [r for r in instr.rows if r.row_type == 'DATA']
        footer_rows  = [r for r in instr.rows if r.row_type == 'FOOTER']

        current_row = anchor_row

        # --- HEADER rows (strict: empty non-EMPTY field = fail) ---
        for tmpl_row in header_rows:
            row_data, ok = self._match_row(
                current_row, anchor_col, tmpl_row, defs, config,
                scanner, tentative_consumed, local_warnings, local_traces,
                logger, strict=True,
            )
            if not ok:
                return None
            if row_data:
                extracted['headers'].append(row_data)
            current_row += 1

        # --- SPLITTER rows (all columns must be empty) ---
        # Use the widest row type so SPLITTER validation covers the full table
        # width and _end_col is correct even when HEADER/SPLITTER have fewer
        # columns than DATA rows.
        num_cols = max(len(r.columns) for r in instr.rows) if instr.rows else 0
        for _ in splitter_rows:
            for c_offset in range(num_cols):
                col = anchor_col + c_offset
                val = scanner.ws.cell(row=current_row, column=col).value
                if not is_empty(val, config.empty_aliases, config.ignore_case_values):
                    return None
                tentative_consumed.add((current_row, col))
            current_row += 1

        # --- DATA rows ---
        data_tmpl = data_rows[0] if data_rows else None
        footer_tmpl_first = footer_rows[0] if footer_rows else None
        skip_if_rows = [r for r in instr.rows if r.row_type == 'SKIP_IF']
        _ser = [r for r in instr.rows if r.row_type == 'SKIP_EMPTY_ROW']
        max_skip_empty = sum(int(r.multiplicity) for r in _ser)

        if data_tmpl:
            is_bounded    = data_tmpl.max_rows is not None
            total_scanned = 0  # physical rows seen (skipped + real), for {n,m} bounds
            consecutive_empty = 0

            while True:
                # Hard ceiling for bounded DATA
                if is_bounded and total_scanned >= data_tmpl.max_rows:
                    # Drain any SKIP_EMPTY_ROW separator rows so footer matching
                    # starts on the actual footer row, not the separator.
                    while (consecutive_empty < max_skip_empty and
                           self._row_is_end_of_data(
                               current_row, anchor_col, data_tmpl, config, scanner)):
                        consecutive_empty += 1
                        current_row += 1
                    break

                if self._row_is_end_of_data(current_row, anchor_col, data_tmpl, config, scanner):
                    if consecutive_empty < max_skip_empty:
                        # SKIP_EMPTY_ROW budget — cross this empty separator row
                        consecutive_empty += 1
                        current_row += 1
                        continue
                    logger.footer_detected(current_row, anchor_col,
                                           scanner.cell_value(current_row, anchor_col))
                    current_row += 1
                    break

                consecutive_empty = 0

                if footer_tmpl_first and self._row_matches_footer(
                    current_row, anchor_col, footer_tmpl_first, defs, config, scanner
                ):
                    logger.footer_detected(current_row, anchor_col,
                                           scanner.cell_value(current_row, anchor_col))
                    break

                # SKIP_IF — silently skip matching rows (still counts toward bounds)
                if skip_if_rows and self._row_matches_any_skip_if(
                    current_row, anchor_col, skip_if_rows, config, scanner, defs
                ):
                    logger.data_row_skipped(current_row)
                    total_scanned += 1
                    current_row   += 1
                    if data_tmpl.multiplicity == '1':
                        break
                    continue

                row_data, ok = self._match_row(
                    current_row, anchor_col, data_tmpl, defs, config,
                    scanner, tentative_consumed, local_warnings, local_traces,
                    logger, strict=False,
                )
                if not ok:
                    return None

                logger.data_row(current_row, len(data_tmpl.columns))
                extracted['data'].append(row_data)
                total_scanned += 1
                current_row   += 1

                if data_tmpl.multiplicity == '1':
                    break

            # Warn if fewer physical rows than the declared minimum were found
            if is_bounded and total_scanned < data_tmpl.min_rows:
                local_warnings.append(
                    logger.warn_data_min_not_reached(data_tmpl.min_rows, total_scanned)
                )

        # --- FOOTER rows (strict: empty non-EMPTY field = fail) ---
        for tmpl_row in footer_rows:
            row_data, ok = self._match_row(
                current_row, anchor_col, tmpl_row, defs, config,
                scanner, tentative_consumed, local_warnings, local_traces,
                logger, strict=True,
            )
            if not ok:
                return None
            if row_data:
                extracted['footers'].append(row_data)
            current_row += 1

        # Full match — commit consumed cells and record span for caller
        scanner.consumed |= tentative_consumed
        extracted['_end_row'] = current_row - 1
        extracted['_end_col'] = anchor_col + num_cols - 1
        return extracted

    # -------------------------------------------------------------------------
    # Row-level matching helpers
    # -------------------------------------------------------------------------

    def _match_row(self, sheet_row: int, anchor_col: int, tmpl_row: TemplateRow,
                   defs: dict, config: Config, scanner: SheetScanner,
                   tentative_consumed: set, local_warnings: list, local_traces: list,
                   logger: Logger, strict: bool = False) -> tuple:
        """
        Match a single template row against a sheet row.
        strict=True: empty value in a non-EMPTY field causes immediate failure (HEADER/FOOTER).
        strict=False: empty value is warned but allowed (DATA).
        Returns (row_data: dict, success: bool).

        For DATA rows (strict=False) a per-field extraction trace is appended to
        local_traces; the caller commits it only if the whole mini-table matches.
        """
        row_data = {}

        for c_offset, tmpl_col in enumerate(tmpl_row.columns):
            col = anchor_col + c_offset
            val = scanner.ws.cell(row=sheet_row, column=col).value

            if tmpl_col.field == 'EMPTY':
                if not is_empty(val, config.empty_aliases, config.ignore_case_values):
                    return {}, False
                tentative_consumed.add((sheet_row, col))
                continue

            if tmpl_col.field == 'IGNORE':
                tentative_consumed.add((sheet_row, col))
                continue

            fd = defs.get(tmpl_col.field)

            if is_empty(val, config.empty_aliases, config.ignore_case_values):
                # required (not-null) check — fatal regardless of strict mode
                if fd is not None and fd.required:
                    logger.fatal(
                        f"Required field {fd.name!r} has an empty/null value",
                        location=cell_ref(sheet_row, col, logger.sheet_name),
                        expected=f'a non-empty value for {fd.name!r} (type: {fd.type})',
                        found='empty cell',
                    )
                if strict:
                    return {}, False  # HEADER/FOOTER: missing field = no match
                # nullable: accept null silently — no warning, no emoji on the trace line
                if fd is None or not fd.nullable:
                    fd_type = fd.type if fd else 'unknown'
                    local_warnings.append(
                        logger.warn_empty_field(sheet_row, col, tmpl_col.field, fd_type)
                    )
                row_data[tmpl_col.field] = None
                local_traces.append(
                    logger.trace_field(sheet_row, col, tmpl_col.field, None, ok=None)
                )
            else:
                trace_ok: bool | None = True
                trace_regex = ''
                if fd is None:
                    local_warnings.append(
                        logger.warn_undefined_field(sheet_row, col, tmpl_col.field)
                    )
                    trace_ok = False
                else:
                    # Apply trim-whitespace before validation and storage.
                    val = _apply_trim(val, fd, config)
                    ok = _validate_field(fd, val, config, self._max_cell_len)
                    if not ok:
                        if strict:
                            return {}, False  # HEADER/FOOTER: wrong value = no match
                        local_warnings.append(
                            logger.warn_validation(sheet_row, col, tmpl_col.field,
                                                   fd.type, fd.regex, val)
                        )
                    trace_ok, trace_regex = ok, fd.regex
                row_data[tmpl_col.field] = val
                if not strict:
                    local_traces.append(
                        logger.trace_field(sheet_row, col, tmpl_col.field, val,
                                           ok=trace_ok, regex=trace_regex)
                    )

            tentative_consumed.add((sheet_row, col))

        return row_data, True

    def _row_matches_any_skip_if(self, sheet_row: int, anchor_col: int,
                                  skip_if_rows: list, config: Config,
                                  scanner, defs: dict | None = None) -> bool:
        """Return True if the sheet row matches ANY SKIP_IF template (OR logic)."""
        return any(
            self._row_matches_skip_if(sheet_row, anchor_col, tmpl, config, scanner, defs or {})
            for tmpl in skip_if_rows
        )

    def _row_matches_skip_if(self, sheet_row: int, anchor_col: int,
                              skip_tmpl: TemplateRow, config: Config,
                              scanner, defs: dict | None = None) -> bool:
        """Return True if every non-IGNORE column in skip_tmpl matches its condition.

        Conditions (AND logic — ALL non-IGNORE columns must match):
          IGNORE → always match (skip this column check).
          EMPTY  → cell must be empty / null / in config.empty_aliases.
          <name> → cell must match the label field named ``<name>`` in defs.
                   An empty cell does NOT match a label condition (label must
                   find a value).  If ``<name>`` is unknown in defs the column
                   check is skipped (treated as IGNORE) so the SKIP_IF degrades
                   gracefully when defs are incomplete.
        """
        defs = defs or {}
        for c_offset, tmpl_col in enumerate(skip_tmpl.columns):
            if tmpl_col.field == 'IGNORE':
                continue
            col = anchor_col + c_offset
            val = scanner.ws.cell(row=sheet_row, column=col).value
            if tmpl_col.field == 'EMPTY':
                if not is_empty(val, config.empty_aliases, config.ignore_case_values):
                    return False
                continue
            # Label-based condition: the cell must match the named label field
            fd = defs.get(tmpl_col.field)
            if fd is None:
                continue   # unknown label → treat as IGNORE
            if is_empty(val, config.empty_aliases, config.ignore_case_values):
                return False  # empty cell never matches a label condition
            if not _validate_field(fd, _apply_trim(val, fd, config), config, self._max_cell_len):
                return False
        return True

    def _row_is_end_of_data(self, sheet_row: int, anchor_col: int,
                             data_tmpl: TemplateRow, config: Config,
                             scanner: SheetScanner) -> bool:
        for c_offset, tmpl_col in enumerate(data_tmpl.columns):
            if tmpl_col.field == 'IGNORE':
                continue
            col = anchor_col + c_offset
            val = scanner.ws.cell(row=sheet_row, column=col).value
            if not is_empty(val, config.empty_aliases, config.ignore_case_values):
                return False
        return True

    def _row_matches_footer(self, sheet_row: int, anchor_col: int,
                             footer_tmpl: TemplateRow, defs: dict,
                             config: Config, scanner: SheetScanner) -> bool:
        """Check if a row matches the first FOOTER template row (no cell consumption)."""
        for c_offset, tmpl_col in enumerate(footer_tmpl.columns):
            col = anchor_col + c_offset
            val = scanner.ws.cell(row=sheet_row, column=col).value

            if tmpl_col.field == 'EMPTY':
                if not is_empty(val, config.empty_aliases, config.ignore_case_values):
                    return False
                continue

            if tmpl_col.field == 'IGNORE':
                continue

            fd = defs.get(tmpl_col.field)
            if fd is None:
                continue
            if is_empty(val, config.empty_aliases, config.ignore_case_values):
                return False
            ok = _validate_field(fd, _apply_trim(val, fd, config), config, self._max_cell_len)
            if not ok:
                return False

        return True
