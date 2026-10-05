"""`grepxcel lint` — inspect an Excel file before extraction.

Reports, with a ✓ / ⚠ / ✗ / ℹ checklist, potential issues in a data file
that might cause extraction to fail or produce unexpected results. Checks:

* File access and format (extension, magic bytes, encryption/IRM)
* Sheet dimensions — declared vs real used extent
* Merged cells
* Formula cells (cached values may be stale)
* Empty sheets
* Multi-sheet inventory

Also includes advisory notes about known limitations not yet detected
(e.g. Microsoft Information Protection, password-protected sheets).
"""
from __future__ import annotations

import os
import sys
import warnings
import zipfile

from .color import MARK_FAIL, MARK_INFO, MARK_OK, MARK_WARN, colorize_marks, should_color
from .security import SecurityError, _check_zip_safety

OK, WARN, FAIL, INFO = 'ok', 'warn', 'fail', 'info'
_MARK = {OK: MARK_OK, WARN: MARK_WARN, FAIL: MARK_FAIL, INFO: MARK_INFO}

# OLE Compound Document magic (D0 CF 11 E0 A1 B1 1A E1)
_OLE_MAGIC = b'\xd0\xcf\x11\xe0'
_ZIP_MAGIC = b'PK\x03\x04'
# .xls is rejected for a different reason than .xlsm/.xlsb: it's not a
# security choice, it's a format-support gap — .xls is a legacy OLE2/BIFF8
# binary container, not the ZIP+XML structure everything else here reads,
# so grepxcel's openpyxl-based pipeline cannot open it at all, macros or
# not. It gets its own message (see lint_file()) rather than being folded
# into the macro-risk wording below, which would wrongly imply a clean
# .xls file could work if it just had no macros.
_MACRO_EXTENSIONS = {'.xlsm', '.xlsb'}

Result = tuple


def lint_file(path: str) -> list[Result]:
    """Run all lint checks on a single file. Returns a list of results."""
    results: list[Result] = []

    # ── 1. File existence and type ────────────────────────────────────────
    if not os.path.exists(path):
        results.append((FAIL, 'file access', f'File does not exist: {path}'))
        _add_advisory(results)
        return results

    if not os.path.isfile(path):
        results.append((FAIL, 'file access', f'Not a regular file: {path}'))
        _add_advisory(results)
        return results

    _, ext = os.path.splitext(path)
    ext = ext.lower()

    if ext == '.xls':
        results.append((FAIL, 'file format',
                         "'.xls' files are not accepted. The legacy Excel 97-2003 "
                         "binary format is not supported — grepxcel reads modern "
                         ".xlsx files only (a format-support limitation, not a "
                         "security block). In Excel: File > Save As > Excel "
                         "Workbook (.xlsx), then run grepxcel again on the "
                         "converted file."))
        _add_advisory(results)
        return results

    if ext in _MACRO_EXTENSIONS:
        results.append((FAIL, 'file format',
                         f'{ext!r} files are not accepted — macro-enabled and binary '
                         f'Excel formats are blocked. Save as .xlsx first.'))
        _add_advisory(results)
        return results

    if ext != '.xlsx':
        results.append((FAIL, 'file format',
                         f'Unsupported file type {ext!r}. Only .xlsx files are accepted.'))
        _add_advisory(results)
        return results

    # ── 2. Magic bytes — ZIP vs OLE vs corrupt ────────────────────────────
    try:
        with open(path, 'rb') as fh:
            header = fh.read(8)
    except OSError as exc:
        results.append((FAIL, 'file access', f'Cannot read file: {exc}'))
        _add_advisory(results)
        return results

    if header[:4] == _OLE_MAGIC:
        results.append((FAIL, 'file format',
                         'File is an OLE Compound Document, not a standard .xlsx (ZIP). '
                         'This usually means the file is encrypted, password-protected, '
                         'or restricted by Microsoft Information Protection (MIP/IRM). '
                         'To use it with grepxcel: (1) open the file in Excel, '
                         '(2) remove protection or save a decrypted copy as .xlsx, '
                         'then re-run lint.'))
        _add_advisory(results)
        return results

    if header[:4] != _ZIP_MAGIC:
        results.append((FAIL, 'file format',
                         f'File does not have a valid .xlsx (ZIP) signature. '
                         f'It may be corrupted or renamed from a different format.'))
        _add_advisory(results)
        return results

    # ── 3. ZIP integrity + ZIP bomb guard ────────────────────────────────
    try:
        _check_zip_safety(path, max_uncompressed_mb=50.0)
    except SecurityError as exc:
        results.append((FAIL, 'file integrity', str(exc)))
        _add_advisory(results)
        return results

    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
    except zipfile.BadZipFile:
        results.append((FAIL, 'file integrity',
                         'File has ZIP magic bytes but is not a valid ZIP archive. '
                         'It may be corrupted.'))
        _add_advisory(results)
        return results

    _check_macro_content(names, results)
    _check_invisible_data(names, results)

    file_mb = os.path.getsize(path) / (1024 * 1024)
    results.append((OK, 'file access', f'{path} — {file_mb:.1f} MB'))

    # ── 4. Open with openpyxl ─────────────────────────────────────────────
    try:
        import openpyxl
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            wb = openpyxl.load_workbook(path, data_only=True)
            wb_raw = openpyxl.load_workbook(path, data_only=False)
    except Exception as exc:
        results.append((FAIL, 'workbook load',
                         f'openpyxl cannot open this file: {exc}. '
                         f'The file may be corrupted, encrypted, or in an '
                         f'unsupported format variant.'))
        _add_advisory(results)
        return results

    # ── 5. Sheet inventory ────────────────────────────────────────────────
    sheet_names = wb.sheetnames
    n_sheets = len(sheet_names)
    if n_sheets == 1:
        results.append((OK, 'sheets', f'1 sheet: {sheet_names[0]!r}'))
    else:
        results.append((OK, 'sheets',
                         f'{n_sheets} sheets: {", ".join(repr(s) for s in sheet_names)}. '
                         f'Use --sheet to select (default: active sheet).'))

    # ── 6. Per-sheet checks ───────────────────────────────────────────────
    for ws, ws_raw in zip(wb.worksheets, wb_raw.worksheets):
        _check_sheet(ws, ws_raw, results)

    wb.close()
    wb_raw.close()

    # ── 7. Advisory notes ─────────────────────────────────────────────────
    _add_advisory(results)

    return results


def _check_sheet(ws, ws_raw, results: list[Result]) -> None:
    """Run per-sheet checks: dimensions, extent, merged cells, formulas.

    ws is loaded with data_only=True (cached values), ws_raw with
    data_only=False (preserves formula text).
    """
    title = ws.title
    declared_rows = ws.max_row or 0
    declared_cols = ws.max_column or 0

    # Empty sheet
    if declared_rows == 0 or declared_cols == 0:
        results.append((WARN, f'sheet {title!r} data',
                         'Sheet is empty — nothing to extract.'))
        return

    # Real used extent (from data_only view)
    used_rows = 0
    used_cols = 0
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=declared_rows,
                         max_col=declared_cols, values_only=True), start=1):
        for col_idx, val in enumerate(row, start=1):
            if val is not None:
                used_rows = row_idx
                if col_idx > used_cols:
                    used_cols = col_idx

    # Formula cells (from raw view — preserves formula text)
    formula_cells = []
    raw_rows = ws_raw.max_row or 0
    raw_cols = ws_raw.max_column or 0
    for row in ws_raw.iter_rows(min_row=1, max_row=raw_rows,
                                max_col=raw_cols):
        for cell in row:
            if cell.data_type == 'f':
                formula_cells.append(cell.coordinate)

    if used_rows == 0:
        results.append((WARN, f'sheet {title!r} data',
                         'Sheet has formatting but no data — empty for extraction.'))
        return

    # Dimension report
    if (declared_rows > used_rows * 2 or declared_cols > used_cols * 2) and (
            declared_rows > used_rows + 10 or declared_cols > used_cols + 10):
        results.append((WARN, f'sheet {title!r} dimensions',
                         f'Declared {declared_rows}×{declared_cols} but only '
                         f'{used_rows}×{used_cols} contains data — inflated by '
                         f'empty formatted cells. grepxcel uses real extent, '
                         f'so this is handled, but the file is larger than needed.'))
    else:
        results.append((OK, f'sheet {title!r} dimensions',
                         f'{used_rows} rows × {used_cols} columns'))

    # Merged cells
    merged = list(ws.merged_cells.ranges)
    if merged:
        examples = ', '.join(str(m) for m in merged[:5])
        suffix = f' (and {len(merged) - 5} more)' if len(merged) > 5 else ''
        results.append((WARN, f'sheet {title!r} merged cells',
                         f'{len(merged)} merged region(s): {examples}{suffix}. '
                         f'grepxcel reads the top-left cell value; other cells in '
                         f'the merge appear empty.'))
    else:
        results.append((OK, f'sheet {title!r} merged cells', 'none'))

    # Formula cells
    if formula_cells:
        examples = ', '.join(formula_cells[:5])
        suffix = f' (and {len(formula_cells) - 5} more)' if len(formula_cells) > 5 else ''
        results.append((WARN, f'sheet {title!r} formulas',
                         f'{len(formula_cells)} formula cell(s): {examples}{suffix}. '
                         f'grepxcel reads cached values (data_only=True). If the file '
                         f'was last saved by a tool that doesn\'t cache formula results, '
                         f'these cells may appear empty. Open and re-save in Excel to fix.'))
    else:
        results.append((OK, f'sheet {title!r} formulas', 'none'))


# Zip parts whose presence means part of the workbook's data is invisible to
# grepxcel — checked directly against the raw ZIP namelist (already opened
# for the integrity check below) rather than left as an unconditional,
# always-shown note. Each fires only when the file actually has that part.
_INVISIBLE_DATA_SIGNALS = (
    ('xl/model/', (INFO, 'advisory',
        'This workbook has a Data Model (Power Pivot). grepxcel reads worksheet '
        'cells only — data, measures, and calculated columns that live only in '
        'the Data Model are not read, even if a PivotTable on a worksheet '
        'displays their results.')),
    ('xl/connections.xml', (INFO, 'advisory',
        'This workbook has an external data connection (ODBC/OLEDB/Power Query). '
        'grepxcel reads the cached worksheet values only — it does not refresh '
        'or read the connection itself, so results reflect whatever was cached '
        'the last time the workbook was saved.')),
    ('xl/externalLinks/', (INFO, 'advisory',
        "This workbook references another workbook (e.g. a formula like "
        "'=[Book2.xlsx]Sheet1!A1'). grepxcel reads the cached result of that "
        'link only — it does not open or follow the referenced file.')),
    ('xl/embeddings/', (INFO, 'advisory',
        'This workbook has an embedded object (e.g. an inserted file or '
        'document). grepxcel only reads cell values — embedded objects are '
        'not extracted or inspected.')),
)


def _check_macro_content(names: list[str], results: list[Result]) -> None:
    """Report macro parts by content, so the reason is visible before extraction.

    `lint` deliberately still opens and reports on such a file rather than
    stopping: its whole job is telling you what a file contains, and a user whose
    extraction was refused needs to see *why* here. `extract` refuses it — see
    `security._check_no_executable_content`.

    The part list lives in `security` so the two cannot drift apart.
    """
    from .security import _EXECUTABLE_PARTS

    lowered = [(n, n.lower()) for n in names]
    for needle, description in _EXECUTABLE_PARTS:
        for original, low in lowered:
            hit = low.startswith(needle) if needle.endswith('/') else low == needle
            if not hit:
                continue
            results.append((FAIL, 'macro content',
                             f'{original} — {description}. Extraction will refuse '
                             f'this file whatever its extension says, because Excel '
                             f'stores macros only in .xlsm/.xlsb: a .xlsx holding '
                             f'this was renamed, not saved that way. grepxcel does '
                             f'not scan for malware and never runs macro content — '
                             f'this reports what is present, not that it is harmful. '
                             f'To extract, re-save as .xlsx in Excel (which drops '
                             f'macros) once you trust the file.'))
            break


def _check_invisible_data(names: list[str], results: list[Result]) -> None:
    """Flag zip parts that hold data grepxcel never reads, so a user isn't
    silently missing part of the picture with no indication of it."""
    for prefix, result in _INVISIBLE_DATA_SIGNALS:
        if any(n.startswith(prefix) for n in names):
            results.append(result)


def _add_advisory(results: list[Result]) -> None:
    """Append advisory notes about conditions not yet auto-detected."""
    results.append((INFO, 'advisory',
                     'Microsoft Information Protection (MIP/IRM): if your organisation '
                     'applies sensitivity labels that encrypt or restrict Excel files, '
                     'grepxcel cannot read them directly. Open the file in Excel and '
                     'save an unprotected copy, or ask your IT admin for a decrypted '
                     'export. Lint detects the OLE signature of encrypted files but '
                     'cannot distinguish IRM labels from other encryption.'))
    results.append((INFO, 'advisory',
                     'Password-protected workbooks or sheets: grepxcel does not prompt '
                     'for passwords. A workbook-level password produces an encrypted '
                     'OLE file (detected above). A sheet-level password still allows '
                     'reading cell values — extraction works, but the file author may '
                     'not intend the data to be extracted.'))
    results.append((INFO, 'advisory',
                     'Conditional formatting, data validation, charts, and VBA modules '
                     'are ignored by grepxcel. They do not affect extraction but may '
                     'indicate the file is more complex than a flat data sheet. '
                     '(A plain PivotTable built from a worksheet range is fully '
                     'readable — its source cells are ordinary data. A Data Model '
                     '/ connection / embedded object is flagged separately above, '
                     'only when the file actually has one.)'))


# ── directory expansion ───────────────────────────────────────────────────────

_EXCEL_EXTENSIONS = {'.xlsx', '.xlsm', '.xlsb', '.xls'}


def _expand_paths(paths: list[str], recursive: bool) -> list[str]:
    """Expand any directory entries in *paths* to the Excel files they contain.

    Non-directory entries are kept as-is (even if they don't exist — lint_file
    will report the missing-file error).  Directories are scanned for files
    with Excel extensions; with recursive=True, subdirectories are included.
    Files within each directory are returned in sorted order.
    """
    expanded: list[str] = []
    for p in paths:
        if not os.path.isdir(p):
            expanded.append(p)
            continue
        if recursive:
            for root, _dirs, files in os.walk(p):
                _dirs.sort()
                for f in sorted(files):
                    if os.path.splitext(f)[1].lower() in _EXCEL_EXTENSIONS:
                        expanded.append(os.path.join(root, f))
        else:
            for f in sorted(os.listdir(p)):
                if os.path.splitext(f)[1].lower() in _EXCEL_EXTENSIONS:
                    expanded.append(os.path.join(p, f))
    return expanded


# ── runner (CLI entry point) ──────────────────────────────────────────────────

def _compact_line(path: str, results: list[Result]) -> str:
    """One-line summary for a single file: mark + path [+ issue categories]."""
    fails = [r for r in results if r[0] == FAIL]
    warns = [r for r in results if r[0] == WARN]
    if fails:
        issues = ' · '.join(dict.fromkeys(r[1] for r in fails))
        return f'  {_MARK[FAIL]}  {path}   {issues}'
    if warns:
        issues = ' · '.join(dict.fromkeys(r[1] for r in warns))
        return f'  {_MARK[WARN]}  {path}   {issues}'
    return f'  {_MARK[OK]}  {path}'


def run_lint(files: list[str], out=None, recursive: bool = False,
             verbose: bool = False) -> int:
    """Lint one or more files or directories, print results, return exit code.

    Default (verbose=False): one summary line per file.
    With verbose=True: full checklist detail (all checks printed).
    Output goes to stdout so piping works correctly with |less, |wc, etc.
    """
    out = out or sys.stdout
    color = should_color(out)
    any_fail = False

    expanded = _expand_paths(files, recursive)

    if not expanded:
        dirs = [p for p in files if os.path.isdir(p)]
        if dirs:
            for d in dirs:
                flag = ' (recursive)' if recursive else ' (pass -r to recurse)'
                print(colorize_marks(
                    f'  {_MARK[INFO]}  {d} — no Excel files found{flag}',
                    color), file=out)
        return 0

    for path in expanded:
        results = lint_file(path)
        if any(r[0] == FAIL for r in results):
            any_fail = True

        if verbose:
            print(f'\ngrepxcel lint — {path}\n' + '─' * 62, file=out)
            for status, name, detail in results:
                mark = _MARK[status]
                print(colorize_marks(
                    f'  {mark}  {name:<32} {detail}', color), file=out)
            print('─' * 62, file=out)
        else:
            print(colorize_marks(_compact_line(path, results), color), file=out)

    return 1 if any_fail else 0
