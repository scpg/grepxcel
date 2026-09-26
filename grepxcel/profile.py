"""`grepxcel profile` — census every cell's actual type in a data file.

No pattern required. Answers "what is actually in this file" before a
pattern is even written — storage type, semantic subtype, and the Layer-4
edge cases (errors, formulas, text-forced-numerics, blanks, merged-cell
members), using the canonical classifier in `cell_taxonomy.py`.

Same job description as `lint` (a raw-file inspection tool, not part of
the extract/pattern pipeline): scans every sheet by default, accepts
multiple files/directories, and uses the same 0/1 exit-code convention
(1 if any error cell was found, 0 otherwise) — not `extract --strict`'s
0/2 convention.
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import PatternFill

from .cell_taxonomy import CellProfile, classify_value
from .color import MARK_FAIL, MARK_INFO, MARK_OK, MARK_WARN, colorize_marks, should_color
from .lint import _expand_paths  # same directory/recursion convention as lint
from .security import DEFAULT_MAX_UNCOMPRESSED_MB, SecurityError, validate_file

DEFAULT_MAX_ROWS = 2048
DEFAULT_MAX_COLS = 1024

# Layer-4 "issue" flags — the broader bucket for --issues-only, beyond
# Excel's own error concept.
_ISSUE_FLAGS = frozenset({'error', 'text_forced_numeric'})


@dataclass
class ProfileCell:
    ref: str
    sheet: str
    profile: CellProfile


@dataclass
class SheetProfile:
    sheet: str
    cells: list[ProfileCell] = field(default_factory=list)
    truncated: bool = False


def _merged_member_coords(ws) -> set[str]:
    """Every coordinate covered by a merged range except its top-left cell —
    those read back as None in openpyxl and must not be miscounted as
    independent blank cells."""
    members: set[str] = set()
    for merged_range in ws.merged_cells.ranges:
        cells = list(merged_range.cells)
        for row, col in cells[1:]:
            members.add(openpyxl.utils.get_column_letter(col) + str(row))
    return members


def profile_sheet(ws, ws_raw, max_rows: int = DEFAULT_MAX_ROWS,
                  max_cols: int = DEFAULT_MAX_COLS) -> SheetProfile:
    """Classify every non-empty cell in one sheet.

    *ws* is loaded with data_only=True (cached values); *ws_raw* with
    data_only=False (preserves formula text) — same dual-load convention
    as lint._check_sheet, needed because a formula cell's data_only value
    is its cached result, which alone can't tell us it was a formula.
    """
    declared_rows = ws.max_row or 0
    declared_cols = ws.max_column or 0
    truncated = declared_rows > max_rows or declared_cols > max_cols
    scan_rows = min(declared_rows, max_rows)
    scan_cols = min(declared_cols, max_cols)

    merged_members = _merged_member_coords(ws)
    result = SheetProfile(sheet=ws.title, truncated=truncated)

    if scan_rows == 0 or scan_cols == 0:
        return result

    for row in ws_raw.iter_rows(min_row=1, max_row=scan_rows, max_col=scan_cols):
        for raw_cell in row:
            ref = raw_cell.coordinate
            is_formula = raw_cell.data_type == 'f'
            cached_cell = ws.cell(row=raw_cell.row, column=raw_cell.column)
            value = cached_cell.value

            if ref in merged_members and value is None:
                # Structural placeholder, not real data — skip it entirely
                # rather than counting it as a blank cell.
                continue
            if value is None and not is_formula:
                continue  # genuinely empty cell — not part of the census

            prof = classify_value(value, cached_cell.number_format, is_formula=is_formula)
            if ref in merged_members:
                prof.flags = prof.flags + ['merged_member']
            result.cells.append(ProfileCell(ref=ref, sheet=ws.title, profile=prof))

    return result


def profile_workbook(path: str, sheet_name: str | None = None,
                     max_rows: int = DEFAULT_MAX_ROWS,
                     max_cols: int = DEFAULT_MAX_COLS) -> list[SheetProfile]:
    """Profile every sheet in *path* (or just *sheet_name*, if given)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    wb_raw = openpyxl.load_workbook(path, data_only=False)
    try:
        names = wb.sheetnames
        if sheet_name is not None:
            resolved = sheet_name
            if resolved not in names:
                try:
                    resolved = names[int(sheet_name)]
                except (ValueError, IndexError):
                    raise SecurityError(f'Sheet {sheet_name!r} not found in {path}')
            names = [resolved]

        sheets = []
        for name in names:
            ws = wb[name]
            ws_raw = wb_raw[name]
            sheets.append(profile_sheet(ws, ws_raw, max_rows, max_cols))
        return sheets
    finally:
        wb.close()
        wb_raw.close()


# ── grouping / filtering ──────────────────────────────────────────────────────

def group_by_type(cells: list[ProfileCell]) -> dict[tuple[str, str], list[ProfileCell]]:
    """Group cells by (storage_type, semantic_type), preserving first-seen order."""
    groups: dict[tuple[str, str], list[ProfileCell]] = {}
    for c in cells:
        key = (c.profile.storage_type, c.profile.semantic_type)
        groups.setdefault(key, []).append(c)
    return groups


def filter_cells(cells: list[ProfileCell], mode: str | None) -> list[ProfileCell]:
    """mode: None (all), 'errors' (strict Excel errors), 'issues' (errors + suspicious)."""
    if mode is None:
        return cells
    if mode == 'errors':
        return [c for c in cells if 'error' in c.profile.flags]
    if mode == 'issues':
        return [c for c in cells if set(c.profile.flags) & _ISSUE_FLAGS]
    raise ValueError(f'unknown filter mode {mode!r}')


def has_error(cells: list[ProfileCell]) -> bool:
    return any('error' in c.profile.flags for c in cells)


# ── rendering ──────────────────────────────────────────────────────────────

def _sample_refs(cells: list[ProfileCell], limit: int = 3) -> str:
    refs = [c.ref for c in cells[:limit]]
    text = ', '.join(refs)
    if len(cells) > limit:
        text += f', … +{len(cells) - limit} more'
    return text


def render_summary(all_cells: list[ProfileCell], color: bool = False) -> str:
    lines = []
    groups = group_by_type(all_cells)
    for (storage, semantic), cells in groups.items():
        n = len(cells)
        noun = 'cell' if n == 1 else 'cells'
        lines.append(f'  {storage}/{semantic}  ({n} {noun}): {_sample_refs(cells)}')
    if not lines:
        return '  (no cells matched)'
    return '\n'.join(lines)


def render_full(all_cells: list[ProfileCell]) -> str:
    lines = []
    groups = group_by_type(all_cells)
    for (storage, semantic), cells in groups.items():
        n = len(cells)
        noun = 'cell' if n == 1 else 'cells'
        lines.append(f'{storage}/{semantic}  ({n} {noun})')
        for c in cells:
            flags = f'  flags: {",".join(c.profile.flags)}' if c.profile.flags else ''
            lines.append(f'  {c.ref:6} {c.profile.raw_repr!r:30} fmt: {c.profile.number_format}{flags}')
    if not lines:
        return '  (no cells matched)'
    return '\n'.join(lines)


def to_json_records(sheets: list[SheetProfile], mode: str | None) -> list[dict]:
    records = []
    for sp in sheets:
        for c in filter_cells(sp.cells, mode):
            records.append({
                'sheet': c.sheet,
                'ref': c.ref,
                'storage_type': c.profile.storage_type,
                'semantic_type': c.profile.semantic_type,
                'python_type': c.profile.python_type,
                'number_format': c.profile.number_format,
                'raw_value': c.profile.raw_repr,
                'flags': c.profile.flags,
            })
    return records


# ── xlsx colored report ──────────────────────────────────────────────────────

_PROFILE_FILLS = {
    'error':               PatternFill('solid', fgColor='FFC7CE'),  # red
    'text_forced_numeric': PatternFill('solid', fgColor='FFE699'),  # amber
    'formula':             PatternFill('solid', fgColor='D0E8FF'),  # blue
    'merged_member':       PatternFill('solid', fgColor='E0E0E0'),  # grey
}


def _fill_for(profile: CellProfile) -> PatternFill | None:
    # Priority order: error > text_forced_numeric > formula > merged_member.
    for flag in ('error', 'text_forced_numeric', 'formula', 'merged_member'):
        if flag in profile.flags:
            return _PROFILE_FILLS[flag]
    return None


def write_colored_xlsx(data_path: str, sheets: list[SheetProfile], out_path: str) -> str:
    """A colored *copy* of the actual data file — never mutates the original."""
    wb = openpyxl.load_workbook(data_path)
    for sp in sheets:
        ws = wb[sp.sheet]
        for c in sp.cells:
            fill = _fill_for(c.profile)
            if fill:
                ws[c.ref].fill = fill
    wb.save(out_path)
    return out_path


def write_json(sheets: list[SheetProfile], mode: str | None, out_path: str) -> str:
    records = to_json_records(sheets, mode)
    with open(out_path, 'w', encoding='utf-8') as fh:
        json.dump(records, fh, indent=2, ensure_ascii=False)
    return out_path


# ── CLI entry point ──────────────────────────────────────────────────────────

def run_profile(files: list[str], out=None, recursive: bool = False,
                sheet: str | None = None, verbose: bool = False, quiet: bool = False,
                errors_only: bool = False, issues_only: bool = False,
                fmt: str = 'json', output: str | None = None, force: bool = False,
                max_file_mb: float = 5, max_uncompressed_mb: float = DEFAULT_MAX_UNCOMPRESSED_MB,
                max_rows: int = DEFAULT_MAX_ROWS, max_cols: int = DEFAULT_MAX_COLS) -> int:
    """Profile one or more files or directories, print/write results, return exit code.

    Exit code matches `lint`'s convention: 1 if any error-flagged cell was
    found across the scanned files, 0 otherwise — independent of which
    display filter (--errors-only/--issues-only) was requested. The filter
    controls what's printed/written; the exit code reflects what's
    actually in the file.
    """
    out = out or sys.stdout
    color = should_color(out)
    mode = 'errors' if errors_only else ('issues' if issues_only else None)
    any_error = False

    expanded = _expand_paths(files, recursive)
    if not expanded:
        dirs = [p for p in files if os.path.isdir(p)]
        for d in dirs:
            flag = ' (recursive)' if recursive else ' (pass -r to recurse)'
            print(colorize_marks(f'  {MARK_INFO}  {d} — no Excel files found{flag}', color), file=out)
        return 0

    multi = len(expanded) > 1
    if multi and output and fmt != 'json':
        print(colorize_marks(
            f'  {MARK_FAIL}  Multiple input files: --format xlsx needs one output '
            f'file per input — pass a directory to -o, not a file path.', color), file=out)
        return 1
    out_dir = output if (multi and output) else None
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    for path in expanded:
        try:
            validate_file(path, max_file_mb, max_uncompressed_mb)
        except SecurityError as exc:
            print(colorize_marks(f'  {MARK_FAIL}  {path}   {exc}', color), file=out)
            any_error = True
            continue

        try:
            sheets = profile_workbook(path, sheet, max_rows, max_cols)
        except Exception as exc:
            print(colorize_marks(f'  {MARK_FAIL}  {path}   Could not read file: {exc}', color), file=out)
            any_error = True
            continue

        all_cells = [c for sp in sheets for c in sp.cells]
        if has_error(all_cells):
            any_error = True

        this_out = None
        if output:
            this_out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0]
                                    + ('.xlsx' if fmt == 'xlsx' else '.json')) if out_dir else output
            if os.path.exists(this_out) and not force:
                print(colorize_marks(
                    f'  {MARK_FAIL}  {this_out} already exists — use --force to overwrite.',
                    color), file=out)
                return 1
            if fmt == 'xlsx':
                write_colored_xlsx(path, sheets, this_out)
            else:
                write_json(sheets, mode, this_out)
            mark = MARK_OK
            print(colorize_marks(f'  {mark}  {path} → {this_out}', color), file=out)
            continue

        filtered = filter_cells(all_cells, mode)
        if quiet:
            n = len(filtered)
            mark = MARK_WARN if any('error' in c.profile.flags for c in filtered) else MARK_OK
            print(colorize_marks(f'  {mark}  {path}   {n} cell(s)', color), file=out)
        elif verbose:
            print(f'\ngrepxcel profile — {path}\n' + '─' * 62, file=out)
            print(render_full(filtered), file=out)
        else:
            print(f'\ngrepxcel profile — {path}\n' + '─' * 62, file=out)
            print(render_summary(filtered, color), file=out)

    return 1 if any_error else 0
