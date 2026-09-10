"""`grepxcel validate-pattern` — statically validate a pattern file.

Runs the same PatternParser used by extraction (so every structural / security /
regex / type / multiplicity / comment rule is enforced), then adds the two checks
the parser alone doesn't make — an empty extraction sequence and references to
undefined fields — so a "valid" verdict really means the pattern is usable.

Works for .xlsx and .csv patterns alike (the parser reads both into one grid).
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field

from .color import MARK_FAIL, MARK_OK, MARK_WARN, colorize_marks, paint, should_color
from .models import CellInstruction, TableInstruction, SeekInstruction, DirectionInstruction
from .pattern_parser import PatternError, PatternParser
from .security import SecurityError

# Patterns that strongly suggest regex intent (backslash-escapes, lookahead)
_REGEX_TELL = re.compile(r'\\[()[\]{}|+*.?^$]|[(][?]')

_MARK_OK, _MARK_WARN, _MARK_FAIL = MARK_OK, MARK_WARN, MARK_FAIL


@dataclass
class CheckResult:
    path: str
    valid: bool = False
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    config: object = None
    defs: dict | None = None
    sequence: list | None = None

    @property
    def n_fields(self) -> int:
        return len(self.defs) if self.defs else 0

    @property
    def n_steps(self) -> int:
        return len(self.sequence) if self.sequence else 0


def _referenced_fields(sequence) -> set[str]:
    """Field names referenced by the extraction sequence (excludes IGNORE/EMPTY)."""
    refs: set[str] = set()
    for instr in sequence:
        if isinstance(instr, CellInstruction):
            if instr.field not in ('IGNORE', 'EMPTY'):
                refs.add(instr.field)
        elif isinstance(instr, TableInstruction):
            for trow in instr.rows:
                for col in trow.columns:
                    if col.field not in ('IGNORE', 'EMPTY'):
                        refs.add(col.field)
    return refs


def check_pattern(path: str) -> CheckResult:
    """Parse and statically validate a pattern file. Never raises — every problem
    is captured in result.errors (fatal) or result.warnings (non-fatal)."""
    result = CheckResult(path=path)
    try:
        config, defs, sequence = PatternParser().parse(path)
    except (PatternError, SecurityError) as exc:
        result.errors.append(str(exc))
        return result
    except FileNotFoundError:
        result.errors.append(f'File not found: {path}')
        return result
    except Exception as exc:  # pragma: no cover - defensive
        result.errors.append(f'{type(exc).__name__}: {exc}')
        return result

    result.config, result.defs, result.sequence = config, defs, sequence

    if not sequence:
        result.errors.append(
            'No extraction steps — the pattern has no START: section, or the '
            'START: … END: block is empty.'
        )

    referenced = _referenced_fields(sequence)
    for name in sorted(referenced - set(defs)):
        result.errors.append(
            f"Field {name!r} is referenced in the extraction sequence but never "
            f"defined with a var:/lbl: row."
        )
    for name in sorted(set(defs) - referenced):
        result.warnings.append(f"Field {name!r} is defined but never used.")

    for key in config.unknown_config_keys:
        result.warnings.append(
            f"Unknown config key {key!r} — ignored. "
            f"Valid keys: pattern.version (or version), read.direction, "
            f"currency.sign, ignore.case, trim.whitespace, lbl.match, var.match, empty.aliases."
        )

    for name, fd in defs.items():
        if fd.role == 'lbl':
            effective_mode = fd.lbl_match if fd.lbl_match is not None else config.lbl_match
            if effective_mode != 'regexp' and _REGEX_TELL.search(fd.regex):
                plain = re.sub(r'\\(.)', r'\1', fd.regex)
                result.warnings.append(
                    f"lbl: field {name!r} pattern {fd.regex!r} looks like a regex "
                    f"but lbl.match mode is {effective_mode!r}. "
                    f"In literal/glob mode backslash-escapes are matched literally. "
                    f"Did you mean {plain!r}? "
                    f"Add lbl:regexp or set config: | lbl.match | regexp to use regex."
                )
        elif fd.role == 'var' and fd.var_mode in ('literal', 'glob'):
            # glob/literal without a pattern is a no-op — warn so the author notices.
            if fd.regex in ('', '.*'):
                result.warnings.append(
                    f"var: field {name!r} has mode {fd.var_mode!r} but column D is "
                    f"empty (pattern '.*'). Add a {fd.var_mode} pattern in column D "
                    f"or change to plain var: for type-only validation."
                )

    result.valid = not result.errors
    return result


def _render_table_grid(rows, out, color: bool = False) -> None:
    """Render table rows as a columnar grid: one column-position per line,
    row types side by side so you can see what each position maps to."""
    if not rows:
        return

    # ── build plain-text header labels ────────────────────────────────────────
    headers = []
    for trow in rows:
        if trow.row_type == 'SKIP_IF':
            headers.append('SKIP_IF')
        else:
            headers.append(f'{trow.row_type}:{trow.multiplicity}')

    # ── column widths (plain text only — ANSI codes must not inflate these) ───
    n_cols = max(len(trow.columns) for trow in rows)
    widths = []
    for ri, trow in enumerate(rows):
        w = len(headers[ri])
        for ci in range(n_cols):
            if ci < len(trow.columns):
                w = max(w, len(trow.columns[ci].field))
        widths.append(w)

    # ── color helpers — pad FIRST (plain length), then paint ──────────────────
    _ROW_COLOR = {
        'HEADER': 'cyan', 'DATA': 'green', 'FOOTER': 'dim', 'SKIP_IF': 'yellow',
    }
    _SENTINEL_FIELDS = {'EMPTY', 'IGNORE'}

    def _pad_paint(plain: str, color_name: str, width: int) -> str:
        """Right-pad *plain* to *width*, then apply color (preserving alignment)."""
        padded = f'{plain:<{width}}'
        return paint(padded, color_name, color)

    # ── header row ─────────────────────────────────────────────────────────────
    parts = []
    for i, h in enumerate(headers):
        row_type = h.split(':')[0]
        c = _ROW_COLOR.get(row_type, 'cyan')
        parts.append(_pad_paint(h, c, widths[i]))
    print(f'        {"  ".join(parts)}', file=out)

    # ── column rows (one per column position) ─────────────────────────────────
    for ci in range(n_cols):
        parts = []
        for ri, trow in enumerate(rows):
            val = trow.columns[ci].field if ci < len(trow.columns) else ''
            if val in _SENTINEL_FIELDS:
                parts.append(_pad_paint(val, 'dim', widths[ri]))
            elif val:
                parts.append(_pad_paint(val, 'cyan', widths[ri]))
            else:
                parts.append(' ' * widths[ri])
        print(f'        {"  ".join(parts)}', file=out)


def render_result(result: CheckResult, verbose: bool = False, out=None) -> None:
    """Print a human-readable report for one CheckResult."""
    out = out or sys.stderr
    color = should_color(out)
    fpath = paint(result.path, 'bold', color)
    if result.valid:
        stats = paint(f'({result.n_fields} field(s), {result.n_steps} extraction step(s))',
                      'dim', color)
        valid = paint('VALID', 'bold_green', color)
        print(colorize_marks(f'{_MARK_OK}  {fpath}  —  {valid} {stats}', color), file=out)
    else:
        invalid = paint('INVALID', 'bold_red', color)
        print(colorize_marks(f'{_MARK_FAIL}  {fpath}  —  {invalid}', color), file=out)
    for err in result.errors:
        print(colorize_marks(f'   {_MARK_FAIL} {paint(err, "red", color)}', color), file=out)
    for warn in result.warnings:
        print(colorize_marks(f'   {_MARK_WARN} {paint(warn, "yellow", color)}', color), file=out)

    if verbose and result.defs is not None:
        cfg = result.config
        # ── helpers ──────────────────────────────────────────────────────────
        def _key(k):   return paint(f'{k}', 'dim', color)
        def _sec(s):   return paint(s, 'dim', color)
        def _faint(v): return paint(str(v), 'dim', color)

        # ── config block ──────────────────────────────────────────────────────
        if cfg.pattern_version_explicit:
            ver = str(cfg.pattern_version)
        else:
            ver = f'{cfg.pattern_version} {_faint("(defaulted — no pattern.version declared)")}'
        aliases_val = (', '.join(cfg.empty_aliases)
                       if cfg.empty_aliases else _faint('(none)'))
        print(f'\n   {_sec("config:")}', file=out)
        print(f'     {_key("pattern.version")} {ver}', file=out)
        print(f'     {_key("read.direction")}  {cfg.read_direction}', file=out)
        print(f'     {_key("currency.sign")}   {cfg.currency_sign}', file=out)
        print(f'     {_key("ignore.case")}     {cfg.ignore_case}', file=out)
        print(f'     {_key("trim.whitespace")} {cfg.trim_whitespace}', file=out)
        print(f'     {_key("lbl.match")}       {cfg.lbl_match}', file=out)
        print(f'     {_key("var.match")}       {cfg.var_match}', file=out)
        print(f'     {_key("empty.aliases")}   {aliases_val}', file=out)

        # ── fields block ──────────────────────────────────────────────────────
        print(f'   {_sec("fields:")}', file=out)
        for name, fd in result.defs.items():
            tags = []
            if fd.role == 'lbl' and fd.lbl_match is not None:
                tags.append(fd.lbl_match)
            if fd.role == 'var' and fd.var_mode is not None:
                tags.append(fd.var_mode)
            if fd.required:
                tags.append('not-null')
            mode_tag = _faint(f' [{", ".join(tags)}]') if tags else ''
            role_color = 'yellow' if fd.role == 'lbl' else 'green'
            role  = paint(f'{fd.role:<4}', role_color, color)
            fname = paint(f'{name:<24}', 'cyan', color)
            ftype = _faint(f'{fd.type:<10}')
            regex = _faint(f'/{fd.regex}/')
            print(f'     {role} {fname} {ftype} {regex}{mode_tag}', file=out)

        # ── extraction sequence ───────────────────────────────────────────────
        print(f'   {_sec("extraction sequence:")}', file=out)
        arrow = paint('->', 'dim', color)
        for instr in (result.sequence or []):
            if isinstance(instr, CellInstruction):
                tgt   = instr.target or instr.multiplicity
                step  = paint(f'cell:{tgt:<6}', 'cyan', color)
                fname = paint(instr.field, 'cyan', color)
                print(f'     {step} {arrow} {fname}', file=out)
            elif isinstance(instr, SeekInstruction):
                step = paint(f'seek:{instr.target}', 'yellow', color)
                print(f'     {step}', file=out)
            elif isinstance(instr, DirectionInstruction):
                step = paint(f'dir:{instr.direction}', 'dim', color)
                print(f'     {step}', file=out)
            elif isinstance(instr, TableInstruction):
                step = paint(f'table:{instr.multiplicity}', 'green', color)
                print(f'     {step}', file=out)
                _render_table_grid(instr.rows, out, color)


def run_validate(paths: list[str], verbose: bool = False, quiet: bool = False,
                 out=None) -> int:
    """Validate each pattern file; return 0 if all valid, 1 otherwise.

    quiet=True suppresses the '✓ VALID' confirmation line; warnings and errors
    are still printed so the caller knows what failed.  Exit code is unchanged.
    """
    out = out or sys.stderr
    all_valid = True
    for idx, path in enumerate(paths):
        if idx:
            print('', file=out)
        result = check_pattern(path)
        if quiet and result.valid and not result.warnings:
            # Silent on clean success — only surface problems.
            pass
        else:
            render_result(result, verbose=verbose, out=out)
        all_valid = all_valid and result.valid
    return 0 if all_valid else 1
