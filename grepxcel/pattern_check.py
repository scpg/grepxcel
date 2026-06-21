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

from .color import colorize_marks, should_color
from .models import CellInstruction, TableInstruction, SeekInstruction, DirectionInstruction
from .pattern_parser import PatternError, PatternParser
from .security import SecurityError

# Patterns that strongly suggest regex intent (backslash-escapes, lookahead)
_REGEX_TELL = re.compile(r'\\[()[\]{}|+*.?^$]|[(][?]')

_MARK_OK, _MARK_WARN, _MARK_FAIL = '✓', '⚠', '✗'


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

    for name, fd in defs.items():
        if fd.role != 'lbl':
            continue
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

    result.valid = not result.errors
    return result


def _render_table_grid(rows, out) -> None:
    """Render table rows as a columnar grid: one column-position per line,
    row types side by side so you can see what each position maps to."""
    if not rows:
        return
    headers = []
    for trow in rows:
        if trow.row_type == 'SKIP_IF':
            headers.append('SKIP_IF')
        else:
            headers.append(f'{trow.row_type}:{trow.multiplicity}')
    n_cols = max(len(trow.columns) for trow in rows)
    widths = []
    for ri, trow in enumerate(rows):
        w = len(headers[ri])
        for ci in range(n_cols):
            if ci < len(trow.columns):
                w = max(w, len(trow.columns[ci].field))
        widths.append(w)
    parts = [f'{h:<{widths[i]}}' for i, h in enumerate(headers)]
    print(f'        {"  ".join(parts)}', file=out)
    for ci in range(n_cols):
        parts = []
        for ri, trow in enumerate(rows):
            val = trow.columns[ci].field if ci < len(trow.columns) else ''
            parts.append(f'{val:<{widths[ri]}}')
        print(f'        {"  ".join(parts)}', file=out)


def render_result(result: CheckResult, verbose: bool = False, out=None) -> None:
    """Print a human-readable report for one CheckResult."""
    out = out or sys.stderr
    color = should_color(out)
    if result.valid:
        print(colorize_marks(f'{_MARK_OK}  {result.path}  —  VALID '
              f'({result.n_fields} field(s), {result.n_steps} extraction step(s))',
              color), file=out)
    else:
        print(colorize_marks(f'{_MARK_FAIL}  {result.path}  —  INVALID', color), file=out)
    for err in result.errors:
        print(colorize_marks(f'   {_MARK_FAIL} {err}', color), file=out)
    for warn in result.warnings:
        print(colorize_marks(f'   {_MARK_WARN} {warn}', color), file=out)

    if verbose and result.defs is not None:
        cfg = result.config
        ver = (f'{cfg.pattern_version}' if cfg.pattern_version_explicit
               else f'{cfg.pattern_version} (defaulted — no pattern.version declared)')
        print('\n   config:', file=out)
        print(f'     pattern.version {ver}', file=out)
        print(f'     read.direction  {cfg.read_direction}', file=out)
        print(f'     currency.sign   {cfg.currency_sign}', file=out)
        print(f'     ignore.case     {cfg.ignore_case}', file=out)
        print(f'     lbl.match       {cfg.lbl_match}', file=out)
        print('   fields:', file=out)
        for name, fd in result.defs.items():
            mode_tag = f' [{fd.lbl_match}]' if fd.lbl_match is not None else ''
            print(f'     {fd.role:<4} {name:<24} {fd.type:<10} /{fd.regex}/{mode_tag}', file=out)
        print('   extraction sequence:', file=out)
        for instr in (result.sequence or []):
            if isinstance(instr, CellInstruction):
                tgt = instr.target or instr.multiplicity
                print(f'     cell:{tgt:<6} -> {instr.field}', file=out)
            elif isinstance(instr, SeekInstruction):
                print(f'     seek:{instr.target}', file=out)
            elif isinstance(instr, DirectionInstruction):
                print(f'     dir:{instr.direction}', file=out)
            elif isinstance(instr, TableInstruction):
                print(f'     table:{instr.multiplicity}', file=out)
                _render_table_grid(instr.rows, out)


def run_validate(paths: list[str], verbose: bool = False, out=None) -> int:
    """Validate each pattern file; return 0 if all valid, 1 otherwise."""
    out = out or sys.stderr
    all_valid = True
    for idx, path in enumerate(paths):
        if idx:
            print('', file=out)
        result = check_pattern(path)
        render_result(result, verbose=verbose, out=out)
        all_valid = all_valid and result.valid
    return 0 if all_valid else 1
