"""
Pattern test suite — run a pattern against many files and report reliability.

Usage:
    grepxcel test -p pattern.xlsx samples/
    grepxcel test -p pattern.xlsx samples/ --recursive
    grepxcel test -p pattern.xlsx samples/ --format json

``grepxcel test`` answers the question "is my pattern ready for production?"
by running extraction against a directory of representative .xlsx files and
reporting field-level reliability (how many files each field extracted cleanly).

Exit codes:
    0  — all files passed (no errors, no warnings)
    1  — some files had warnings or partial extraction
    2  — one or more files failed completely, OR --strict and any issue

Output (default — human-readable):
    Tested 47 files
    ✅ 44 passed — all fields extracted cleanly
    ⚠️  2 warnings — partial extraction (some fields missing)
    ❌  1 failed  — extraction raised an error

    Field reliability:
      inv.number    47/47  ████████████████████  100%
      inv.date      45/47  ███████████████████░   96%
      vendor.name   47/47  ████████████████████  100%
      line_items    44/47  ██████████████████░░   94%

    Failures:
      ❌ northco_special.xlsx — anchor 'Invoice Number:' not found
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Any


# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class FileResult:
    path: str
    status: str            # 'pass' | 'warn' | 'fail'
    error: str | None
    result: dict | None
    issues: list[str]      # field-level warnings / missing fields


@dataclass
class PatternTestReport:
    pattern_path: str
    total: int
    passed: int
    warned: int
    failed: int
    file_results: list[FileResult]
    # field_name → (count_present, count_total)
    field_counts: dict[str, tuple[int, int]] = field(default_factory=dict)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _collect_xlsx_files(directory: str, recursive: bool) -> list[str]:
    """Return sorted list of .xlsx files, excluding Excel temp files."""
    result = []
    if recursive:
        for root, _, files in os.walk(directory):
            for f in sorted(files):
                if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                    result.append(os.path.join(root, f))
    else:
        for f in sorted(os.listdir(directory)):
            if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                result.append(os.path.join(directory, f))
    return result


def _extract_field_names(result: dict, prefix: str = '') -> set[str]:
    """Recursively collect all scalar leaf field paths from an extraction result."""
    names: set[str] = set()
    for key, value in result.items():
        if key.startswith('_'):
            continue
        full = f'{prefix}.{key}' if prefix else key
        if isinstance(value, dict):
            names.update(_extract_field_names(value, full))
        elif isinstance(value, list):
            # table — count as present if at least one data row
            names.add(full)
        else:
            names.add(full)
    return names


def _count_present(result: dict, prefix: str = '') -> set[str]:
    """Return field paths whose value is non-None and non-empty."""
    present: set[str] = set()
    for key, value in result.items():
        if key.startswith('_'):
            continue
        full = f'{prefix}.{key}' if prefix else key
        if isinstance(value, dict):
            present.update(_count_present(value, full))
        elif isinstance(value, list):
            if value:  # non-empty table list
                present.add(full)
        elif value is not None and value != '':
            present.add(full)
    return present


# ── Core runner ───────────────────────────────────────────────────────────────

def run_tests(
    pattern_path: str,
    directory: str,
    *,
    recursive: bool = False,
    sheet: str | None = None,
    strict: bool = False,
    max_size_mb: float = 5.0,
    max_uncompressed_mb: float = 50.0,
    on_progress: Any = None,
) -> PatternTestReport:
    """Run the pattern against all .xlsx files in *directory*.

    Args:
        pattern_path:        Path to the pattern file (.xlsx or .csv).
        directory:           Directory containing test .xlsx files.
        recursive:           Also descend into subdirectories.
        sheet:               Sheet name or 0-based index; default active sheet.
        strict:              Treat any missing field as a failure.
        max_size_mb:         Compressed file size limit passed to the engine.
        max_uncompressed_mb: Uncompressed content limit.
        on_progress:         Optional callable(path, FileResult) called after
                             each file is processed.

    Returns:
        PatternTestReport with per-file results and field reliability counts.

    Raises:
        FileNotFoundError:  if pattern_path or directory does not exist.
        ValueError:         if no .xlsx files are found in directory.
    """
    from .engine import Engine
    from .logger import Logger, VerbosityLevel

    pattern_path = os.path.abspath(pattern_path)
    directory = os.path.abspath(directory)

    if not os.path.isfile(pattern_path):
        raise FileNotFoundError(f'Pattern file not found: {pattern_path}')
    if not os.path.isdir(directory):
        raise FileNotFoundError(f'Test directory not found: {directory}')

    files = _collect_xlsx_files(directory, recursive)
    if not files:
        raise ValueError(f'No .xlsx files found in {directory}')

    # Track all field names seen across all results
    all_fields: set[str] = set()
    field_present: dict[str, int] = {}
    field_total: dict[str, int] = {}

    file_results: list[FileResult] = []

    for path in files:
        logger = Logger(level=VerbosityLevel.QUIET)
        engine = Engine()
        issues: list[str] = []

        try:
            result = engine.process(
                pattern_path, path,
                logger=logger,
                sheet=sheet,
                output_format='nested',
                max_file_mb=max_size_mb,
                max_uncompressed_mb=max_uncompressed_mb,
            )
            # Collect WARNING/ERROR records from the logger
            for rec in logger._records:
                if rec.severity in ('WARNING', 'ERROR'):
                    issues.append(rec.message)

            fields = _extract_field_names(result)
            present = _count_present(result)
            all_fields.update(fields)

            for f in fields:
                field_total[f] = field_total.get(f, 0) + 1
                if f in present:
                    field_present[f] = field_present.get(f, 0) + 1

            missing = fields - present
            if missing or issues:
                if strict or missing:
                    status = 'fail' if (strict and missing) else 'warn'
                else:
                    status = 'warn'
                issues.extend(f'missing: {m}' for m in sorted(missing))
            else:
                status = 'pass'

            fr = FileResult(
                path=path,
                status=status,
                error=None,
                result=result,
                issues=issues,
            )

        except Exception as exc:
            fr = FileResult(
                path=path,
                status='fail',
                error=str(exc),
                result=None,
                issues=[],
            )
            # Count all known fields as not-present for this file
            for f in all_fields:
                field_total[f] = field_total.get(f, 0) + 1

        file_results.append(fr)
        if on_progress:
            on_progress(path, fr)

    # Build per-field counts (fields that only appeared in some files still
    # get a denominator equal to the number of files they were seen in)
    field_counts = {
        f: (field_present.get(f, 0), field_total[f])
        for f in sorted(field_total.keys())
    }

    passed = sum(1 for r in file_results if r.status == 'pass')
    warned = sum(1 for r in file_results if r.status == 'warn')
    failed = sum(1 for r in file_results if r.status == 'fail')

    return PatternTestReport(
        pattern_path=pattern_path,
        total=len(file_results),
        passed=passed,
        warned=warned,
        failed=failed,
        file_results=file_results,
        field_counts=field_counts,
    )


# ── Formatters ────────────────────────────────────────────────────────────────

_BAR_WIDTH = 20


def _bar(count: int, total: int, width: int = _BAR_WIDTH) -> str:
    filled = round(count / total * width) if total else 0
    return '█' * filled + '░' * (width - filled)


def format_human(report: PatternTestReport, color: bool = True) -> str:
    """Return a human-readable multi-line string summarising the test run."""
    lines = []

    ok = '✅' if color else 'OK'
    warn = '⚠️ ' if color else 'WARN'
    fail = '❌' if color else 'FAIL'

    lines.append(f'\nTested {report.total} file{"s" if report.total != 1 else ""}  '
                 f'— pattern: {os.path.basename(report.pattern_path)}\n')

    if report.passed == report.total:
        lines.append(f'{ok} All {report.total} passed — every field extracted cleanly')
    else:
        if report.passed:
            lines.append(f'{ok} {report.passed} passed')
        if report.warned:
            lines.append(f'{warn} {report.warned} {"warning" if report.warned == 1 else "warnings"} '
                         f'— partial extraction (some fields missing or null)')
        if report.failed:
            lines.append(f'{fail} {report.failed} failed '
                         f'— extraction error or required field missing')

    if report.field_counts:
        lines.append('\nField reliability:')
        max_name = max(len(k) for k in report.field_counts)
        for fname, (present, total) in sorted(report.field_counts.items()):
            pct = present / total * 100 if total else 0
            bar = _bar(present, total)
            lines.append(
                f'  {fname:<{max_name}}  {present:>{len(str(total))}}/{total}  '
                f'{bar}  {pct:5.1f}%'
            )

    failures = [r for r in report.file_results if r.status == 'fail']
    warnings = [r for r in report.file_results if r.status == 'warn']

    if failures:
        lines.append('\nFailures:')
        for r in failures:
            msg = r.error or ', '.join(r.issues[:3])
            lines.append(f'  {fail} {os.path.basename(r.path)} — {msg}')

    if warnings:
        lines.append('\nWarnings:')
        for r in warnings:
            msg = ', '.join(r.issues[:3])
            lines.append(f'  {warn} {os.path.basename(r.path)} — {msg}')

    lines.append('')
    return '\n'.join(lines)


def format_json(report: PatternTestReport) -> str:
    """Return a JSON string with the full test report."""
    data = {
        'pattern': report.pattern_path,
        'total': report.total,
        'passed': report.passed,
        'warned': report.warned,
        'failed': report.failed,
        'field_reliability': {
            k: {'present': p, 'total': t, 'pct': round(p / t * 100, 1) if t else 0}
            for k, (p, t) in report.field_counts.items()
        },
        'files': [
            {
                'path': r.path,
                'status': r.status,
                'error': r.error,
                'issues': r.issues,
            }
            for r in report.file_results
        ],
    }
    return json.dumps(data, indent=2)
