"""grepxcel self-test — post-install verification.

Runs a short suite of checks against the bundled examples using only the
installed package and the stdlib — no pytest, no network, no external files.

Exit codes: 0 = all pass, 1 = one or more failures.
"""
import sys

import grepxcel
from grepxcel.color import (
    MARK_FAIL, MARK_INFO, MARK_OK,
    colorize_marks, should_color,
)
from grepxcel.examples_generator import EXAMPLES

# Expected structural assertions per example: which groups must exist,
# whether they are dicts (scalars) or lists (tables), and a sample of
# mandatory scalar sub-fields.
_EXPECTED: dict[str, dict] = {
    '01_simple_invoice': {
        'scalars': {'invoice': ['number', 'date'], 'client': ['name']},
    },
    '02_product_catalog': {
        'tables': ['catalog'],
    },
    '03_expense_report': {
        'scalars': {'employee': ['name', 'department']},
        'tables': ['expense'],
    },
    '04_loan_schedule': {
        'scalars': {'loan': ['amount']},
        'tables': ['schedule'],
    },
}


def _check_result(name: str, result: dict) -> list[str]:
    spec = _EXPECTED.get(name, {})
    issues: list[str] = []
    for grp, fields in spec.get('scalars', {}).items():
        if grp not in result:
            issues.append(f"group '{grp}' missing")
            continue
        if not isinstance(result[grp], dict):
            issues.append(f"group '{grp}' expected dict, got {type(result[grp]).__name__}")
            continue
        for f in fields:
            if f not in result[grp]:
                issues.append(f"field '{grp}.{f}' missing")
    for grp in spec.get('tables', []):
        if grp not in result:
            issues.append(f"table group '{grp}' missing")
        elif not isinstance(result.get(grp), list) or not result[grp]:
            issues.append(f"table '{grp}' is empty or wrong type")
    return issues


def run_self_test(verbose: bool = False) -> int:
    """Run checks; return 0 on full pass, 1 on any failure."""
    out = sys.stderr
    color = should_color(out)

    def _mark(mark, text):
        print(colorize_marks(f'  {mark}  {text}', color), file=out)

    _mark(MARK_INFO, f'grepxcel {grepxcel.__version__}  —  self-test')
    print(file=out)

    total = passed = 0

    for ex in EXAMPLES:
        name = ex['name']

        # validate-pattern (parse only, no extraction)
        total += 1
        try:
            from grepxcel.pattern_parser import PatternParser
            PatternParser().parse(ex['pattern'])
            _mark(MARK_OK, f'validate  {name}')
            passed += 1
        except Exception as exc:
            _mark(MARK_FAIL, f'validate  {name}')
            print(f'           {exc}', file=out)

        # extract + structural assertions
        total += 1
        try:
            result = grepxcel.extract(ex['pattern'], ex['data'])
            issues = _check_result(name, result)
            if issues:
                _mark(MARK_FAIL, f'extract   {name}  ({ex["description"]})')
                for msg in issues:
                    print(f'           {msg}', file=out)
            else:
                _mark(MARK_OK, f'extract   {name}  ({ex["description"]})')
                passed += 1
        except Exception as exc:
            _mark(MARK_FAIL, f'extract   {name}  ({ex["description"]})')
            if verbose:
                import traceback
                traceback.print_exc(file=out)
            else:
                print(f'           {exc}', file=out)

    print(file=out)
    if passed == total:
        _mark(MARK_OK, f'All {total} checks passed.')
        return 0
    _mark(MARK_FAIL, f'{total - passed}/{total} checks FAILED.')
    return 1
