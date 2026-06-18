#!/usr/bin/env python3
"""Apply color formatting to one or more pattern .xlsx files.

Usage:
    python3 scripts/colorize_pattern.py pattern.xlsx [pattern2.xlsx ...]
    python3 scripts/colorize_pattern.py tests/fixtures/*/pattern-from-draft.xlsx

Colors rows by type (config:, lbl:, var:, START:/END:, table rows, comments)
without changing any cell content. Column widths are auto-adjusted.
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

from grepxcel.pattern_colors import colorize_pattern_file


def main():
    if len(sys.argv) < 2:
        print(__doc__.strip(), file=sys.stderr)
        sys.exit(2)

    paths = sys.argv[1:]
    for path in paths:
        if not path.endswith('.xlsx'):
            print(f'  skip  {path} (not .xlsx)', file=sys.stderr)
            continue
        if not os.path.isfile(path):
            print(f'  ✗  {path} (not found)', file=sys.stderr)
            continue
        colorize_pattern_file(path)
        print(f'  ✓  {path}', file=sys.stderr)

    print(f'\nColorized {len(paths)} file(s).', file=sys.stderr)


if __name__ == '__main__':
    main()
