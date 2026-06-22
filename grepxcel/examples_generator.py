"""Generate working examples that new users can explore immediately."""

import importlib.resources
import os
import shutil
import sys

_EXAMPLES_PACKAGE = 'grepxcel.examples'

EXAMPLES = [
    {
        'name': '01_simple_invoice',
        'source_fixture': '01_simple_invoice',
        'description': 'Key-value extraction (cell:next) — 8 scalar fields',
        'pattern': '',
        'data': '',
    },
    {
        'name': '02_product_catalog',
        'source_fixture': '02_product_catalog',
        'description': 'Table extraction (HEADER/DATA) — 4-column repeating table',
        'pattern': '',
        'data': '',
    },
    {
        'name': '03_expense_report',
        'source_fixture': '05_expense_report',
        'description': 'Key-value + table + footer — 18 fields with Subtotal',
        'pattern': '',
        'data': '',
    },
    {
        'name': '04_loan_schedule',
        'source_fixture': '11_loan_schedule',
        'description': 'Key-value header + amortization table — 24 fields',
        'pattern': '',
        'data': '',
    },
]


def _resolve_bundled_paths():
    """Fill in the absolute paths to bundled xlsx files."""
    for ex in EXAMPLES:
        if ex['pattern']:
            continue
        pkg = f"{_EXAMPLES_PACKAGE}.{ex['name']}"
        try:
            ref = importlib.resources.files(pkg)
        except (ModuleNotFoundError, TypeError):
            pkg_dir = os.path.join(
                os.path.dirname(__file__), 'examples', ex['name'],
            )
            ex['pattern'] = os.path.join(pkg_dir, 'pattern.xlsx')
            ex['data'] = os.path.join(pkg_dir, 'data.xlsx')
            continue
        ex['pattern'] = str(ref.joinpath('pattern.xlsx'))
        ex['data'] = str(ref.joinpath('data.xlsx'))


_resolve_bundled_paths()


def generate_examples(output_dir: str) -> str:
    """Copy bundled examples into *output_dir*.

    Each example gets its own subdirectory with pattern.xlsx, data.xlsx,
    and a README.txt showing the exact extract command to run.

    Raises SystemExit if *output_dir* already exists and is non-empty.
    Returns the output directory path.
    """
    if os.path.isdir(output_dir) and os.listdir(output_dir):
        print(
            f"Error: directory '{output_dir}' already exists and is not empty.\n"
            f"Remove it first or choose a different path with -o.",
            file=sys.stderr,
        )
        raise SystemExit(1)

    os.makedirs(output_dir, exist_ok=True)
    print(f"Creating examples in: {output_dir}\n", file=sys.stderr)

    for ex in EXAMPLES:
        name = ex['name']
        dest = os.path.join(output_dir, name)
        os.makedirs(dest, exist_ok=True)

        shutil.copy2(ex['pattern'], os.path.join(dest, 'pattern.xlsx'))
        shutil.copy2(ex['data'], os.path.join(dest, 'data.xlsx'))

        readme = _build_readme(ex)
        with open(os.path.join(dest, 'README.txt'), 'w') as f:
            f.write(readme)

        print(f"  {name}/  — {ex['description']}", file=sys.stderr)

    print(
        f"\nDone! To try an example:\n"
        f"  cd {output_dir}/01_simple_invoice\n"
        f"  grepxcel extract -p pattern.xlsx data.xlsx\n",
        file=sys.stderr,
    )
    return output_dir


def _build_readme(ex: dict) -> str:
    name = ex['name']
    desc = ex['description']
    return (
        f"{name}\n"
        f"{'=' * len(name)}\n\n"
        f"{desc}\n\n"
        f"Quick start\n"
        f"-----------\n"
        f"  grepxcel extract -p pattern.xlsx data.xlsx\n\n"
        f"Verbose (see per-field trace):\n"
        f"  grepxcel extract -p pattern.xlsx data.xlsx -v\n\n"
        f"Validate the pattern (no extraction):\n"
        f"  grepxcel validate-pattern pattern.xlsx -v\n\n"
        f"Save output to JSON:\n"
        f"  grepxcel extract -p pattern.xlsx data.xlsx -o .\n"
    )
