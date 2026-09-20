"""Generate working examples that new users can explore immediately."""

import importlib.resources
import os
import shutil
import sys

_EXAMPLES_PACKAGE = 'grepxcel.examples'

# Each entry maps a user-facing example name to a source fixture in
# tests/fixtures/<source_fixture>/. The bundled files in grepxcel/examples/
# are copies of the fixture's highest-priority pattern and data.xlsx, renamed
# to the simpler pattern.xlsx / data.xlsx for a clean new-user experience.
# Priority: pattern-manual.xlsx > pattern-from-claude.xlsx > pattern-from-draft.xlsx
#
# Mapping: example name  →  source fixture
#   01_simple_invoice    →  tests/fixtures/01_simple_invoice/
#   02_product_catalog   →  tests/fixtures/02_product_catalog/
#   03_expense_report    →  tests/fixtures/05_expense_report/   (renumbered for UX)
#   04_loan_schedule     →  tests/fixtures/11_loan_schedule/    (renumbered for UX)
#
# To update bundled examples after changing a fixture pattern:
#   Run scripts/generate_fixtures.py (updates pattern-from-draft + data files)
#   Then copy the highest-priority pattern and updated data:
#     cp tests/fixtures/<source_fixture>/<source_fixture>_data.xlsx \
#        grepxcel/examples/<example_name>/data.xlsx
#     cp tests/fixtures/<source_fixture>/<source_fixture>_pattern-manual.xlsx \
#        grepxcel/examples/<example_name>/pattern.xlsx
#   (or _pattern-from-claude.xlsx if no manual pattern exists)
EXAMPLES = [
    {
        'name': '01_simple_invoice',
        'source_fixture': '01_simple_invoice',  # → tests/fixtures/01_simple_invoice/
        'description': 'Key-value extraction (cell:next) — 8 scalar fields',
        'title': 'Simple Invoice',
        'what_it_shows': (
            'A pattern that extracts scalar header fields from a typical invoice:\n'
            'invoice number, date, vendor name, client details, and amount subtotals.\n'
            'All fields are scalars — single-cell values, no repeating table rows.\n'
            'Start here if you are new to grepxcel.'
        ),
        'extra_commands': [
            ('Save output to a file',
             'grepxcel extract -p pattern.xlsx data.xlsx -o output.json'),
            ('Fail if any field is missing (CI mode)',
             'grepxcel extract -p pattern.xlsx data.xlsx --strict'),
        ],
        'pattern': '',
        'data': '',
    },
    {
        'name': '02_product_catalog',
        'source_fixture': '02_product_catalog',  # → tests/fixtures/02_product_catalog/
        'description': 'Table extraction (HEADER/DATA) — 4-column repeating table',
        'title': 'Product Catalog',
        'what_it_shows': (
            'A pattern that extracts a repeating table of products:\n'
            'SKU, product name, unit price, and stock level.\n'
            'Table patterns use HEADER/DATA rows to describe column structure.\n'
            'This is the right starting point for any sheet with repeating rows.'
        ),
        'extra_commands': [
            ('Export directly to CSV (single-table patterns)',
             'grepxcel extract -p pattern.xlsx data.xlsx --format csv'),
            ('Save JSON to a file',
             'grepxcel extract -p pattern.xlsx data.xlsx -o output.json'),
        ],
        'pattern': '',
        'data': '',
    },
    {
        'name': '03_expense_report',
        'source_fixture': '05_expense_report',  # → tests/fixtures/05_expense_report/ (renumbered)
        'description': 'Key-value + table + footer — 18 fields with Subtotal',
        'title': 'Expense Report',
        'what_it_shows': (
            'A mixed pattern combining three common shapes:\n'
            '  - Header scalars (employee name, date, department)\n'
            '  - A repeating table of expense line items\n'
            '  - A footer row with the reimbursable subtotal\n'
            'This is the most common real-world pattern shape.'
        ),
        'extra_commands': [
            ('See the full extraction trace',
             'grepxcel extract -p pattern.xlsx data.xlsx -v'),
            ('Export as colored Excel report',
             'grepxcel extract -p pattern.xlsx data.xlsx --format xlsx -o .'),
        ],
        'pattern': '',
        'data': '',
    },
    {
        'name': '04_loan_schedule',
        'source_fixture': '11_loan_schedule',  # → tests/fixtures/11_loan_schedule/ (renumbered)
        'description': 'Key-value header + amortization table — 24 fields',
        'title': 'Loan Schedule',
        'what_it_shows': (
            'A header block (loan amount, interest rate, term, monthly payment)\n'
            'followed by a large amortization table (payment number, date,\n'
            'principal, interest, and remaining balance for each period).\n'
            'Tests pattern performance on larger tables with many rows.'
        ),
        'extra_commands': [
            ('Get DataFrames in Python (requires pandas)',
             "python3 -c \""
             "import grepxcel; frames = grepxcel.extract_df('pattern.xlsx', 'data.xlsx'); "
             "print(frames['schedule'])\""),
        ],
        'pattern': '',
        'data': '',
    },
    {
        'name': '05_ap_aging',
        'source_fixture': '23_ap_aging',  # → tests/fixtures/23_ap_aging/
        'description': 'Enterprise AP Aging — merged header, metadata scalars, SKIP_IF subtotals',
        'title': 'AP Aging Report',
        'what_it_shows': (
            'A realistic ERP-export accounts-payable aging report:\n'
            'a merged company-name header, metadata rows (report date, period),\n'
            'a transaction table with interleaved vendor subtotal rows and\n'
            'a grand total row that are filtered out with SKIP_IF.\n'
            'Demonstrates seek:, absolute cell references for scalars,\n'
            'and SKIP_IF with a regexp label to clean up ERP noise.'
        ),
        'extra_commands': [
            ('Extract only pending invoices (filter in Python)',
             "python3 -c \""
             "import grepxcel; r = grepxcel.extract('pattern.xlsx', 'data.xlsx'); "
             "pending = [t for t in r['txn'][0]['data'] if t['payment_status']=='Pending']; "
             "print(f'{len(pending)} pending invoices')\""),
            ('Export to CSV for downstream ingestion',
             'grepxcel extract -p pattern.xlsx data.xlsx --format csv'),
        ],
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

    Returns the output directory path.
    """
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
    title = ex.get('title', name)
    heading = f"{name} — {title}"
    what = ex.get('what_it_shows', ex['description'])
    extra = ex.get('extra_commands', [])

    lines = [
        heading,
        '=' * len(heading),
        '',
        'What this example shows',
        '-----------------------',
        what,
        '',
        'Files',
        '-----',
        '  pattern.xlsx  — the pattern: describes what to look for',
        '  data.xlsx     — a sample data file',
        '',
        'Try it',
        '------',
        'Extract and print JSON:',
        '  grepxcel extract -p pattern.xlsx data.xlsx',
        '',
        'See the field-by-field extraction trace:',
        '  grepxcel extract -p pattern.xlsx data.xlsx -v',
        '',
        'Validate the pattern without extracting:',
        '  grepxcel validate-pattern pattern.xlsx -v',
        '',
    ]

    if extra:
        lines.append('More options')
        lines.append('------------')
        for label, cmd in extra:
            lines.append(f'{label}:')
            lines.append(f'  {cmd}')
            lines.append('')

    lines += [
        'Build a pattern for your own files',
        '-----------------------------------',
        'Browser wizard (click cells to classify):',
        "  grepxcel web-wizard your-file.xlsx        # requires: pip install 'grepxcel[web]'",
        '',
        'AI-assisted starter pattern:',
        '  grepxcel draft your-file.xlsx -o my-pattern.xlsx',
        '',
        'Test a pattern against a folder of files:',
        '  grepxcel test -p pattern.xlsx your-folder/',
        '',
        'Learn more',
        '----------',
        '  grepxcel quickstart       — guided tutorial in your terminal',
        '  grepxcel doctor           — check your environment is ready',
        '  grepxcel docs             — write the full pattern reference to disk',
        '',
    ]
    return '\n'.join(lines)
