"""The ``grepxcel quickstart`` command — a guided tutorial in your terminal."""

import sys


def _color(text: str, code: str) -> str:
    if not sys.stderr.isatty():
        return text
    return f'\033[{code}m{text}\033[0m'


def _bold(text: str) -> str:
    return _color(text, '1')


def _dim(text: str) -> str:
    return _color(text, '2')


def _cyan(text: str) -> str:
    return _color(text, '36')


def _green(text: str) -> str:
    return _color(text, '32')


def run_quickstart() -> int:
    from . import __version__
    lines = [
        '',
        _bold(f'  grepxcel {__version__} — Quick Start Guide'),
        '',
        _bold('  What is grepxcel?'),
        '',
        '  grepxcel extracts structured data from Excel files. You describe',
        '  the layout of the file once in a pattern, and grepxcel reads any',
        '  matching file and gives you clean JSON.',
        '',
        _dim('  Think of it as: pattern + Excel file = JSON output.'),
        '',
        '─' * 64,
        '',
        _bold('  Step 1: Generate example files'),
        '',
        _cyan('    $ grepxcel generate-examples'),
        _cyan('    $ cd grepxcel-examples/01_simple_invoice'),
        '',
        '  This creates a folder with a pattern file and a data file',
        '  you can try immediately.',
        '',
        '─' * 64,
        '',
        _bold('  Step 2: Run your first extraction'),
        '',
        _cyan('    $ grepxcel extract -p pattern.xlsx data.xlsx'),
        '',
        '  You will see JSON output like this:',
        '',
        _green('    {'),
        _green('      "inv": { "number": "AB123456", "date": "2026-03-15" },'),
        _green('      "client": { "name": "Alice Wonderland" },'),
        _green('      "amount": { "net": 120, "vat": 24, "gross": 144 }'),
        _green('    }'),
        '',
        '─' * 64,
        '',
        _bold('  Step 3: Understand the pattern file'),
        '',
        '  Open pattern.xlsx (or pattern.csv) in Excel or a text editor.',
        '  Each row plays one of four roles:',
        '',
        f'    {_bold("config:")}  Settings (scan direction, currency symbol)',
        f'    {_bold("lbl:")}     Landmarks — text used to find your place',
        f'    {_bold("var:")}     Fields — values you want to extract',
        f'    {_bold("doc:")}     Notes — ignored by the engine',
        '',
        '  Between START: and END: you list the cells to read, in order.',
        '',
        '─' * 64,
        '',
        _bold('  Step 4: Create your own pattern'),
        '',
        '  Option A — let AI draft one for you:',
        _cyan('    $ grepxcel draft your-file.xlsx -o my-pattern.xlsx'),
        '',
        '  Option B — start from the reference:',
        _cyan('    $ grepxcel docs -o pattern-reference.xlsx'),
        '  Open it — every instruction is explained with examples.',
        '',
        '  Option C — validate as you go:',
        _cyan('    $ grepxcel validate-pattern my-pattern.xlsx -v'),
        '',
        '─' * 64,
        '',
        _bold('  Useful commands'),
        '',
        f'    {_cyan("grepxcel extract -p pat.xlsx data.xlsx")}    Extract data',
        f'    {_cyan("grepxcel extract ... --strict")}              Fail on any issue',
        f'    {_cyan("grepxcel extract ... -v")}                    See step-by-step log',
        f'    {_cyan("grepxcel lint data.xlsx")}                    Check a file first',
        f'    {_cyan("grepxcel docs")}                              Write local guide (.xlsx + .docx)',
        f'    {_cyan("grepxcel doctor")}                            Verify your setup',
        '',
        '─' * 64,
        '',
        _bold('  Learn more'),
        '',
        '  Pattern reference:   https://github.com/scpg/grepxcel/blob/main/docs/pattern-file.md',
        '  CLI reference:       https://github.com/scpg/grepxcel/blob/main/docs/cli-reference.md',
        '  Design philosophy:   https://github.com/scpg/grepxcel/blob/main/docs/MINDSET.md',
        '  Local guide:         grepxcel docs  (writes pattern-reference.xlsx + grepxcel-guide.docx)',
        '',
        _bold('  TAB completion'),
        '',
        '  Add to ~/.bashrc or ~/.zshrc to enable:',
        '    eval "$(register-python-argcomplete grepxcel)"',
        '',
    ]
    print('\n'.join(lines), file=sys.stderr)
    return 0
