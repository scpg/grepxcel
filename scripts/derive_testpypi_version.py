#!/usr/bin/env python3
"""Derive the next unused TestPyPI pre-release version.

Usage:
    python3 scripts/derive_testpypi_version.py [base-version]

Prints one version string on stdout, e.g. 0.5.0rc47.

The rc counter is scoped to the BASE VERSION, not to the CI run. We ask
TestPyPI which pre-releases it already holds for the current __version__ and
take the next free number, so a bump to 0.6.0 restarts at rc1 and the number
means "how many times this version was rehearsed" — the only thing it can
usefully mean.

The previous scheme used github.run_number, a per-workflow counter that never
resets, which is why TestPyPI holds 0.1.0rc16 next to 0.3.0rc17: the version
moved and the counter did not.

Deliberately stdlib-only, and deliberately does NOT call ensure_venv(): this
runs on a bare GitHub Actions runner before the project is installed.
"""
import json
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

PACKAGE      = 'grepxcel'
INDEX_JSON   = f'https://test.pypi.org/pypi/{PACKAGE}/json'
VERSION_FILE = Path(__file__).resolve().parent.parent / PACKAGE / '__init__.py'
TIMEOUT      = 30


def read_base_version() -> str:
    """The X.Y.Z declared in the package, ignoring any pre-release suffix."""
    text  = VERSION_FILE.read_text(encoding='utf-8')
    match = re.search(r"""^__version__\s*=\s*['"](\d+\.\d+\.\d+)""", text, re.M)
    if not match:
        sys.exit(f'Error: no __version__ = "X.Y.Z" found in {VERSION_FILE}')
    return match.group(1)


def published_releases() -> list[str]:
    """Every version string TestPyPI already holds for this project.

    A 404 means the project has never been uploaded — an empty history, not an
    error. Every other failure is fatal on purpose: guessing "empty" after a
    transient network fault would hand back an rc number that already exists,
    skip-existing would swallow the collision, and the smoke test would then
    happily validate a months-old build as if it were the new one.
    """
    try:
        # nosec B310 — fixed https:// module constant, never user input.
        with urllib.request.urlopen(INDEX_JSON, timeout=TIMEOUT) as response:  # nosec B310
            return list(json.load(response).get('releases', {}))
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return []
        sys.exit(f'Error: TestPyPI returned HTTP {exc.code} for {INDEX_JSON}')
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        sys.exit(f'Error: could not reach TestPyPI ({exc})')


def next_candidate(base: str, releases: list[str]) -> str:
    """base plus the lowest rc number not already taken for that base."""
    pattern = re.compile(rf'^{re.escape(base)}rc(\d+)$')
    used    = [int(m.group(1)) for m in map(pattern.match, releases) if m]
    return f'{base}rc{max(used, default=0) + 1}'


def main() -> None:
    base = sys.argv[1] if len(sys.argv) > 1 else read_base_version()
    print(next_candidate(base, published_releases()))


if __name__ == '__main__':
    main()
