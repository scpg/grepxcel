#!/usr/bin/env python3
"""The pre-PR gate: run the fast suite, then the oracle type-matrix suite.

Usage:
    .venv/bin/python3 scripts/oracle_check.py

Runs, in order, stopping at the first failure:
  1. the normal fast suite      (.venv/bin/pytest tests/ -q)
  2. the oracle matrix suite    (.venv/bin/pytest -m oracle tests/oracle -q)

Prints a one-line verdict and exits 0 (green) or 1 (red). The oracle suite is
deliberately NOT in CI — it is too slow for the per-push loop — so this command
is what stands between a branch and `gh pr create`.

Narrow the oracle run while debugging a failure (the gate itself always runs the
full matrix):
    GREPXCEL_ORACLE_SEEDS=1 GREPXCEL_ORACLE_TYPES=integer \\
    GREPXCEL_ORACLE_SHAPES=cells1 .venv/bin/pytest -m oracle tests/oracle -q
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

import subprocess
import time

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
PYTEST = os.path.join(ROOT, '.venv', 'bin', 'pytest')

def _oracle_cmd() -> list:
    """The oracle step, parallelised across cores when pytest-xdist is present.

    ~27k generated cases are almost entirely CPU-bound workbook round-trips with
    no shared state: each case builds into its own directory (keyed by pid) and
    every test uses tmp_path, so the suite distributes cleanly. Measured on a
    32-core machine: 124s serial, 26s with -n auto, identical counts.

    The FAST suite is deliberately left serial. It takes ~23s, so there is
    little to win, and it contains at least one test that is order-sensitive
    under xdist: proxy_support._truststore_inject() mutates SSL state
    process-wide, and a worker process shared across tests can carry that into
    a later, fully-mocked backend test (seen once as X509:
    NO_CERTIFICATE_OR_CRL_FOUND under -n 16). That is a pre-existing isolation
    bug, not a parallelism bug — but the gate should not be the place it shows up.
    """
    cmd = [PYTEST, '-m', 'oracle', 'tests/oracle', '-q']
    try:
        import xdist  # noqa: F401
    except ImportError:
        return cmd
    return cmd + ['-n', 'auto']


STEPS = [
    ('fast suite', [PYTEST, 'tests/', '-q']),
    ('oracle matrix suite', _oracle_cmd()),
]

GREEN = 'ORACLE GATE: GREEN — safe to open a PR'
RED = 'ORACLE GATE: RED — do not open a PR'


def main() -> int:
    if not os.path.exists(PYTEST):
        print(f'Error: {PYTEST} not found. Set the venv up with:\n'
              f'  python3 -m venv .venv\n'
              f'  .venv/bin/pip install -e ".[dev]"', file=sys.stderr)
        print(RED)
        return 1

    for name, cmd in STEPS:
        print(f'── {name}: {" ".join(cmd[1:])}', flush=True)
        started = time.monotonic()
        # nosec B603 — fixed interpreter/pytest paths built from __file__, no
        # shell, no user input in argv.
        proc = subprocess.run(cmd, cwd=ROOT)  # nosec B603
        elapsed = time.monotonic() - started
        if proc.returncode != 0:
            print(f'\n{name} FAILED (exit {proc.returncode}) after {elapsed:.1f}s')
            print(RED)
            return 1
        print(f'   {name} passed in {elapsed:.1f}s\n', flush=True)

    print(GREEN)
    return 0


if __name__ == '__main__':
    sys.exit(main())
