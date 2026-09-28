"""Clean up the generated case tree when the oracle session ends.

The matrix suite builds every case under a pid-keyed directory in the system
temp dir (`test_oracle_matrix._ROOT`) and, until this existed, never removed it.
One run left ~75 MB per process; under `-n auto` that is one directory per xdist
worker, so a 32-core run leaked ~2.4 GB per invocation. On a machine where
/tmp is tmpfs that is resident memory, and it accumulates silently across runs
until a later run dies partway through with `No space left on device` — which
surfaces as thousands of unrelated oracle assertion failures rather than as a
disk error, and sends you looking for a regression that is not there.

Set `GREPXCEL_ORACLE_KEEP=1` to keep the tree. That matters: when a case fails,
the generated .xlsx and pattern are the evidence, and a suite that deletes them
on the way out makes its own failures harder to diagnose. The gate prints the
path it kept so it can be found.
"""
import os
import shutil

import pytest


def _keep_requested() -> bool:
    return os.environ.get('GREPXCEL_ORACLE_KEEP', '').strip().lower() not in (
        '', '0', 'false', 'no',
    )


@pytest.fixture(scope='session', autouse=True)
def _remove_generated_cases():
    """Remove the case tree after the session, unless asked to keep it.

    Only ever touches the directory this process owns — the path carries
    `os.getpid()`, so a concurrent run's tree is never removed.
    """
    yield

    from tests.oracle import test_oracle_matrix as matrix

    root = getattr(matrix, '_ROOT', None)
    if not root or not os.path.isdir(root):
        return

    if _keep_requested():
        print(f'\nGREPXCEL_ORACLE_KEEP set — generated cases kept at: {root}')
        return

    # Never raise from teardown: a failure to clean up must not turn a green
    # run red, and must not mask the real result of the run.
    shutil.rmtree(root, ignore_errors=True)
