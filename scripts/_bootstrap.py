"""
Auto-activates the project virtual environment without requiring the user
to run 'source .venv/bin/activate' first.

Each script in scripts/ calls ensure_venv() as its first action. If the
script is already running inside the venv, the call is a no-op. Otherwise
the process is re-executed with the venv Python — transparent to the user.
"""

import os
import subprocess
import sys


def ensure_venv() -> None:
    root     = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    venv_py  = os.path.join(root, '.venv', 'bin', 'python3')

    if not os.path.exists(venv_py):
        print(
            'Error: virtual environment not found.\n'
            'Set it up with:\n'
            '  python3 -m venv .venv\n'
            '  .venv/bin/pip install -e .',
            file=sys.stderr,
        )
        sys.exit(1)

    if os.path.abspath(sys.executable) != os.path.abspath(venv_py):
        # nosec B603 — fixed venv interpreter path (built from __file__, not user
        # input), no shell; sys.argv is passed as argv to python, never parsed by
        # a shell. Re-exec of the same script under the venv interpreter.
        sys.exit(subprocess.run([venv_py] + sys.argv).returncode)  # nosec B603
