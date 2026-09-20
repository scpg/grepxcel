"""Conftest for smoke tests.

tests/conftest.py inserts the project root into sys.path[0] so the source
tree's grepxcel/ is importable in the regular test suite.  Smoke tests must
import the INSTALLED package from site-packages instead, so we undo that
insertion here (this conftest is loaded after the parent one).
"""
import os
import sys

_project_root = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..')
)
try:
    sys.path.remove(_project_root)
except ValueError:
    pass  # already absent (e.g. running with --import-mode=importlib alone)
