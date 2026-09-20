"""Conftest for smoke tests.

tests/conftest.py inserts the project root into sys.path[0] so the source
tree's grepxcel/ is importable in the regular test suite.  Smoke tests must
import the INSTALLED package from site-packages instead, so we undo that
insertion here (this conftest is loaded after the parent one).

The insertion is stored as an unnormalized path (e.g. '/…/tests/..') so a
simple sys.path.remove() against the normalized form misses it.  We filter by
comparing normalized paths instead.
"""
import os
import sys

_project_root = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..')
)
sys.path[:] = [p for p in sys.path if os.path.normpath(p) != _project_root]
