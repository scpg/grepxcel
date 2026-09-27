"""`os._exit` must stay behind the one named seam, so the test suite can report.

The wizard ends its own process on shutdown, which is legitimate: uvicorn does not
reliably hand back control when the request that triggered the shutdown is still in
flight. What is not legitimate is calling `os._exit` inline from a request handler.

That is not a style point. `POST /api/shutdown` arms a daemon thread that exits 0.3s
later, and three tests exercise that endpoint on the accept path. With the call inline,
the thread killed the pytest process 0.3s after those tests — which landed after the
final test and before pytest wrote its epilogue. The suite printed progress to 100%,
skipped the FAILURES section, and exited 0 while a test was genuinely failing. A test
suite that cannot report a failure is worse than no test suite, because it is trusted.

`tests/conftest.py` replaces `_terminate_process` for the whole session. That guard only
works while every exit path goes through it, which is what this file pins.
"""
import ast
from pathlib import Path

import pytest

_SOURCE = Path(__file__).resolve().parents[2] / 'grepxcel' / 'wizard_api.py'

#: The single function allowed to halt the interpreter.
_SEAM = '_terminate_process'


def _hard_exit_calls():
    """Every `os._exit(...)` call, paired with the function enclosing it."""
    tree = ast.parse(_SOURCE.read_text(encoding='utf-8'))
    found = []
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for inner in ast.walk(node):
            if not isinstance(inner, ast.Call):
                continue
            func = inner.func
            if (isinstance(func, ast.Attribute) and func.attr == '_exit'
                    and isinstance(func.value, ast.Name) and func.value.id == 'os'):
                found.append((node.name, inner.lineno))
    return found


def test_the_seam_exists():
    """Guards against the AST walk silently passing because the name changed."""
    tree = ast.parse(_SOURCE.read_text(encoding='utf-8'))
    names = {n.name for n in ast.walk(tree)
             if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
    assert _SEAM in names, (
        f'{_SEAM}() is gone from wizard_api.py. If the exit was renamed, update '
        f'this test AND the session fixture in tests/conftest.py that replaces it '
        f'— otherwise the suite silently loses its ability to report failures.'
    )


def test_os_exit_is_called_only_from_the_seam():
    offenders = [(fn, line) for fn, line in _hard_exit_calls() if fn != _SEAM]
    assert not offenders, (
        'os._exit() is called outside ' + _SEAM + '():\n'
        + '\n'.join(f'  wizard_api.py:{line} in {fn}()' for fn, line in offenders)
        + '\nRoute it through ' + _SEAM + '() instead. An inline call cannot be '
          'neutralised for tests, and a shutdown handler that exits inline kills '
          'the pytest process mid-run with status 0 — the suite then reports '
          'success while tests are failing.'
    )


def test_there_is_exactly_one_hard_exit():
    """More than one, and the conftest guard covers only some of them."""
    calls = _hard_exit_calls()
    assert len(calls) == 1, f'expected a single os._exit call, found {calls}'


@pytest.mark.parametrize('module', ['wizard_api', 'mcp_server', 'engine', 'cli'])
def test_no_module_calls_os_exit_outside_the_seam(module):
    """The same masking would apply anywhere else it appeared."""
    path = _SOURCE.parent / f'{module}.py'
    if not path.exists():
        pytest.skip(f'{module}.py not present')
    tree = ast.parse(path.read_text(encoding='utf-8'))
    bad = [
        node.lineno for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute) and node.func.attr == '_exit'
        and isinstance(node.func.value, ast.Name) and node.func.value.id == 'os'
    ]
    if module == 'wizard_api':
        # Its single permitted call lives in the seam, checked above.
        assert len(bad) == 1, f'wizard_api.py os._exit count changed: {bad}'
    else:
        assert not bad, (
            f'{module}.py calls os._exit at line(s) {bad}. Use sys.exit so pytest '
            f'can observe it, or add a named seam and neutralise it in conftest.'
        )
