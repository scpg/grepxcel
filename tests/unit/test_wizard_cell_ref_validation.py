"""A cell ref is validated at entry and never built into JavaScript.

PR #86 closed the CSRF and DNS-rebinding half of a two-part finding: a page on
another origin could drive the wizard's API. It did not close the other half.
`/api/classify`, `/api/classify-batch` and `/api/note` validated `action`
against an allow-list but accepted *any* non-empty string as a ref, and
wizard.js spliced that value straight into

    onclick="jumpToCell('<ref>')"

An apostrophe in the ref closes the JS string literal and whatever follows runs.
`escHtml()` would not have helped: it escaped `& < > "` and not `'`, and the
value sat inside single quotes.

Three layers, each tested here:

1. the ref is rejected at the endpoint, so a malformed value never reaches
   session state, the session log, or a sink added later;
2. the handler no longer interpolates into JavaScript at all — the ref goes in
   a data attribute and the click is bound by delegation;
3. `escHtml()` now escapes the apostrophe, as backup for every other caller.

Reachability after #86 is a crafted *pattern file* loaded through the wizard's
own Load Pattern feature, not a drive-by from another site. Note also that
`.upper()` on the ref incidentally mangled most script payloads, since JS
identifiers are case-sensitive — that is an accident, not a control, and
nothing should depend on it.
"""
import re
from pathlib import Path

import pytest

# wizard_api imports FastAPI, which the main CI test job does not install — see
# the "CI blind spot" section of docs/STATUS.md. The checks that only read
# wizard.js as text need no dependency at all, and those are the ones guarding
# the injection sink, so they are deliberately left to run everywhere.
try:
    from fastapi import HTTPException

    from grepxcel.wizard_api import _CELL_REF_RE, _require_cell_ref
    _API_OK = True
except ImportError:                                  # pragma: no cover
    _API_OK = False

_skip_no_api = pytest.mark.skipif(
    not _API_OK,
    reason='fastapi / uvicorn not installed (pip install "grepxcel[web]")',
)

_JS = Path(__file__).resolve().parents[2] / 'grepxcel' / 'static' / 'wizard.js'


# ── layer 1: the endpoint rejects it ─────────────────────────────────────────

@_skip_no_api
@pytest.mark.parametrize('ref', ['A1', 'B12', 'Z99', 'AB12', 'XFD1048576', 'A1048576'])
def test_well_formed_refs_are_accepted(ref):
    assert _require_cell_ref(ref) == ref


@_skip_no_api
@pytest.mark.parametrize('ref', [
    "A1'); alert(1); //",       # the payload this exists to stop
    "A1'",                      # the single character that broke the sink
    'A1"',
    'A1<script>',
    '<img src=x onerror=1>',
    'ABCD1',                    # four column letters; Excel's max is XFD
    'A0',                       # row 0 does not exist
    'A01',                      # leading zero would key state differently
    '1A',
    'A',
    '1',
    '',
    'a1',                       # callers upper-case first; this skipped that
    'A1 ',
    ' A1',
    'A1:B2',                    # a range is not a cell
    'Sheet1!A1',
    'A1\n',
    'A1' + chr(0),              # a literal NUL here would corrupt this file
])
def test_malformed_refs_are_refused(ref):
    with pytest.raises(HTTPException) as excinfo:
        _require_cell_ref(ref)
    assert excinfo.value.status_code == 400


@_skip_no_api
def test_the_refusal_names_the_offending_value():
    """So a user with a bad pattern file can see which ref is wrong."""
    with pytest.raises(HTTPException) as excinfo:
        _require_cell_ref('NOPE!', 'ref in refs')
    detail = str(excinfo.value.detail)
    assert 'NOPE!' in detail and 'ref in refs' in detail


@_skip_no_api
def test_none_is_refused_rather_than_crashing():
    with pytest.raises(HTTPException):
        _require_cell_ref(None)


@_skip_no_api
def test_the_pattern_is_anchored_at_both_ends():
    r"""An unanchored pattern would match a ref with a payload appended.

    The tail anchor must be `\Z`, not `$`: Python's `$` also matches just
    before a trailing newline, so `^...$` accepted "A1\n" — which the
    parametrized case above caught after this was first written with `$`.
    """
    assert _CELL_REF_RE.pattern.startswith('^')
    assert _CELL_REF_RE.pattern.endswith(r'\Z'), (
        'use \\Z rather than $ — $ permits a trailing newline'
    )


# ── layer 2: no ref is interpolated into JavaScript ──────────────────────────

def test_no_inline_onclick_builds_javascript_from_a_ref():
    """The sink itself. An inline handler built by string interpolation is the
    shape of the original defect, so it is banned rather than escaped."""
    source = _JS.read_text(encoding='utf-8')
    offenders = [
        line.strip()
        for line in source.splitlines()
        if re.search(r'onclick\s*=\s*["\'][^"\']*\$\{', line)
    ]
    assert not offenders, (
        'a template literal is being interpolated into an inline onclick:\n  '
        + '\n  '.join(offenders)
        + '\nUse a data attribute and a delegated listener instead — escaping '
          'is not sufficient when the value lands inside a JS string.'
    )


def test_the_provenance_chip_uses_a_data_attribute():
    source = _JS.read_text(encoding='utf-8')
    assert 'data-ref="${escHtml(r)}"' in source, (
        'the provenance chip no longer carries its ref in a data attribute — '
        'if it was refactored, keep the click delegated rather than inline'
    )


def test_a_delegated_listener_handles_the_chip():
    """Otherwise the data attribute is set and nothing ever reads it, and the
    chips silently stop working."""
    source = _JS.read_text(encoding='utf-8')
    assert ".et-ref[data-ref]" in source and 'chip.dataset.ref' in source


# ── layer 3: escHtml covers the apostrophe ───────────────────────────────────

def test_eschtml_escapes_the_apostrophe():
    """It escaped & < > " but not ' — which is the character that broke out of
    a single-quoted JS string inside a double-quoted attribute."""
    source = _JS.read_text(encoding='utf-8')
    body = source.split('function escHtml', 1)[1].split('}', 1)[0]
    for char, entity in (('&', '&amp;'), ('<', '&lt;'), ('>', '&gt;'),
                         ('"', '&quot;'), ("'", '&#39;')):
        assert entity in body, f'escHtml no longer escapes {char!r}'
