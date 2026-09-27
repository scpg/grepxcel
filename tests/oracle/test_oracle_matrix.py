"""The oracle type-matrix entry point (spec §11).

One parametrized case per (direction x type x shape[x via|header_mode] x storage
case x number format x seed); six test functions, one per subcommand oracle, all
sharing one generated case per ``case_id`` via :func:`case_for`.

Excluded from the default run by the ``oracle`` marker (``pytest.ini`` sets
``addopts = -m "not oracle"``). Run it with::

    .venv/bin/pytest -m oracle tests/oracle -q

Narrow it while debugging with environment variables::

    GREPXCEL_ORACLE_SEEDS=1              # default 3
    GREPXCEL_ORACLE_TYPES=integer,url    # default: all 11
    GREPXCEL_ORACLE_SHAPES=cells1,table1 # default: all 6 (spec's phase-3 subset)

A failure always prints its ``case_id``, which contains the seed, so any red is
replayable by regenerating that one case — no committed fixture needed.

## Known limitations (§12) — documented, not solved here

* ``str`` (formula-cached string) and ``inlineStr`` storage cannot be written by
  openpyxl: a written formula has no cached value under ``data_only=True``, and
  openpyxl always emits shared strings. Real coverage lives in
  ``tests/fixtures/24_type_tests`` (50 ``str/string`` cells).
* ``image`` / richData cells: same, real fixtures only.
* Pattern-side variants (``var:literal``/``glob``, ``lbl:regexp``, ``not-null``,
  ``trim-whitespace``, ``nullable``, ``ignore.case``) are fixed to defaults; they
  are the v2 axis.
* ``FOOTER``, ``SPLITTER``, ``SKIP_IF``, ``SKIP_EMPTY_ROW``, ``DATA:{n,m}`` and
  ``table:{n,m}`` are out of scope — each has its own unit tests.
* ``validate_type``'s ``float.is_integer()`` branch is unreachable through a real
  .xlsx (openpyxl returns an ``int`` for an integral float) — see
  ``catalog.StorageCase.round_trip``.
"""
from __future__ import annotations

import os
import tempfile

import pytest

from grepxcel import cli
from tests.oracle import catalog, oracles
from tests.oracle.generator import (
    HEADER_MODES, SCALAR_SHAPES, SHAPES, TABLE_SHAPES, VIAS,
    CaseParams, build_case,
)

pytestmark = pytest.mark.oracle

DIRECTIONS = ('LR', 'TD')


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name, '').strip()
    try:
        return max(1, int(raw)) if raw else default
    except ValueError:
        return default


def _env_list(name: str, default: tuple) -> tuple:
    raw = os.environ.get(name, '').strip()
    if not raw:
        return default
    chosen = tuple(p.strip() for p in raw.split(',') if p.strip())
    return chosen or default


SEEDS = range(_env_int('GREPXCEL_ORACLE_SEEDS', 3))
TYPES = _env_list('GREPXCEL_ORACLE_TYPES', tuple(catalog.EXTRACTABLE_TYPES))
SHAPES_UNDER_TEST = _env_list('GREPXCEL_ORACLE_SHAPES', SHAPES)


def _build_params() -> list[CaseParams]:
    out: list[CaseParams] = []
    for direction in DIRECTIONS:
        for type_name in TYPES:
            for case in catalog.STORAGE_MATRIX[type_name]:
                for fmt in catalog.formats_for(type_name, case):
                    for seed in SEEDS:
                        for shape in SHAPES_UNDER_TEST:
                            if shape in SCALAR_SHAPES:
                                for via in VIAS:
                                    out.append(CaseParams(
                                        type_name, shape, case.key, fmt,
                                        direction, seed, via=via))
                            elif shape in TABLE_SHAPES:
                                modes = (('same_header',) if shape == 'table1'
                                         else HEADER_MODES)
                                for mode in modes:
                                    out.append(CaseParams(
                                        type_name, shape, case.key, fmt,
                                        direction, seed, header_mode=mode))
    return out


PARAMS = _build_params()
IDS = [p.case_id for p in PARAMS]

# ── one generated case per case_id, shared by all six oracles (§11) ───────
_ROOT = os.path.join(tempfile.gettempdir(), f'grepxcel-oracle-{os.getpid()}')
_CASES: dict[str, object] = {}


def case_for(params: CaseParams):
    """Build (or reuse) the case for *params*. Generation happens once per
    case_id however many oracles ask for it."""
    m = _CASES.get(params.case_id)
    if m is None:
        m = build_case(params, os.path.join(_ROOT, params.case_id))
        _CASES[params.case_id] = m
    return m


# ── 10.1 validate-pattern ────────────────────────────────────────────────
@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_validate_pattern(params):
    m = case_for(params)
    with pytest.raises(SystemExit) as exc:
        cli.main(['validate-pattern', m.pattern_path, '-q'])
    assert exc.value.code == 0, (
        f'{params.case_id}: validate-pattern exited {exc.value.code}'
    )


# ── 10.2 lint ────────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_lint(params):
    oracles.check_lint(case_for(params))


# ── 10.3 profile ─────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_profile(params):
    oracles.check_profile(case_for(params))


# ── 10.4 extract ─────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_extract(params):
    oracles.check_extract(case_for(params))


# ── 10.5 schema ──────────────────────────────────────────────────────────
@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_schema(params):
    oracles.check_schema(case_for(params))


# ── 10.6 strict exit code ────────────────────────────────────────────────
@pytest.mark.parametrize('params', PARAMS, ids=IDS)
def test_strict(params):
    oracles.check_strict(case_for(params))


# ── the matrix itself ────────────────────────────────────────────────────
def test_matrix_is_not_vacuous():
    """Guards against an env filter or a catalog edit silently emptying the
    run — a suite that collects nothing passes without testing anything."""
    assert len(PARAMS) >= 100, f'only {len(PARAMS)} cases in the matrix'
    assert len(set(IDS)) == len(IDS), 'duplicate case_id in the matrix'
