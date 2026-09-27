"""The five oracles — one per subcommand, each checked against the Manifest.

Each ``check_*`` takes a :class:`Manifest` and raises ``AssertionError`` with the
case_id (which carries the seed) on disagreement. The manifest is the single
ground truth; ``validate-pattern``, ``lint``, ``profile``, ``extract`` and
``schema`` are five independent code paths that must all agree with it.

Every assertion here is about **rendered output** — the list ``lint_file``
returns, the cells ``profile_workbook`` produces, the dict ``Engine().process``
returns, the records the logger actually holds — never an internal flag read
back out of the object that computed it. That rule exists because of a real bug:
``SheetProfile.truncated`` was computed correctly and printed by no output path,
and the test of the day asserted the flag's value rather than its appearance.

## The two defects this module found on its first full run

Both are fixed now; these 1518 cases are the regression net for them.

1. **§10.4, 1320 cases.** ``utils.validate_type`` computes a precise reason for
   every rejection and ``engine._validate_field`` discarded it
   (``ok, _ = validate_type(...)``), so every cause — a bool in an integer field,
   a ``javascript:`` URL, a cell over ``--max-cell-len``, a regex timeout, a
   genuine mismatch — surfaced as the same generic ``'Value does not match the
   expected pattern'``. ``engine._validate_field_with_reason()`` now carries it
   into ``warn_validation``'s ``message``.
2. **§10.5, 198 cases.** ``validate_type`` accepts ``1``/``0``, ``"1"``/``"0"``
   and ``"yes"``/``"no"`` as ``boolean``, while ``schema._TYPE_MAP`` declared
   only ``{"type": ["boolean", "null"]}`` — so the schema generated from a
   pattern rejected 3 of the 4 boolean forms that same pattern's extraction
   accepts. The declared types are now widened to match.
"""
from __future__ import annotations

import json

import jsonschema
import pytest

from grepxcel import cli
from grepxcel.engine import Engine
from grepxcel.lint import FAIL, WARN, lint_file
from grepxcel.logger import Logger, Severity, VerbosityLevel
from grepxcel.profile import profile_workbook
from grepxcel.schema import generate_schema
from tests.integration.test_schema_validation import _JSONEncoder
from tests.oracle.manifest import Manifest, normalise

#: §10.4's message-accuracy assertion. True = assert the reason substring, as
#: the spec requires. Never set this False to make a run green: a rejected value
#: whose warning does not say *why* is the exact failure mode this suite exists
#: to catch.
REASON_TEXT_REQUIRED = True

#: Cache keyed by case_id so the extract-consuming oracles share one engine run.
_EXTRACT_CACHE: dict[str, tuple] = {}


def run_extract(m: Manifest) -> tuple[dict, list]:
    """``Engine().process`` plus the WARNING/ERROR records it raised."""
    hit = _EXTRACT_CACHE.get(m.case_id)
    if hit is not None:
        return hit
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=m.pattern_path, data_file=m.data_path,
                              logger=lg)
    issues = [r for r in lg._records
              if r.severity in (Severity.WARNING, Severity.ERROR)]
    _EXTRACT_CACHE[m.case_id] = (result, issues)
    return result, issues


def _record_text(rec) -> str:
    """Everything a user would actually see for one record."""
    parts = [getattr(rec, attr, None) for attr in
             ('message', 'hint', 'expected', 'found', 'event')]
    return ' | '.join(str(p) for p in parts if p)


# ── 10.2 lint ─────────────────────────────────────────────────────────────
def check_lint(m: Manifest) -> None:
    """A generated file has no merged cells, no formulas and no declared>used
    extent, so lint must be clean. A WARN or FAIL here is a real finding."""
    results = lint_file(m.data_path)
    bad = [r for r in results if r[0] in (FAIL, WARN)]
    assert not bad, (
        f'{m.case_id}: lint reported {len(bad)} WARN/FAIL result(s) on a '
        f'generated file:\n' + '\n'.join(f'  {sev}: {cat}: {msg}'
                                        for sev, cat, msg in bad)
    )


# ── 10.3 profile ──────────────────────────────────────────────────────────
def check_profile(m: Manifest) -> None:
    """Cell by cell against the manifest, plus set equality of the refs."""
    sheets = profile_workbook(m.data_path)

    for sp in sheets:
        assert not sp.truncated, (
            f'{m.case_id}: profile marked sheet {sp.sheet!r} truncated for a '
            f'30x30 generated file'
        )

    by_ref = {}
    for sp in sheets:
        for cell in sp.cells:
            by_ref[cell.ref] = cell

    assert set(by_ref) == set(m.refs), (
        f'{m.case_id}: profiled ref set != manifest ref set — '
        f'only profiled: {sorted(set(by_ref) - set(m.refs))}, '
        f'only in manifest: {sorted(set(m.refs) - set(by_ref))}'
    )

    for spec in m.cells:
        prof = by_ref[spec.ref].profile
        assert prof.storage_type == spec.profile_storage, (
            f'{m.case_id} {spec.ref} (role={spec.role}, '
            f'wrote {spec.written_value!r:.40} fmt={spec.number_format!r}): '
            f'storage {prof.storage_type!r}, manifest expects '
            f'{spec.profile_storage!r}'
        )
        assert prof.semantic_type == spec.profile_semantic, (
            f'{m.case_id} {spec.ref} (role={spec.role}, '
            f'wrote {spec.written_value!r:.40} fmt={spec.number_format!r}): '
            f'semantic {prof.semantic_type!r}, manifest expects '
            f'{spec.profile_semantic!r}'
        )
        missing = set(spec.profile_flags) - set(prof.flags)
        assert not missing, (
            f'{m.case_id} {spec.ref}: profile flags {prof.flags} are missing '
            f'{sorted(missing)}'
        )


# ── 10.4 extract ──────────────────────────────────────────────────────────
def check_extract(m: Manifest) -> None:
    """Accepted values must come back equal; rejected values must come back
    ANYWAY (the engine writes unconditionally) *and* be reported with a warning
    that names the reason."""
    result, issues = run_extract(m)
    norm = normalise(result)

    # --- scalars -------------------------------------------------------
    for spec in m.cells:
        if spec.role != 'value':
            continue
        assert spec.field in norm, (
            f'{m.case_id}: field {spec.field!r} missing from extract output '
            f'(keys: {sorted(norm)})'
        )
        expected = m.expected_extract[spec.field]
        assert norm[spec.field] == expected, (
            f'{m.case_id} {spec.ref} field {spec.field!r}: extracted '
            f'{norm[spec.field]!r:.60}, manifest expects {expected!r:.60}'
        )
        if spec.extract_accepted is False:
            _assert_rejection_reported(m, spec, issues)

    # --- tables --------------------------------------------------------
    _check_tables(m, result, issues)


def _assert_rejection_reported(m: Manifest, spec, issues) -> None:
    """§10.4's message-accuracy check: the reason text, not merely a warning."""
    at_cell = [r for r in issues if spec.ref in str(getattr(r, 'location', ''))]
    assert at_cell, (
        f'{m.case_id} {spec.ref} field {spec.field!r}: value '
        f'{spec.written_value!r:.40} is rejected by validate_type '
        f'({spec.extract_reason!r}) but NO warning/error was raised for that '
        f'cell. Records: {[_record_text(r) for r in issues]}'
    )
    if not REASON_TEXT_REQUIRED:
        return
    assert any(spec.extract_reason in str(getattr(r, 'message', '') or '')
               for r in at_cell), (
        f'{m.case_id} {spec.ref} field {spec.field!r} '
        f'(type={spec.profile_semantic}, wrote {spec.written_value!r:.40}): '
        f'a warning was raised but its message does not name the reason.\n'
        f'  validate_type reason : {spec.extract_reason!r}\n'
        f'  record messages      : '
        f'{[str(getattr(r, "message", "")) for r in at_cell]}\n'
        f'  full records         : {[_record_text(r) for r in at_cell]}\n'
        f'  This is the message-accuracy check (§10.4): distinct causes must '
        f'produce distinctly-worded messages.'
    )


def _instance_start_ref(instance: dict) -> str:
    """The top-left ref of a table instance's _source range ('Y13:Z18' -> 'Y13')."""
    ref = (instance.get('_source') or {}).get('ref', '')
    return ref.split(':', 1)[0]


def _check_tables(m: Manifest, result: dict, issues) -> None:
    by_group: dict[str, list] = {}
    for spec in m.tables:
        by_group.setdefault(spec.group_key, []).append(spec)

    for group_key, specs in by_group.items():
        expected_instances = specs[0].expected_instances
        assert group_key in result, (
            f'{m.case_id}: table group {group_key!r} missing from extract '
            f'output (keys: {sorted(result)})'
        )
        instances = result[group_key]
        assert len(instances) == expected_instances, (
            f'{m.case_id}: table group {group_key!r} produced '
            f'{len(instances)} instance(s), manifest expects '
            f'{expected_instances} '
            f'(instance refs: {[_instance_start_ref(i) for i in instances]}, '
            f'manifest header refs: {[s.header_ref for s in specs]})'
        )

        want = {s.header_ref: s for s in specs}
        for instance in instances:
            start = _instance_start_ref(instance)
            assert start in want, (
                f'{m.case_id}: table instance at {start!r} matches no manifest '
                f'header ref {sorted(want)}'
            )
            spec = want[start]
            got_rows = normalise(instance.get('data') or [])
            exp_rows = [
                {name.partition('.')[2] or name: value
                 for name, value in zip(spec.columns, row)}
                for row in spec.data_rows
            ]
            assert got_rows == exp_rows, (
                f'{m.case_id}: table {group_key!r} instance at {start}: '
                f'{len(got_rows)} row(s) extracted vs {len(exp_rows)} expected\n'
                f'  extracted: {got_rows!r:.400}\n'
                f'  expected : {exp_rows!r:.400}'
            )

    # rejected table data cells must also be reported, per cell
    for spec in m.cells:
        if spec.role == 'data' and spec.extract_accepted is False:
            _assert_rejection_reported(m, spec, issues)


# ── 10.5 schema ───────────────────────────────────────────────────────────
def check_schema(m: Manifest) -> None:
    """The extraction must validate against the schema generated from its own
    pattern — two independent readings of the same pattern file.

    **Only for cases whose every value is accepted.** A rejected value is still
    written to the output (``engine._process_cell`` is unconditional), so a case
    that deliberately puts ``-0.01`` in an ``integer`` field produces output that
    its own schema correctly refuses — the schema is describing *valid* output
    and is right to reject. That is already this project's stated position:
    ``tests/integration/test_schema_validation._XFAIL_DATA_QUALITY`` xfails the
    fixtures with data-quality issues precisely because "the schema correctly
    rejects these". §10.5 does not say which way to read it, so rather than
    assert an invariant that does not hold, those cases are skipped **visibly**
    — they show up in the run's skip count with this reason instead of quietly
    passing. Note it is not even a uniform rejection: a string in a ``date``
    field satisfies the schema's ``{"type": "string"}`` and only the unvalidated
    ``format`` keyword would catch it, so "rejected ⇒ schema fails" would be
    just as wrong an assertion.
    """
    if m.has_rejection:
        pytest.skip(
            f'{m.case_id}: case holds {len(m.rejected_cells())} value(s) that '
            f'validate_type rejects; the generated schema describes valid '
            f'output, so conformance is not an invariant here'
        )
    result, _ = run_extract(m)
    schema = generate_schema(m.pattern_path)
    instance = json.loads(json.dumps(result, cls=_JSONEncoder))
    try:
        jsonschema.validate(instance=instance, schema=schema)
    except jsonschema.ValidationError as exc:
        raise AssertionError(
            f'{m.case_id}: extract output does not validate against the schema '
            f'generated from its own pattern\n'
            f'  path   : {list(exc.absolute_path)}\n'
            f'  message: {exc.message}'
        ) from exc


# ── 10.6 strict exit code ─────────────────────────────────────────────────
def check_strict(m: Manifest) -> None:
    """--strict exits 2 when any field had an issue, 0 when none did."""
    expected_code = 2 if m.has_rejection else 0
    try:
        cli.main(['extract', '-p', m.pattern_path, m.data_path, '--strict', '-q'])
        code = 0
    except SystemExit as exc:
        code = exc.code if exc.code is not None else 0
    assert code == expected_code, (
        f'{m.case_id}: extract --strict exited {code}, expected '
        f'{expected_code} ('
        f'{len(m.rejected_cells())} rejected cell(s) in the manifest)'
    )
