# Implementer prompt — oracle type-matrix suite

Copy everything below the line into the FIRST message of a fresh Claude Code session
opened on the grepxcel repo. Companion to
`2026-09-27-oracle-type-matrix-tests.md` (the spec).

---

You are implementing a pre-designed test suite for the grepxcel project. The design
is FINAL. Your job is to implement it faithfully, in parallel where the dependency
graph allows, and to test every piece before integrating it. Do not redesign; if
something in the spec is genuinely impossible, STOP and report it — do not improvise.

## Step 1 — read, in this exact order, before writing any code
1. CLAUDE.md (repo root). Obey it without exception. In particular: never run
   `python3 -c "..."` — every throwaway script goes in tmp.local/ with the venv
   bootstrap block; tests call the Python API (cli.main([...]), Engine().process(...)),
   never subprocess to the grepxcel binary; commits are signed with the exact command
   in CLAUDE.md.
2. docs/superpowers/specs/2026-09-27-oracle-type-matrix-tests.md — the spec. Read ALL
   of it. Section 0.5 ("Context for the implementer") explains why the suite exists,
   the real bugs it must catch, and six pitfalls already hit once — read that section
   twice.
3. grepxcel/utils.py::validate_type and grepxcel/cell_taxonomy.py::classify_value —
   the two functions whose behaviour the spec's §5 table encodes. The CODE is the
   authority; the table is design intent. You will verify the table against them.
4. tests/unit/test_random_type_fixtures.py — reuse TYPE_CASES, _BOUNDARY_VALUES,
   RandomPlacer, scan_index, position_before, write_pattern_csv, write_data_xlsx.
   Do not copy them.
5. tests/integration/test_schema_validation.py — reuse its _JSONEncoder and the
   jsonschema validation pattern.
6. docs/pattern-file.md — pattern syntax reference.

## Step 2 — setup
- Create branch: feature/oracle-type-matrix/<your-github-user>. Never touch main.
- Run `.venv/bin/pytest tests/ -q --junit-xml=tmp.local/baseline.xml` and record the
  test count from the <testsuite> line. This number must be UNCHANGED at the end
  (the oracle suite is excluded from the default run by marker).
- Do the pytest.ini marker change (spec §3) FIRST, so nothing you add can leak into
  the fast suite. Verify with --collect-only as §3 instructs.

## Step 3 — implement in parallel waves (each box is independent within its wave;
   use parallel subagents/workers for a wave if your environment supports it)

WAVE 1 (three independent tasks — run them in parallel):
  1a. tests/oracle/catalog.py — encode spec §5, §6, §7 as data. THEN write
      tmp.local/verify_catalog.py that runs EVERY §5 row through validate_type() and
      classify_value() and prints any row where the code disagrees with the table.
      Correct the table to match the code and note each correction in the module
      docstring. This task is not done until that script prints zero disagreements.
  1b. tests/oracle/placement.py — spec §9 placement: 30x30 grid, non-overlap,
      >=1 empty row AND column between every block, scan-order sort for LR and TD,
      value cell as scan-order successor of its anchor. Write
      tests/oracle/test_placement.py proving: no overlap, margins hold, sort order
      matches scan_index for both directions, and m=10 scalars AND m=10 tables place
      successfully for seeds 0..19. Mark these tests with the oracle marker too.
  1c. tests/oracle/manifest.py — the three dataclasses in spec §8 exactly as
      written, plus to_json()/from_json() (default=str) so a case can be inspected
      by hand.

WAVE 2 (depends on wave 1):
  2.  tests/oracle/generator.py — build_case(...) producing pattern.csv, data.xlsx,
      manifest.json for every shape/mode/via in §9. Acceptance: for every generated
      pattern, cli.main(['validate-pattern', path, '-q']) exits 0 (spec §10.1).
      Write tests/oracle/test_generator.py covering every shape x direction x via x
      header_mode at seed 0.

WAVE 3 (depends on wave 2; the five oracles are independent of each other — parallel):
  3a. oracles.py::check_lint      (§10.2)
  3b. oracles.py::check_profile   (§10.3) — cell-by-cell against the manifest, and
      assert the profiled ref set EQUALS the manifest ref set, and not truncated.
  3c. oracles.py::check_extract   (§10.4) — accepted: value equality after
      normalisation; rejected: value STILL present AND a WARNING/ERROR record whose
      message contains the exact reason substring. Never weaken this to "a warning
      exists".
  3d. oracles.py::check_schema    (§10.5) — jsonschema.validate of the normalised
      extract output against generate_schema(pattern).
  3e. oracles.py::check_strict    (§10.6)
  Each oracle gets its own small unit test on one hand-built manifest before it is
  wired into the matrix.

WAVE 4 (depends on wave 3):
  4a. tests/oracle/test_oracle_matrix.py — spec §11. First run it restricted to
      shapes cells1 + table1 only (spec Phase 3). Every red there is either a real
      grepxcel bug or a wrong §5 expectation — investigate and report; do not
      weaken assertions. Only then enable the full matrix.
  4b. scripts/oracle_check.py (§13) and the CLAUDE.md edit (§13) — independent of
      4a; can be done in parallel with it.

## Step 4 — verification you must run and report
- `.venv/bin/pytest tests/ -q --junit-xml=tmp.local/fast.xml` → count equals the
  baseline from Step 2.
- `.venv/bin/pytest -m oracle tests/oracle -q --junit-xml=tmp.local/oracle.xml`
  → 0 failures, 0 errors, wall time under 5 minutes (report the time). If slower,
  reduce GREPXCEL_ORACLE_SEEDS default to 2 and cache per-case builds — do NOT drop
  cases.
- `.venv/bin/python3 scripts/oracle_check.py` → prints `ORACLE GATE: GREEN`.
- Console output in this environment sometimes truncates before pytest's summary
  line: ALWAYS read the <testsuite ...> line of the junit XML for the authoritative
  numbers.

## Step 5 — commits and report
- One signed commit per wave (CLAUDE.md signing command), scoped tests run before
  each commit, full `pytest tests/` before the final push. Do not open the PR
  yourself; leave the branch pushed and report.
- Report exactly per spec §15: per wave, what was built, the exact commands run
  with their last authoritative line, every §5 row you had to correct and why, and
  every case that failed for a reason that looks like a grepxcel bug (with its
  case_id, which contains the seed). Never mark a wave done with a red run.
