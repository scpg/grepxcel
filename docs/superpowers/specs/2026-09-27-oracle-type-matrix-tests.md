# Oracle type-matrix test suite — implementation spec

**Status:** approved design, not yet implemented. **Owner:** SCPG. **Written:** 2026-09-27.
**Audience:** the implementer is expected to be a smaller/cheaper model working from this
document alone, without the design conversation. Every decision below is already made —
do not re-open them; if something here is genuinely impossible, stop and report it rather
than improvising a different design.

## 0. Read this first

- Read `CLAUDE.md` at the repo root and obey it: no `python3 -c`, every script in
  `scripts/` or `tmp.local/` starts with the venv bootstrap block, tests call the Python
  API (`cli.main([...])`, `Engine().process(...)`) — never `subprocess` to the `grepxcel`
  binary (CI does not install the package).
- Read `docs/pattern-file.md` (pattern syntax) and `grepxcel/utils.py::validate_type`
  (the exact acceptance rules the oracles encode) before writing any code.
- Existing test infrastructure to REUSE, not duplicate:
  `tests/unit/test_random_type_fixtures.py` (`TYPE_CASES`, `RandomPlacer`,
  `scan_index`/`position_before`, `write_pattern_csv`, `write_data_xlsx`),
  `tests/integration/test_schema_validation.py` (`_JSONEncoder`, the jsonschema
  validation pattern), `tests/conftest.py`.
- Work on a feature branch (`feature/oracle-type-matrix/<github-user>` per
  `global/git_branch_naming`), never on `main`.

## 0.5 Context for the implementer — how we got here and what mindset to bring

You are joining a long collaboration between the project owner and Claude. This section
is the part of that conversation you need; the rest of this document is the design that
came out of it.

**The owner's testing mindset (this is the thing to internalise).** The owner has driven
grepxcel's testing forward for months and, more than once, has found that a test "that
should have been there since the beginning" was missing. Their view, in their words: a
tool that fails without proper messaging wastes the user's time, and grepxcel's limits
exist to protect against hidden or malicious content in Excel files (ReDoS, ZIP bombs,
oversized sheets) — so a misleading or silent failure defeats the guard almost as badly
as no guard at all. The persistent checklist distilled from this lives in the corkern
memory `project_testing_checklist.md`; the seven items are, in short:
1. boundary values (grid corners, zero/min/max, and the exact value of any code constant
   — test N-1, N, N+1);
2. **error-message root-cause accuracy** — distinct causes must produce distinctly-worded
   messages; build an input that would have *passed* if a guard hadn't fired, and assert
   the message names that guard;
3. every computed diagnostic field must be traced to an actual output path (a flag that is
   set but never printed is a bug);
4. cross-reference consistency wherever a behaviour is described in more than one place;
5. say explicitly what is out of scope instead of letting it be silently absent;
6. verify against the running code — never reason from a docstring or from memory;
7. security-guard limits deserve clearer wording than performance caps.

This suite is items 2, 3 and 6 turned into machinery: several independent code paths are
forced to agree with generated ground truth, and rejected cases assert the *reason text*.

**The real bugs that motivated this design (all found by hand, this session, all now
fixed on the branch).** Each one maps to a design element below — that is why the
element is there:
- `profile` reported the three `IMAGE()` cells in fixture 24 as Excel `#VALUE!` errors.
  Root cause: the classifier only saw the cached scalar; the richData chain that says
  "this is an image" lived in another code path (`engine.scan_image_cells`) and was never
  consulted. → §10.3 profile-as-oracle: generator truth vs classifier must agree cell by cell.
- A `table:1` block silently overran into an unrelated table further down the sheet that
  happened to use the same header text (fixture 19), duplicating a row into the wrong
  group. Multiplicity had been parsed for years but never enforced. → §9 `distinct_header`
  mode for `table2`/`tableN` is the regression net for exactly this.
- `profile_workbook()` computed `SheetProfile.truncated` correctly and no output path ever
  printed it; the existing test asserted the flag's value, not its appearance in output.
  → §10.3 asserts `not truncated` on rendered results, and the general rule: assert
  rendered output, not internal fields.
- A 1500-char string that perfectly matched its regex failed with "does not match" — the
  real cause (over `--max-cell-len`, never evaluated) was collapsed into the generic
  message. Same for regex timeout. → §10.4 asserts the reason substring, never just
  "a warning happened".
- `profile` crashed on any workbook containing a chart sheet (`wb[name]` returns a
  `Chartsheet`, which has no `max_row`). → §10.2/§10.3 run `lint`/`profile` on every
  generated file so structural surprises surface immediately.

**Pitfalls already hit once — do not repeat them.**
- `cell:next` scans *forward*; a pattern whose instruction order does not match the
  sheet's scan order fails with "sheet is exhausted". An earlier attempt to fix this with a
  `seek:sheet` instruction was invalid syntax. The fix is §9's sort-by-`scan_index` rule.
- A hypothesis text strategy produced an unassigned Unicode codepoint (`'퟼'`) that
  `utils.is_empty()` treats as invisible, making a real cell look empty. That is why §7
  restricts strings to printable categories and why this suite uses curated pools, not
  arbitrary generation.
- The engine writes a `var:` value into the output **even when validation rejects it**
  (`engine._process_cell` is unconditional); rejection is reported as a warning. Tests
  that assumed a rejected value becomes `null` were wrong. §10.4 encodes the real behaviour.
- `openpyxl` cannot write `inlineStr`, cannot produce a cached value for a formula, and
  cannot create richData `IMAGE()` cells (§12). Do not spend time trying.
- Running `python3 -c "..."` is banned in this repo (`CLAUDE.md`, no exceptions) and was
  slipped into twice this session by accident. Write a file under `tmp.local/`.
- Full-suite console output in this environment sometimes truncates before the summary
  line; when in doubt, pass `--junit-xml=<file>` and read the `<testsuite ...>` line for
  the authoritative counts.

**Existing test suites this complements (do not duplicate them):**
`tests/unit/test_random_type_fixtures.py` (fixed-seed, 166 tests: 1/2 labels, 1/2
mini-tables, combined label-before-table under LR+TD, grid-corner and per-type boundary
values, `table:{n,m}` enforcement), `tests/unit/test_type_pattern_hypothesis.py`
(hypothesis-driven positions/values for the two simplest shapes), `tests/unit/test_profile.py`,
`tests/unit/test_lint.py`, `tests/integration/test_schema_validation.py`. This suite's
distinct contributions are the m=3..10 random-count shapes, the storage-type axis with
negative cases, and running all five subcommands against one manifest.

**Where the collaboration's memory lives:** corkern (`~/dev/corkern/memory/grepxcel/`),
read automatically at session start; the files most relevant here are
`project_testing_checklist.md`, `feedback_message_accuracy_and_signal_wiring.md`,
`feedback_pre_pr_oracle_gate.md`, `project_oracle_type_matrix.md`.

## 1. Purpose

A "reverse white-box" suite: for every combination of **value type × file shape × OOXML
storage type × number format × read direction**, the suite *knows* what it generated
(a manifest), builds a pattern file that must recognise it, builds the matching Excel
file, then runs **every relevant grepxcel subcommand** against them and checks each one's
output against the manifest. The manifest is the single oracle; `validate-pattern`,
`lint`, `profile`, `extract` and `schema` are five independent code paths that must all
agree with it.

Why it exists (the user's framing): a tool's failure messaging is as load-bearing as its
success behaviour, and several real bugs this project found by hand (image cells reported
as `#VALUE!` errors, `table:1` overrunning into an unrelated table with the same header,
a computed `truncated` flag never printed) would all have been caught automatically by a
suite that forces independent code paths to agree with generated ground truth.

## 2. Decisions already made (do not re-litigate)

| Decision | Choice |
|---|---|
| Runs in CI on every push? | **No.** Too slow for the per-push loop; the normal `pytest tests/` stays fast. |
| When does it run? | **Before every PR, as a gate.** Must be green before `gh pr create`. Enforced by: `scripts/oracle_check.py` (the command), the Branch-workflow step in `CLAUDE.md`, a surfacing corkern memory, and a corkern reminder. |
| CI-on-PR-only job? | **Deferred, optional** (Phase 6). A `pull_request`-triggered job running `pytest -m oracle` is low complexity, but not built now. |
| Randomness | **Fixed seeds only** (`GREPXCEL_ORACLE_SEEDS`, default `3`). No hypothesis in this suite — a failure must be replayable from the seed printed in the failure message. |
| Storage/format loops | **Type-conditional**, never a blind cross product (§5, §6). Each pair has a defined expected outcome, including the negative ones. |
| Pattern-side variants (modes, modifiers, `ignore.case`) | **Fixed to defaults in v1.** Recorded as a v2 axis (§11). |
| Types excluded from extraction | `image` (not generatable by openpyxl), `error` (profile/lint only), `empty` (implicit). |
| Grid | 30 rows × 30 cols, all elements separated by ≥1 empty row AND column (margin), never overlapping. |
| Value source | Curated deterministic pools per type (§7). No arbitrary string generation. |

## 3. File layout to create

```
tests/oracle/__init__.py                 # empty; makes `from tests.oracle import ...` work
tests/oracle/manifest.py                 # dataclasses: CellSpec, TableSpec, Manifest
tests/oracle/catalog.py                  # TYPES, STORAGE_MATRIX, FORMATS, VALUE_POOLS (§4–§7)
tests/oracle/generator.py                # build_case(...) -> Manifest + pattern.csv + data.xlsx
tests/oracle/placement.py                # grid placement, margin/non-overlap, scan-order sort
tests/oracle/oracles.py                  # one function per subcommand, each takes a Manifest
tests/oracle/test_oracle_matrix.py       # the parametrized pytest entry point (marker: oracle)
scripts/oracle_check.py                  # the pre-PR gate command (venv bootstrap block)
```

`pytest.ini` changes (append; keep existing keys):
```
markers =
    oracle: slow generated type-matrix suite; excluded by default, run with -m oracle
addopts = -m "not oracle"
```
Verify after editing: `.venv/bin/pytest tests/ --collect-only -q | tail -1` must show the
same count as before (oracle tests excluded), and
`.venv/bin/pytest -m oracle tests/oracle --collect-only -q | tail -1` must show > 0
(`-m` given on the command line overrides the one in `addopts`).

## 4. Types in scope

`EXTRACTABLE_TYPES = ['string', 'integer', 'number', 'currency', 'percentage', 'date',
'datetime', 'time', 'duration', 'boolean', 'url']` (11).

`PROFILE_ONLY_TYPES = ['error']` — generated as the literal string `'#DIV/0!'` (and the
other 8 codes in `grepxcel/cell_taxonomy.py::_EXCEL_ERRORS`), checked by the `lint` and
`profile` oracles only; never targeted by a `var:` field.

`image` is out of scope: `IMAGE()`/richData cells cannot be written by openpyxl. Real
coverage stays in `tests/fixtures/24_type_tests` + `tests/unit/test_profile.py`.

## 5. Storage matrix — every (type, storage) pair with its expected outcome

"Storage" is how the generator writes the cell. Outcome for `extract` is what
`validate_type()` returns for a `var:` of that type with column D = `.*`
(for `url`, column D = `.*` too). **Reason substrings are quoted from
`grepxcel/utils.py::validate_type` — the oracle asserts the substring is present in the
warning message, so copy them exactly and re-check them against the current source.**

Key: `n` = numeric cell (int/float/date/time written as Python objects), `s` = plain string,
`b` = Python bool. `str`/`inlineStr` are NOT generatable (see §12) — omit.

| type | storage | generator writes | extract outcome | reason substring (if rejected) | profile expects (storage/semantic) |
|---|---|---|---|---|---|
| string | s | text from pool | accepted | | s/string |
| integer | n | int | accepted | | n/integer |
| integer | n | float with `.is_integer()` (e.g. `42.0`) | accepted | | n/integer |
| integer | n | non-integral float (`42.5`) | rejected | `is not a whole number` | n/number |
| integer | s | `"42"` | rejected | `is not integer type` | s/string (+flag `text_forced_numeric`) |
| integer | b | `True` | rejected | `boolean is not integer` | b/boolean |
| number | n | float | accepted | | n/number |
| number | n | int | accepted | | n/integer |
| number | s | `"42.5"` | rejected | `is not numeric` | s/string (+`text_forced_numeric`) |
| number | b | `False` | rejected | `boolean is not number` | b/boolean |
| currency | n | float, currency fmt | accepted | | n/currency |
| currency | s | `"12.50"` | rejected | `is not numeric` | s/string (+`text_forced_numeric`) |
| currency | b | `True` | rejected | `boolean is not currency` | b/boolean |
| percentage | n | float in [0,1] (and one >1), `%` fmt | accepted | | n/percentage |
| percentage | s | `"0.15"` | rejected | `is not numeric` | s/string (+`text_forced_numeric`) |
| percentage | b | `True` | rejected | `boolean is not percentage` | b/boolean |
| date | n | `datetime.date`, date fmt | accepted | | n/datetime (see note A) |
| date | s | `"2024-01-15"` | rejected | `is not a date` | s/string |
| datetime | n | `datetime.datetime`, datetime fmt | accepted | | n/datetime |
| datetime | s | `"2024-01-15 09:30"` | rejected | `is not a datetime` | s/string |
| time | n | `datetime.time`, time fmt | accepted | | n/time |
| time | n | `datetime.timedelta`, `[h]:mm` | accepted (both accepted for either) | | n/duration |
| time | s | `"09:30"` | rejected | `is not a time` | s/string |
| duration | n | `datetime.timedelta`, `[h]:mm` | accepted | | n/duration |
| duration | n | `datetime.time`, `HH:MM` | accepted | | n/time |
| duration | s | `"2:00"` | rejected | `is not a time` | s/string |
| boolean | b | `True`/`False` | accepted | | b/boolean |
| boolean | n | `0`/`1` | accepted | | n/integer (see note B) |
| boolean | n | `2` | rejected | `is not a boolean` | n/integer |
| boolean | s | `"TRUE"`/`"yes"`/`"0"` | accepted | | s/boolean for TRUE/FALSE/YES/NO; s/string+`text_forced_numeric` for `"0"`/`"1"` (note B) |
| boolean | s | `"banana"` | rejected | `is not a boolean` | s/string |
| url | s | `https://example.com/x` | accepted | | s/url |
| url | s | `example.com/x` (no scheme) | rejected | `does not look like a URL (no scheme)` | s/string |
| url | s | `javascript:alert(1)` | rejected | `javascript: URLs are not permitted` | s/string (note C) |
| url | n | `42` | rejected | `does not look like a URL (no scheme)` | n/integer |

Note A — openpyxl returns a `datetime.datetime` (midnight) for a cell written as a
`datetime.date`; `classify_value()` therefore reports `datetime`, not `date`. The manifest
records the *taxonomy's* expectation (`n/datetime`) separately from the *extract*
expectation (accepted as `var: date`). This is the reason the manifest has two independent
expectation fields (§8) — do not "fix" one to match the other.

Note B — `validate_type` accepts `0`/`1`/`"1"` as boolean, but `classify_value` classifies
them as integer / text-forced-numeric string. Same rule as note A: two oracles, two truths,
both recorded.

Note C — `classify_value` treats any string containing `://` as `url`; `javascript:alert(1)`
has no `://`, so profile says `s/string`. Verify against current code when implementing.

**Implementation rule:** encode this table as data in `tests/oracle/catalog.py`
(`STORAGE_MATRIX: dict[type, list[StorageCase]]` where `StorageCase` carries
`writer`, `accepted: bool`, `reason: str | None`, `profile_storage`, `profile_semantic`,
`profile_flags`). Before trusting any row, run it once through `validate_type()` and
`classify_value()` directly in a `tmp.local/` script and correct the table if the code
disagrees — the code is the authority, this table is the design intent.

## 6. Number formats per type (valid set; type-conditional)

Applied only to `n`-storage cases. `classify_value()` (`grepxcel/cell_taxonomy.py`) is the
authority for `profile` expectations; for `date`/`datetime`/`time`/`duration` the Python
object type wins regardless of format, for plain numbers the format decides.

| type | formats to iterate |
|---|---|
| string | `General`, `@` (keep pool strings non-numeric so `@` does not add `text_forced_numeric`) |
| integer | `General`, `0`, `#,##0` |
| number | `General`, `0.00`, `#,##0.00` |
| currency | `$#,##0.00`, `€#,##0.00`, `£#,##0.00` (symbol must be in `cell_taxonomy._CURRENCY_SYMBOLS`) |
| percentage | `0%`, `0.00%` |
| date | `yyyy-mm-dd`, `dd/mm/yyyy`, `d-mmm-yy` |
| datetime | `yyyy-mm-dd hh:mm`, `yyyy-mm-dd hh:mm:ss` |
| time | `HH:MM`, `HH:MM:SS` |
| duration | `[h]:mm`, `[h]:mm:ss`, `[mm]:ss` |
| boolean | `General` |
| url | `General` |

Known heuristic discrepancy, out of scope for this suite: `grepxcel/drafter.py::
_type_from_number_format` treats `#,##0.00` as currency while `classify_value` says number.
The oracle uses `classify_value` only. Do not change either.

**Cross-format re-typing cases (Phase 5, optional):** an `int` written with `[h]:mm` comes
back from openpyxl as a `timedelta`; a float with `%` becomes `percentage` to profile.
These are genuine "the format re-types the value" behaviours worth a small explicit list
(with expected outcomes) — but only after Phases 1–4 are green.

## 7. Value pools (deterministic, per type)

Start from `tests/unit/test_random_type_fixtures.py::TYPE_CASES` and `_BOUNDARY_VALUES`
(import them; do not copy). Rules:
- Every `string` value must be non-empty, non-whitespace, printable (categories L/N/P/S/Zs
  only — an unassigned codepoint was once treated as "empty" by `utils.is_empty()`), and
  must never equal a reserved label (`ANCHOR*`, `HEAD*`, `COLHEAD*`).
- Include per type: one ordinary value, one zero/min-style boundary, one max-style boundary
  (e.g. `datetime.date(1900,1,1)`/`(9999,12,31)`, `time(0,0)`/`time(23,59,59)`,
  `timedelta(0)`/`timedelta(days=999)`, string of exactly `_MAX_REGEX_INPUT_LEN` chars).
- The seed selects which pool entry each cell gets (`random.Random(seed).choice(pool)`).

## 8. Manifest (`tests/oracle/manifest.py`)

```python
@dataclass
class CellSpec:
    ref: str                     # 'B7'
    row: int; col: int
    field: str                   # var: field name, or None for anchors/headers/error cells
    role: str                    # 'anchor' | 'value' | 'header' | 'data' | 'error'
    written_value: object        # exactly what openpyxl was given
    number_format: str | None
    # extract expectations (validate_type truth)
    extract_accepted: bool | None   # None when the cell is not a var: target
    extract_reason: str | None      # substring, when rejected
    # profile expectations (classify_value truth)
    profile_storage: str
    profile_semantic: str
    profile_flags: frozenset[str]   # subset that MUST be present (e.g. {'text_forced_numeric'})

@dataclass
class TableSpec:
    group_key: str               # output key, e.g. 'table_0' or 'items' (see §9)
    header_ref: str
    columns: list[str]           # var: field names, in order
    data_rows: list[list[object]]
    expected_instances: int      # 1 for distinct-header mode; m for same-header mode

@dataclass
class Manifest:
    case_id: str                 # '{type}-{shape}-{storage}-{fmt}-{direction}-seed{n}'
    direction: str               # 'LR' | 'TD'
    pattern_path: str; data_path: str
    cells: list[CellSpec]
    tables: list[TableSpec]
    expected_extract: dict       # the exact nested dict Engine().process must return (normalised, §10.4)
```

The manifest is written next to the generated files as `manifest.json` (`default=str`) so a
failing case can be inspected by hand.

## 9. Shapes and placement (`tests/oracle/placement.py`, `generator.py`)

Six shapes, each parametrised by `direction ∈ {LR, TD}`:

| shape id | content | count m |
|---|---|---|
| `cells1` | 1 scalar | 1 |
| `cells2` | 2 scalars | 2 |
| `cellsN` | m scalars | m = `Random(seed).randint(3, 10)` |
| `table1` | 1 mini-table | 1 |
| `table2` | 2 mini-tables | 2 |
| `tableN` | m mini-tables | m = `Random(seed).randint(3, 10)` |

**Scalar = anchor + value pair**, addressed two ways; both are generated for every
scalar shape (parameter `via ∈ {next, abs}`):
- `via=next`: pattern rows `lbl: | anchor_i | string | ANCHOR_i` and
  `var: | field_i | <type> | .*`; sequence `cell:next anchor_i`, `cell:next field_i`.
  Value cell must be the **scan-order successor** of the anchor: `(r, c+1)` for LR,
  `(r+1, c)` for TD.
- `via=abs`: no anchor; sequence `cell:<ref> field_i`. Refs must be emitted in
  forward scan order (the engine enforces "ordering must be forward").

**Ordering rule (both modes):** generate random positions first, then **sort elements by
`scan_index(pos, direction, 30, 30)`** (import from `test_random_type_fixtures`) and emit
pattern instructions in that order. This is what makes "randomly located" work with a
forward-scanning engine. `read.direction` is written as `config: | read.direction | <dir>`.

**Mini-table:** `k ∈ {1,2,3}` columns (seeded), all of the same type; `HEADER:1` row with
`k` `lbl:` headers; `DATA:*` with `r ∈ {1..5}` (seeded) data rows; no FOOTER, no SKIP_IF
(table-row mechanics are a different suite's concern). Two header modes, both generated
for `table2`/`tableN`:
- `same_header`: every table uses identical header text `COLHEAD_1..k`; ONE `table:*`
  block; `expected_instances = m`, all under one group key.
- `distinct_header`: table *i* uses `HEAD{i}_1..k`; *m* separate `table:1` blocks emitted in
  scan order of their header cells; each `expected_instances = 1`.
  (This mode is the regression net for the fixture-19 class of bug — a `table:1`
  overrunning into a later table.)

Group key: name var fields `t{i}.col{j}` so the output key is `t{i}` (dotted-prefix
naming in `engine._table_group`); for `same_header` use `items.col{j}` → key `items`.

**Placement constraints (enforce, then assert in a unit test of `placement.py` itself):**
- Grid 30×30. Every block (a scalar pair or a `(r+1)×k` table rectangle) is separated from
  every other block by at least one fully empty row and one fully empty column. Reason: a
  `DATA:*` scan runs until an empty row; a table placed directly beneath another would be
  consumed by it.
- No overlap. Retry up to 2000 draws; raise with the seed if impossible.
- `cellsN`/`tableN` with m=10 must be provably placeable in 30×30 with these margins —
  add a unit test that seeds 0..19 all place successfully.

## 10. Oracles (`tests/oracle/oracles.py`) — one function each, all take a `Manifest`

10.1 **validate-pattern** — `cli.main(['validate-pattern', m.pattern_path, '-q'])`
inside `pytest.raises(SystemExit)`; assert exit code `0`.

10.2 **lint** — `from grepxcel.lint import lint_file, OK, WARN, FAIL, INFO`;
`results = lint_file(m.data_path)`; assert no result has severity `FAIL` or `WARN`. A
generated file has no merged cells, no formulas, no declared>used extent — if lint warns,
that is a real finding, not a test bug: report it.

10.3 **profile** — `from grepxcel.profile import profile_workbook`;
`sheets = profile_workbook(m.data_path)`; build `{ref: cell}`; for every `CellSpec`
assert `storage_type == profile_storage`, `semantic_type == profile_semantic`,
`profile_flags ⊆ flags`; assert the set of profiled refs equals the set of manifest refs
(no extra cells, none missing); assert `not sp.truncated` for every sheet.

10.4 **extract** — `Engine().process(pattern_file=..., data_file=..., logger=Logger(level=VerbosityLevel.QUIET))`.
Normalise both sides with the same function before comparing:
`datetime/date/time → .isoformat()`, `timedelta → str()`, everything else unchanged
(mirror `tests/integration/test_schema_validation._JSONEncoder`). Then:
- accepted scalars: `result[field] == expected` after normalisation;
- rejected scalars: the value is STILL present in `result[field]` (the engine writes it
  unconditionally — this is documented current behaviour, see
  `test_random_type_fixtures._run_with_issues`), AND at least one record in
  `logger._records` with severity `WARNING` or `ERROR` whose `message` contains
  `extract_reason`. **This is the message-accuracy check** — assert the reason text, not
  merely that a warning exists;
- tables: `len(result[group_key]) == expected_instances`; each instance's `data` rows equal
  the manifest rows after normalisation (match instances to `TableSpec`s by `_source.ref`).

10.5 **schema** — `from grepxcel.schema import generate_schema`;
`schema = generate_schema(m.pattern_path)`; `jsonschema.validate(json.loads(json.dumps(result, cls=_JSONEncoder)), schema)`
(reuse the encoder from `test_schema_validation.py`). `jsonschema` is already a dev
dependency — do not add anything to `pyproject.toml`.

10.6 **strict exit code (cheap extra)** — for cases with any rejected cell:
`cli.main(['extract', '-p', pattern, data, '--strict', '-q'])` → `SystemExit` code `2`;
for all-accepted cases → `0`.

## 11. Test entry point (`tests/oracle/test_oracle_matrix.py`)

- Module-level `pytestmark = pytest.mark.oracle`.
- Build the case list at import time from the catalog:
  `for direction × type × shape × storage_case × fmt × seed` with the type-conditional
  filters of §5/§6, and for scalar shapes × `via`, for multi-table shapes × `header_mode`.
  Use `pytest.mark.parametrize(..., ids=case_id)` so every case is individually named and
  a failure prints its `case_id` (which contains the seed).
- One test function per oracle (5–6 functions), each parametrized over the same case list,
  each building the case into `tmp_path` via `generator.build_case(...)`. Cache the built
  case per `case_id` in a module-level dict so the five oracles share one generation.
- Expected count with defaults (seeds=3): roughly
  `2 dir × 11 types × (3 scalar shapes × 2 via + 1 table shape + 2 table shapes × 2 modes)`
  `× ~3 storage cases × ~2.5 formats × 3 seeds` ≈ 4,000–6,000 parametrized test IDs across
  the oracle functions. Target runtime: **under 5 minutes** on the dev machine. If it is
  slower, first reduce `GREPXCEL_ORACLE_SEEDS` default to 2 and cache per-case builds;
  do not drop cases.
- `GREPXCEL_ORACLE_SEEDS` (int, default 3) and `GREPXCEL_ORACLE_TYPES` (comma list,
  default all) environment variables narrow the run for local debugging.

## 12. Known limitations — document in the module docstring, do not "solve"

- `str` (formula-cached string) and `inlineStr` storage cannot be produced by openpyxl
  (a written formula has no cached value under `data_only=True`; openpyxl always writes
  shared strings). Covered only by real fixtures (`24_type_tests` has 50 `str/string`
  cells).
- `image` cells: same — real fixtures only.
- Pattern-side variants (`var:literal`/`glob`, `lbl:regexp`, `not-null`, `trim-whitespace`,
  `nullable`, `ignore.case`) are fixed to defaults. Listed as the v2 axis.
- `FOOTER`, `SPLITTER`, `SKIP_IF`, `SKIP_EMPTY_ROW`, `DATA:{n,m}`, `table:{n,m}` are out
  of scope here (covered by their own unit tests).

## 13. The pre-PR gate (`scripts/oracle_check.py`)

Starts with the mandatory `scripts/` bootstrap block. Runs, in order, and stops at the
first failure with a clear message:
1. `.venv/bin/pytest tests/ -q` (the normal fast suite; must pass)
2. `.venv/bin/pytest -m oracle tests/oracle -q` (this suite; must pass)
Prints a final one-line verdict: `ORACLE GATE: GREEN — safe to open a PR` or
`ORACLE GATE: RED — do not open a PR`, exit code 0/1. Use `subprocess` for pytest here
(this is a developer script, not a CI test, so the CI-no-package-install rule does not
apply — but call `.venv/bin/pytest` by explicit path per `CLAUDE.md`).

`CLAUDE.md` → "Branch workflow" section: insert as step 0, before pushing:
`.venv/bin/python3 scripts/oracle_check.py   # must print GREEN before gh pr create`.
Also add one line under "Running tests" explaining the marker and that `pytest tests/`
excludes it by default.

## 14. Phases and acceptance criteria

**Phase 1 — catalog + manifest + placement (no oracles yet).**
Done when: `catalog.py` encodes §5/§6/§7 as data; a `tmp.local/` script has run every
storage row through `validate_type()` and `classify_value()` and the table in §5 has been
corrected to match the code (record any correction in the module docstring);
`placement.py` has unit tests proving non-overlap + margins + scan-order sort for LR and
TD and that m=10 places for seeds 0..19.

**Phase 2 — generator.** Done when `build_case()` produces `pattern.csv`, `data.xlsx`,
`manifest.json` for every shape/mode, and `validate-pattern` passes on every generated
pattern (oracle 10.1 only).

**Phase 3 — oracles 10.2–10.5 wired, `cells1`/`table1` only.** Done when green for all
types and storages on the two simplest shapes. Any red here is most likely a real
grepxcel bug or a wrong expectation in §5 — investigate; do not weaken the assertion.

**Phase 4 — full matrix + `scripts/oracle_check.py` + `CLAUDE.md` edit + pytest.ini
marker.** Done when the full run is green under 5 minutes and `pytest tests/` count is
unchanged from before this work.

**Phase 5 (optional)** — cross-format re-typing cases (§6). **Phase 6 (optional, decide
with the owner)** — a `pull_request`-only CI job.

Each phase: scoped tests per commit, full `pytest tests/` before the PR, signed commits
(`CLAUDE.md` signing command), one PR per phase or one PR for phases 1–4 — owner's call.

## 15. How to report back

For each phase: what was implemented, the exact command run and its last line of output,
every §5 row that had to be corrected and why, and every case that failed for a reason
that looks like a grepxcel bug rather than a test bug (with `case_id`). Never mark a
phase done with a red run.
