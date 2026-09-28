# TODO / decision sheet — schema strictness vs validate_type

**Status:** awaiting owner decision, per category. **Raised:** 2026-09-27.
**Nothing in this document has been implemented.** It exists so the decision can
be made consciously, case by case, rather than inferred from one example.

## The question

`grepxcel extract` and `grepxcel schema` read the *same* pattern file. When
`validate_type` rejects a value but the generated JSON Schema accepts it (or vice
versa), the two commands disagree about what that pattern means.

The oracle suite surfaced this: of its 3696 cases, **1320 hold a deliberately
invalid value**. For those, `oracles.check_schema` currently **skips** rather than
asserting, because conformance is not an invariant when the value is known-bad
(the engine writes rejected values out regardless). The skip is honest but it is
also 1320 cases of lost coverage.

Splitting them by what the schema actually does gives two groups — and the second
group is where the real decisions are.

## Group A — schema already REJECTS (792 cases): no product decision needed

These are pure test-coverage wins. The schema catches the bad value, so the
oracle can assert that instead of skipping.

| type | storage case | cases | example | validate_type reason |
|---|---|---|---|---|
| integer | n-float | 198 | `-0.01` | is not a whole number |
| integer | s-text | 66 | `'0'` | is not integer type |
| integer | b-bool | 66 | `False` | boolean is not integer |
| number | s-text | 66 | `'0'` | is not numeric |
| number | b-bool | 66 | `False` | boolean is not number |
| currency | s-text | 66 | `'0.00'` | is not numeric |
| currency | b-bool | 66 | `False` | boolean is not currency |
| percentage | s-text | 66 | `'1'` | is not numeric |
| percentage | b-bool | 66 | `False` | boolean is not percentage |
| url | n-int | 66 | `42` | does not look like a URL (no scheme) |

**Proposed:** assert `jsonschema.ValidationError` for these. Test-only change.
→ **Decision: ______**

## Group B — schema ACCEPTS a value extraction rejects (528 cases)

Each row is a separate product decision. "Risk" is what a user actually loses if
the schema stays loose: they validate their pipeline output against the generated
schema, it passes, and the data is still wrong.

### B1 — `url`, 132 cases — RECOMMEND TIGHTENING

| storage case | cases | example | validate_type reason | schema today |
|---|---|---|---|---|
| s-javascript | 66 | `'javascript:alert(1)'` | javascript: URLs are not permitted | accepts |
| s-noscheme | 66 | `'foo.org/a/b'` | does not look like a URL (no scheme) | accepts |

`url` is not in `schema._TYPE_MAP` at all, so it falls through to the default
`{'type': ['string','null']}` — any string passes. Note the first row is the
security-relevant one: `validate_type` explicitly refuses `javascript:` URLs and
the schema does not.

Proposed:
```python
'url': {'type': ['string', 'null'],
        'pattern': '^(https?|ftps?|mailto|tel|sms|file):'}
```
Keeps the allow-list in one place conceptually, but now duplicated between
`utils._ALLOWED_URL_SCHEMES` and `schema.py` — a cross-reference consistency risk
worth a test that derives the pattern from the constant rather than repeating it.

→ **Decision: ______**

### B2 — `boolean`, 132 cases — A TRADE-OFF I INTRODUCED TODAY

| storage case | cases | example | validate_type reason | schema today |
|---|---|---|---|---|
| n-two | 66 | `-1` | is not a boolean | accepts |
| s-text | 66 | `'maybe'` | is not a boolean | accepts |

Before 2026-09-27 the schema said `['boolean','null']`, which wrongly rejected
`1`/`0`/`'yes'` — values extraction accepts. I widened it to
`['boolean','integer','string','null']`, which fixed that but now accepts `-1`
and `'maybe'`. **A false rejection was traded for a false acceptance.**

Proposed (the tighter form, matching `validate_type` exactly):
```python
'boolean': {'anyOf': [
    {'type': ['boolean', 'null']},
    {'type': 'integer', 'enum': [0, 1]},
    {'type': 'string', 'pattern': '(?i)^\\s*(true|false|yes|no|1|0)\\s*$'},
]}
```
The `\s*` and `(?i)` are load-bearing: `trim.whitespace` and `ignore.case` mean
`' Yes '` is legitimately valid, so a plain `enum` would reintroduce false
rejections.

→ **Decision: ______**

### B3 — `date` / `datetime` / `time` / `duration` as text, 264 cases — SUPERSEDED

| storage case | cases | example | validate_type reason | schema today |
|---|---|---|---|---|
| date s-text | 66 | `'1900-01-01'` | is not a date | accepts |
| datetime s-text | 66 | `'1900-01-01 00:00'` | is not a datetime | accepts |
| time s-text | 66 | `'00:00'` | is not a time | accepts |
| duration s-text | 66 | `'0:00'` | is not a time | accepts |

The schema accepts these because `date`/`datetime` declare
`{'type':['string','null'], 'format':'date-time'}` and `time`/`duration` declare
`{'type':['string','null']}` — an ISO string satisfies `type: string`, and
`jsonschema` does not enforce `format` unless a format checker is installed.

**These 264 stop being "invalid" once temporal text coercion lands** (approved
separately, same day): an ISO date string becomes a real `datetime`, so
extraction accepts it and the disagreement disappears. No decision needed here
beyond confirming that work covers it.

Open sub-question once coercion exists: should `schema` install a
`FormatChecker` so `format: date-time` is actually enforced? That would make the
schema meaningfully stricter for every date field, not just these cases.

→ **Decision: ______**

## Cross-cutting question

Should `schema.py` derive its constraints from `utils.validate_type`'s own
constants (`_ALLOWED_URL_SCHEMES`, the boolean string set) instead of restating
them? Restating is how the boolean pair drifted apart in the first place —
`validate_type` was widened in ea9afbe and `_TYPE_MAP` was not. A single shared
source plus a test asserting the two agree would prevent the next drift.

→ **Decision: ______**

## How to replay the evidence

```bash
.venv/bin/python3 tmp.local/analyse_schema_skips.py
```
Prints the table above from the running code: every skipped category, its case
count, an example value, `validate_type`'s reason, and whether the generated
schema accepts or rejects it.
