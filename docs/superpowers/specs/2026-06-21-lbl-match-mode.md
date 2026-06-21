# Spec: lbl: match mode — literal default + per-field override

**Date:** 2026-06-21  
**Status:** Implemented (2026-06-21) — all phases complete  
**Motivation:** `lbl:` is a structural assertion ("is this the sheet I expect?"), not a data query. Regex was never the right default there — it leaked implementation details into user patterns and caused silent failures when literal header text contained metacharacters like `(`, `)`, `.`.

---

## What changes

### 1. Global default config key

```
config: | lbl.match | literal     ← NEW DEFAULT (no config line needed)
config: | lbl.match | glob        ← wildcards: * = any text, ? = one char
config: | lbl.match | regexp      ← current full-regex behaviour (opt-in)
```

When no `config: lbl.match` line is present, `literal` is assumed.

### 2. Per-field override in column A

| Column A    | Meaning |
|-------------|---------|
| `lbl:`      | Uses global `config: lbl.match` (defaults to `literal`) |
| `lbl:literal` | Exact string match for this field only |
| `lbl:glob`  | Wildcard match for this field only (`*` = any text, `?` = one char) |
| `lbl:regexp` | Full Python regex for this field only |

**Precedence:** field-level qualifier always overrides global config.

### 3. Modes explained

**`literal`**  
Plain string equality. Metacharacters are treated as literal characters — `Table (reservations)` matches the cell text `Table (reservations)` with no escaping needed. Honours `config: ignore.case`.

**`glob`**  
Shell-style wildcard matching via `fnmatch.translate()` with `re.DOTALL` so `*` matches across newlines (Alt+Enter cells). `*` = any sequence of characters (including empty), `?` = exactly one character. Example: `Table (*)` matches `Table (reservations)` and `Table (orders)`. Honours `ignore.case`.

**`regexp`**  
Current behaviour — full Python `re.search` against the cell value. Metacharacters must be escaped manually. Use this for power cases like `Amount.*EUR` or when migrating an existing pattern without rewriting `lbl:` values.

### 4. Empty column D — all modes

Empty pattern (column D blank) = match any non-empty cell, regardless of mode. This preserves the existing "relaxed default" guarantee and means existing patterns that omit the regex column are unaffected.

---

## Migration from current behaviour

**The breaking case:** existing `lbl:` rows with regex content in column D.

| Existing column D | After change | Action needed |
|---|---|---|
| *(empty)* | Still matches anything | None |
| `.*` | Would match literal `.*` | Delete the value (empty = match anything) |
| `Amount.*EUR` | Would match literal `Amount.*EUR` | Change column A to `lbl:regexp` |
| `Breaks\n\(minutes\)` | Would match literal `Breaks\n\(minutes\)` | Change column A to `lbl:regexp`, OR change column D to `Breaks\n(minutes)` and keep `lbl:literal` |

**Easiest migration path for a whole pattern file:** add `config: | lbl.match | regexp` — zero field changes needed, exact old behaviour.

**Best long-term approach for existing patterns:** run `validate-pattern`, follow the warnings field by field.

---

## validate-pattern warnings (new)

When the resolved mode for an `lbl:` field is `literal` or `glob` and column D value contains unambiguous regex tokens, emit a WARN:

- `.*`, `\d`, `\w`, `\s`, `\b` — regex metacharacter sequences
- Unescaped `(`, `[`, `^`, `$`, `+`, `?` — regex special characters

Warning text example:
```
WARN  lbl: 'col_breaks' — value 'Breaks\n\(minutes\)' looks like a regex pattern
      but lbl.match is 'literal'. Use lbl:regexp or rewrite as 'Breaks\n(minutes)'.
```

---

## Implementation checklist

### Phase 1 — Core model + parser ✓
- [x] `grepxcel/models.py` — added `LBL_MATCH_MODES` frozenset and `Config.lbl_match: str = 'literal'`; `FieldDef.lbl_match: str | None = None`
- [x] `grepxcel/pattern_parser.py` — parse `config: | lbl.match | <mode>`; validate value; `isinstance` guard for None cells
- [x] `grepxcel/pattern_parser.py` — recognize `lbl:literal`, `lbl:glob`, `lbl:regexp` in column A; suffix parsed and stored on `FieldDef`; unknown suffixes raise PatternError
- [x] `grepxcel/pattern_parser.py` — `check_regex_safety` skipped for `lbl:` fields when mode is not `regexp`

### Phase 2 — Engine match dispatcher ✓
- [x] `grepxcel/engine.py` — `_match_lbl()` and `_resolve_lbl_mode()` added
- [x] `grepxcel/engine.py` — all lbl anchor checks now go through `_match_lbl`

### Phase 3 — validate-pattern warnings ✓
- [x] `grepxcel/pattern_check.py` — `_REGEX_TELL` regex detects backslash-escapes and lookahead
- [x] Warns when lbl.match mode is literal/glob but pattern looks like regex; suggests plain text and regexp opt-in
- [x] Verbose output shows `lbl.match` in config block and `[mode]` tags on fields with per-field override

### Phase 4 — Drafter ✓
- [x] `grepxcel/drafter.py` — prompt updated: lbl: column D is now described as "plain text verbatim, no regex escaping"

### Phase 5 — Fixtures + tests ✓
- [x] Fixture 11 (`pattern-from-draft.csv` + `.xlsx`) — updated `Term \(months\):` → `Term (months):`
- [x] 45 new unit tests: `_match_lbl` (all modes × edge cases), `_resolve_lbl_mode`, parser lbl: suffixes, validate-pattern warnings
- [x] 1251 tests passing, 7 skipped, 4 xfailed

### Phase 6 — Docs (partial)
- [ ] `grepxcel docs` / `pattern-reference.xlsx` — document the three modes and the column A variants
- [ ] CHANGELOG entry

---

## Known risks

| Risk | Mitigation |
|---|---|
| Glob `*` across newlines (Alt+Enter cells) | Use `re.DOTALL` in compiled glob pattern |
| `fnmatch` case sensitivity | Compile with `re.IGNORECASE` when `ignore.case` is set |
| `[abc]` in glob mode acting as regex character class | `fnmatch.translate` handles this correctly — it is a character class in glob too |
| Fixture migration misses a field | `validate-pattern` warnings catch it before tests run |
| `col:` rows (column anchors) have same problem | Deferred — `col:` has same structural role but lower metacharacter incidence; tackle in a follow-up |

---

## Out of scope (follow-up)

- `col:literal`, `col:glob`, `col:regexp` — same pattern for column anchors
- Per-field `lbl:re` alias (decided: one canonical name `lbl:regexp`, no aliases)
- Fuzzy / similarity matching mode
