"""Tests for `grepxcel validate-pattern` (grepxcel.pattern_check)."""
import io

import openpyxl

from grepxcel.pattern_check import check_pattern, run_validate


def _write(rows, tmp_path, name='pattern.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    path = str(tmp_path / name)
    wb.save(path)
    return path


_VALID = [['var:', 'x.v', 'string', '.*'], ['START:'], ['cell:A1', 'x.v'], ['END:']]


def test_valid_pattern(tmp_path):
    r = check_pattern(_write(_VALID, tmp_path))
    assert r.valid and not r.errors
    assert r.n_fields == 1 and r.n_steps == 1


def test_unknown_type_is_invalid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'currncy', '.*'], ['START:'], ['cell:A1', 'x.v'], ['END:']], tmp_path))
    assert not r.valid
    assert any('Unknown field type' in e for e in r.errors)


def test_undefined_field_is_invalid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'string', '.*'], ['START:'], ['cell:A1', 'ghost'], ['END:']], tmp_path))
    assert not r.valid
    assert any("'ghost'" in e and 'never defined' in e for e in r.errors)


def test_empty_sequence_is_invalid(tmp_path):
    r = check_pattern(_write([['var:', 'x.v', 'string', '.*']], tmp_path))  # no START:
    assert not r.valid
    assert any('No extraction steps' in e for e in r.errors)


def test_unused_field_warns_but_valid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'string', '.*'], ['var:', 'y.z', 'string', '.*'],
         ['START:'], ['cell:A1', 'x.v'], ['END:']], tmp_path))
    assert r.valid                              # unused is a warning, not an error
    assert any("'y.z'" in w for w in r.warnings)


def test_redos_regex_is_invalid(tmp_path):
    r = check_pattern(_write(
        [['var:', 'x.v', 'string', r'(a+)+$'], ['START:'], ['cell:A1', 'x.v'], ['END:']], tmp_path))
    assert not r.valid
    assert any('regex' in e.lower() for e in r.errors)


def test_missing_file(tmp_path):
    r = check_pattern(str(tmp_path / 'nope.xlsx'))
    assert not r.valid
    assert r.errors


def test_run_validate_exit_codes(tmp_path):
    good = _write(_VALID, tmp_path, 'good.xlsx')
    bad = _write([['var:', 'x.v', 'currncy', '.*'], ['START:'], ['cell:A1', 'x.v'], ['END:']],
                 tmp_path, 'bad.xlsx')
    assert run_validate([good], out=io.StringIO()) == 0
    assert run_validate([bad], out=io.StringIO()) == 1
    assert run_validate([good, bad], out=io.StringIO()) == 1   # any invalid → non-zero


def test_verbose_shows_fields_and_sequence(tmp_path):
    buf = io.StringIO()
    run_validate([_write(_VALID, tmp_path)], verbose=True, out=buf)
    out = buf.getvalue()
    assert 'VALID' in out and 'fields:' in out and 'extraction sequence:' in out and 'x.v' in out


def test_verbose_shows_seek_instruction(tmp_path):
    """seek: appears in the verbose extraction-sequence listing."""
    rows = [
        ['var:', 'x', 'string', '.*'],
        ['START:'],
        ['seek:G5'],
        ['cell:next', 'x'],
        ['END:'],
    ]
    buf = io.StringIO()
    run_validate([_write(rows, tmp_path)], verbose=True, out=buf)
    out = buf.getvalue()
    assert 'seek:G5' in out


def test_verbose_table_columnar_format(tmp_path):
    """Table rows render as a columnar grid, one column-position per line."""
    rows = [
        ['lbl:', 'h1', 'string', 'Name'],
        ['lbl:', 'h2', 'string', 'Age'],
        ['var:', 'name', 'string', '.*'],
        ['var:', 'age', 'integer', r'\d+'],
        ['START:'],
        ['table:1'],
        ['', 'HEADER:1', 'h1', 'h2'],
        ['', 'DATA:*', 'name', 'age'],
        ['END:'],
    ]
    buf = io.StringIO()
    run_validate([_write(rows, tmp_path)], verbose=True, out=buf)
    out = buf.getvalue()
    lines = out.splitlines()
    # Find the table block
    table_idx = next(i for i, l in enumerate(lines) if 'table:1' in l)
    # Next line should be the header row with row-type labels
    header_line = lines[table_idx + 1]
    assert 'HEADER:1' in header_line
    assert 'DATA:*' in header_line
    # Column positions follow, one per line
    col1_line = lines[table_idx + 2]
    assert 'h1' in col1_line and 'name' in col1_line
    col2_line = lines[table_idx + 3]
    assert 'h2' in col2_line and 'age' in col2_line


def test_verbose_table_with_skip_if(tmp_path):
    """SKIP_IF rows appear in the columnar table grid."""
    rows = [
        ['lbl:', 'h', 'string', 'Val'],
        ['var:', 'v', 'string', '.*'],
        ['START:'],
        ['table:1'],
        ['', 'HEADER:1', 'h'],
        ['', 'SKIP_IF', 'EMPTY'],
        ['', 'DATA:*', 'v'],
        ['END:'],
    ]
    buf = io.StringIO()
    run_validate([_write(rows, tmp_path)], verbose=True, out=buf)
    out = buf.getvalue()
    lines = out.splitlines()
    table_idx = next(i for i, l in enumerate(lines) if 'table:1' in l)
    header_line = lines[table_idx + 1]
    assert 'SKIP_IF' in header_line
    # The single column position should show all three row types
    col_line = lines[table_idx + 2]
    assert 'h' in col_line and 'EMPTY' in col_line and 'v' in col_line


# ── lbl.match warnings ────────────────────────────────────────────────────────

def _lbl_pattern(rows, tmp_path):
    """Write rows and return check_pattern result."""
    return check_pattern(_write(rows, tmp_path))


def _lbl_base(lbl_pattern, lbl_col_a='lbl:'):
    return [
        [lbl_col_a, 'h', 'string', lbl_pattern],
        ['var:', 'v', 'string', '.*'],
        ['START:'],
        ['cell:1', 'h'],
        ['cell:1', 'v'],
        ['END:'],
    ]


def test_regex_escaped_paren_warns_in_literal_mode(tmp_path):
    r = _lbl_pattern(_lbl_base(r'Term \(months\):'), tmp_path)
    assert r.valid
    assert any('looks like a regex' in w for w in r.warnings)


def test_regex_escaped_dot_warns_in_literal_mode(tmp_path):
    r = _lbl_pattern(_lbl_base(r'Version\.2'), tmp_path)
    assert r.valid
    assert any('looks like a regex' in w for w in r.warnings)


def test_lookahead_warns_in_literal_mode(tmp_path):
    r = _lbl_pattern(_lbl_base(r'(?:Amount)'), tmp_path)
    assert r.valid
    assert any('looks like a regex' in w for w in r.warnings)


def test_plain_literal_no_warning(tmp_path):
    r = _lbl_pattern(_lbl_base('Term (months):'), tmp_path)
    assert r.valid
    assert not any('looks like a regex' in w for w in r.warnings)


def test_regexp_override_suppresses_warning(tmp_path):
    r = _lbl_pattern(_lbl_base(r'Term \(months\):', 'lbl:regexp'), tmp_path)
    assert r.valid
    assert not any('looks like a regex' in w for w in r.warnings)


def test_global_regexp_mode_suppresses_warning(tmp_path):
    rows = [
        ['config:', 'lbl.match', 'regexp'],
        ['lbl:', 'h', 'string', r'Term \(months\):'],
        ['var:', 'v', 'string', '.*'],
        ['START:'], ['cell:1', 'h'], ['cell:1', 'v'], ['END:'],
    ]
    r = _lbl_pattern(rows, tmp_path)
    assert r.valid
    assert not any('looks like a regex' in w for w in r.warnings)


def test_verbose_shows_lbl_match_config(tmp_path):
    buf = io.StringIO()
    run_validate([_write(_lbl_base('Name'), tmp_path)], verbose=True, out=buf)
    assert 'lbl.match' in buf.getvalue()


def test_verbose_shows_per_field_mode_tag(tmp_path):
    buf = io.StringIO()
    run_validate([_write(_lbl_base('Hello *', 'lbl:glob'), tmp_path)], verbose=True, out=buf)
    assert '[glob]' in buf.getvalue()


# ── verbose config: complete defaults ────────────────────────────────────────

def _run_verbose(rows, tmp_path):
    """Write pattern rows, run validate-pattern -v, return output string."""
    buf = io.StringIO()
    run_validate([_write(rows, tmp_path)], verbose=True, out=buf)
    return buf.getvalue()


def test_verbose_config_shows_all_eight_keys(tmp_path):
    """All 8 config keys must appear in verbose output for any valid pattern."""
    out = _run_verbose(_VALID, tmp_path)
    for key in ('pattern.version', 'read.direction', 'currency.sign',
                'ignore.case', 'trim.whitespace', 'lbl.match', 'var.match',
                'empty.aliases'):
        assert key in out, f'Missing config key in -v output: {key!r}'


def test_verbose_config_defaults_pattern_version(tmp_path):
    """pattern.version without an explicit declaration shows the 'defaulted' note."""
    out = _run_verbose(_VALID, tmp_path)
    assert 'defaulted' in out


def test_verbose_config_default_read_direction(tmp_path):
    out = _run_verbose(_VALID, tmp_path)
    assert 'read.direction  LR' in out


def test_verbose_config_default_currency_sign(tmp_path):
    out = _run_verbose(_VALID, tmp_path)
    assert 'currency.sign   €' in out


def test_verbose_config_default_ignore_case(tmp_path):
    out = _run_verbose(_VALID, tmp_path)
    assert 'ignore.case     False' in out


def test_verbose_config_default_trim_whitespace(tmp_path):
    """trim.whitespace defaults to False and must appear in verbose output."""
    out = _run_verbose(_VALID, tmp_path)
    assert 'trim.whitespace False' in out


def test_verbose_config_default_lbl_match(tmp_path):
    out = _run_verbose(_VALID, tmp_path)
    assert 'lbl.match       literal' in out


def test_verbose_config_default_var_match(tmp_path):
    """var.match defaults to regexp and must appear in verbose output."""
    out = _run_verbose(_VALID, tmp_path)
    assert 'var.match       regexp' in out


def test_verbose_config_default_empty_aliases_none(tmp_path):
    """empty.aliases shows (none) when no aliases are configured."""
    out = _run_verbose(_VALID, tmp_path)
    assert 'empty.aliases   (none)' in out


def test_verbose_config_section_key_order(tmp_path):
    """Config keys must appear in the documented order in -v output."""
    out = _run_verbose(_VALID, tmp_path)
    keys = ('pattern.version', 'read.direction', 'currency.sign',
            'ignore.case', 'trim.whitespace', 'lbl.match', 'var.match',
            'empty.aliases')
    positions = [out.index(k) for k in keys]
    assert positions == sorted(positions), (
        f'Config keys out of order. Positions: {list(zip(keys, positions))}'
    )


# ── verbose config: non-default values ───────────────────────────────────────

def _cfg_pattern(config_rows, tmp_path):
    """Build a pattern with the given config rows prepended to _VALID fields."""
    rows = config_rows + list(_VALID)
    return _run_verbose(rows, tmp_path)


def test_verbose_config_trim_whitespace_true(tmp_path):
    out = _cfg_pattern([['config:', 'trim.whitespace', 'yes']], tmp_path)
    assert 'trim.whitespace True' in out


def test_verbose_config_var_match_glob(tmp_path):
    out = _cfg_pattern([['config:', 'var.match', 'glob']], tmp_path)
    assert 'var.match       glob' in out


def test_verbose_config_var_match_literal(tmp_path):
    out = _cfg_pattern([['config:', 'var.match', 'literal']], tmp_path)
    assert 'var.match       literal' in out


def test_verbose_config_empty_aliases_single(tmp_path):
    out = _cfg_pattern([['config:', 'empty.aliases', 'N/A']], tmp_path)
    assert 'empty.aliases   N/A' in out


def test_verbose_config_empty_aliases_multiple(tmp_path):
    """Multiple empty.aliases config rows are shown comma-separated."""
    rows = [
        ['config:', 'empty.aliases', 'N/A'],
        ['config:', 'empty.aliases', 'TBD'],
        ['config:', 'empty.aliases', '-'],
    ]
    out = _cfg_pattern(rows, tmp_path)
    assert 'N/A' in out and 'TBD' in out and '-' in out


def test_verbose_config_ignore_case_true(tmp_path):
    out = _cfg_pattern([['config:', 'ignore.case', 'yes']], tmp_path)
    assert 'ignore.case     True' in out


def test_verbose_config_pattern_version_explicit(tmp_path):
    """Explicitly declared pattern.version must NOT show the 'defaulted' note."""
    out = _cfg_pattern([['config:', 'pattern.version', '1']], tmp_path)
    assert 'pattern.version 1' in out
    assert 'defaulted' not in out


# ── empty.aliases validation warnings ────────────────────────────────────────

def _alias_pattern(alias_rows, tmp_path, extra_config=None):
    """Build and check a pattern with the given empty.aliases config rows."""
    rows = (extra_config or []) + alias_rows + list(_VALID)
    return check_pattern(_write(rows, tmp_path))


def test_alias_known_excel_error_no_warning(tmp_path):
    """A well-formed Excel error string like '#N/A' must not trigger any warning."""
    r = _alias_pattern([['config:', 'empty.aliases', '#N/A']], tmp_path)
    assert r.valid
    assert not any('#N/A' in w and 'not a recognised' in w for w in r.warnings)
    assert not any('#N/A' in w and 'capitalisation' in w for w in r.warnings)


def test_alias_all_known_excel_errors_accepted(tmp_path):
    """Every canonical Excel error string is recognised without a warning."""
    known = ['#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!']
    for err in known:
        r = _alias_pattern([['config:', 'empty.aliases', err]], tmp_path)
        assert r.valid, f'Unexpected invalid for {err!r}'
        assert not any('not a recognised' in w for w in r.warnings), \
            f'False-positive "not recognised" warning for {err!r}'


def test_alias_unknown_hash_string_warns(tmp_path):
    """A '#'-prefixed alias that is not a known Excel error triggers a warning."""
    r = _alias_pattern([['config:', 'empty.aliases', '#UNKNOWN!']], tmp_path)
    assert r.valid   # warning, not error
    assert any('not a recognised' in w for w in r.warnings)


def test_alias_wrong_case_warns_without_ignore_case(tmp_path):
    """'#n/a' (wrong case) warns about capitalisation when ignore.case is off."""
    r = _alias_pattern([['config:', 'empty.aliases', '#n/a']], tmp_path)
    assert r.valid
    assert any('capitalisation' in w for w in r.warnings)


def test_alias_wrong_case_no_warn_with_ignore_case(tmp_path):
    """'#n/a' with ignore.case yes: no capitalisation warning (case won't matter)."""
    r = _alias_pattern(
        [['config:', 'empty.aliases', '#n/a']],
        tmp_path,
        extra_config=[['config:', 'ignore.case', 'yes']],
    )
    assert r.valid
    assert not any('capitalisation' in w for w in r.warnings)


def test_alias_duplicate_warns(tmp_path):
    """Two identical aliases produce a 'Duplicate' warning."""
    r = _alias_pattern([
        ['config:', 'empty.aliases', 'N/A'],
        ['config:', 'empty.aliases', 'N/A'],
    ], tmp_path)
    assert r.valid
    assert any('Duplicate' in w for w in r.warnings)


def test_alias_duplicate_case_insensitive_warns(tmp_path):
    """'N/A' and 'n/a' are duplicates when ignore.case is on."""
    r = _alias_pattern(
        [['config:', 'empty.aliases', 'N/A'],
         ['config:', 'empty.aliases', 'n/a']],
        tmp_path,
        extra_config=[['config:', 'ignore.case', 'yes']],
    )
    assert r.valid
    assert any('Duplicate' in w for w in r.warnings)


def test_alias_duplicate_case_sensitive_no_warn(tmp_path):
    """'N/A' and 'n/a' are NOT duplicates when ignore.case is off."""
    r = _alias_pattern([
        ['config:', 'empty.aliases', 'N/A'],
        ['config:', 'empty.aliases', 'n/a'],
    ], tmp_path)
    assert r.valid
    assert not any('Duplicate' in w for w in r.warnings)


def test_alias_regex_metacharacter_warns(tmp_path):
    """An alias containing regex metacharacters (e.g. '.*') warns that it is
    matched as a plain string, not a regex."""
    r = _alias_pattern([['config:', 'empty.aliases', '.*']], tmp_path)
    assert r.valid
    assert any('metacharacter' in w for w in r.warnings)


def test_alias_plain_string_no_metachar_warning(tmp_path):
    """A plain alias like 'N/A' or '-' has no metacharacter warning."""
    r = _alias_pattern([['config:', 'empty.aliases', 'N/A']], tmp_path)
    assert r.valid
    assert not any('metacharacter' in w for w in r.warnings)


def test_alias_no_warnings_when_no_aliases(tmp_path):
    """A pattern with no empty.aliases config produces no alias-related warnings."""
    r = check_pattern(_write(list(_VALID), tmp_path))
    assert r.valid
    assert not any('alias' in w.lower() for w in r.warnings)
