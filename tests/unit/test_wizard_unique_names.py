"""Generated field names must be unique, and a repeat must be reported.

A sheet that repeats the same labelled block — three identical test columns side
by side — used to produce the same generated name three times, because names are
slugified from the cell's own text and nothing checked for a collision.

That is not cosmetic. `engine._process_cell` does
``result['cells'][field] = value`` unconditionally, so the last write wins and
the earlier values are discarded with no warning anywhere: the pattern validates,
the run reports zero warnings, and two thirds of the data never appears.

Two independent defences, tested here:
  1. `wizard_api._unique_name` — stop the wizard generating the collision.
  2. `pattern_check._overwritten_var_fields` — catch it in ANY pattern, whether
     the wizard, a hand edit or another tool produced it.
"""
import pytest

from grepxcel.pattern_check import _overwritten_var_fields, check_pattern
from grepxcel.wizard_api import _unique_name
from tests.unit.test_random_type_fixtures import write_pattern_csv


class TestUniqueName:
    def test_first_use_is_unchanged(self):
        assert _unique_name('total', {}, 'A1') == 'total'

    def test_a_collision_is_numbered(self):
        choices = {'A1': {'name': 'total'}}
        assert _unique_name('total', choices, 'B1') == 'total_2'

    def test_numbering_continues_past_existing_suffixes(self):
        choices = {'A1': {'name': 'total'}, 'B1': {'name': 'total_2'}}
        assert _unique_name('total', choices, 'C1') == 'total_3'

    def test_suffix_is_applied_after_the_number(self):
        """Label names end in _label; the counter must not land after it, or
        'x_label_2' and 'x_2_label' would both appear for the same base."""
        choices = {'A1': {'name': 'total_label'}}
        assert _unique_name('total', choices, 'B1', '_label') == 'total_2_label'

    def test_reclassifying_the_same_cell_keeps_its_name(self):
        """own_ref is excluded, so editing a cell does not bump it every time."""
        choices = {'A1': {'name': 'total'}}
        assert _unique_name('total', choices, 'A1') == 'total'

    def test_unrelated_names_do_not_collide(self):
        choices = {'A1': {'name': 'subtotal'}}
        assert _unique_name('total', choices, 'B1') == 'total'

    def test_three_repeated_blocks_get_three_names(self):
        """The reported case: the same label text classified in three columns."""
        choices = {}
        names = []
        for ref in ('A2', 'D2', 'G2'):
            name = _unique_name('boolean_bool_test', choices, ref)
            choices[ref] = {'name': name}
            names.append(name)
        assert names == ['boolean_bool_test', 'boolean_bool_test_2',
                         'boolean_bool_test_3']
        assert len(set(names)) == 3


def _pattern_with_repeats(tmp_path, *, distinct_names: bool):
    """Three cells feeding one var: field, or three separate fields."""
    rows = [['var:', 'v' if not distinct_names else 'v1', 'string', '.*']]
    if distinct_names:
        rows += [['var:', 'v2', 'string', '.*'], ['var:', 'v3', 'string', '.*']]
    rows += [[], ['START:']]
    targets = ['v1', 'v2', 'v3'] if distinct_names else ['v', 'v', 'v']
    for ref, field in zip(('A1', 'B1', 'C1'), targets):
        rows.append([f'cell:{ref}', field])
    rows.append(['END:'])
    path = str(tmp_path / 'pattern.csv')
    write_pattern_csv(path, rows)
    return path


class TestValidatePatternCatchesOverwrites:
    def test_repeated_writes_are_detected(self, tmp_path):
        path = _pattern_with_repeats(tmp_path, distinct_names=False)
        result = check_pattern(path)
        dupes = _overwritten_var_fields(result.sequence, result.defs)
        assert dupes == {'v': ['cell:A1', 'cell:B1', 'cell:C1']}

    def test_the_warning_says_data_is_discarded(self, tmp_path):
        """Message accuracy: 'duplicate field name' would understate it — the
        user needs to know values are being thrown away."""
        path = _pattern_with_repeats(tmp_path, distinct_names=False)
        result = check_pattern(path)
        warning = next(w for w in result.warnings if "'v'" in w)
        assert 'written 3 times' in warning
        assert 'discarded' in warning
        assert 'cell:A1' in warning and 'cell:C1' in warning

    def test_the_pattern_is_still_valid(self, tmp_path):
        """A warning, not an error — the pattern runs, it just loses data."""
        path = _pattern_with_repeats(tmp_path, distinct_names=False)
        assert check_pattern(path).valid is True

    def test_distinct_names_produce_no_warning(self, tmp_path):
        path = _pattern_with_repeats(tmp_path, distinct_names=True)
        result = check_pattern(path)
        assert _overwritten_var_fields(result.sequence, result.defs) == {}
        assert not [w for w in result.warnings if 'written' in w]

    def test_a_label_reused_as_an_anchor_is_not_flagged(self, tmp_path):
        """lbl: fields carry no output value, so re-anchoring loses nothing."""
        path = str(tmp_path / 'p.csv')
        write_pattern_csv(path, [
            ['lbl:', 'a', 'string', 'ANCHOR'],
            ['var:', 'v1', 'string', '.*'],
            ['var:', 'v2', 'string', '.*'],
            [], ['START:'],
            ['cell:A1', 'a'], ['cell:B1', 'v1'],
            ['cell:A2', 'a'], ['cell:B2', 'v2'],
            ['END:'],
        ])
        result = check_pattern(path)
        assert _overwritten_var_fields(result.sequence, result.defs) == {}

    def test_a_table_column_is_not_flagged(self, tmp_path):
        """A DATA: field is written once per row by design."""
        path = str(tmp_path / 'p.csv')
        write_pattern_csv(path, [
            ['lbl:', 'h', 'string', 'COLHEAD'],
            ['var:', 'items.c', 'string', '.*'],
            [], ['START:'], ['table:*'],
            ['', 'HEADER:1', 'h'], ['', 'DATA:*', 'items.c'], ['END:'],
        ])
        result = check_pattern(path)
        assert _overwritten_var_fields(result.sequence, result.defs) == {}


def test_valid_config_keys_are_listed_accurately(tmp_path):
    """The unknown-key message restated its list and drifted: date.format and
    time.format were accepted by the parser while being reported as invalid."""
    path = str(tmp_path / 'p.csv')
    write_pattern_csv(path, [
        ['config:', 'not.a.real.key', 'x'],
        ['var:', 'v', 'string', '.*'],
        [], ['START:'], ['cell:A1', 'v'], ['END:'],
    ])
    warning = next(w for w in check_pattern(path).warnings if 'Unknown config key' in w)
    for key in ('date.format', 'time.format', 'read.direction', 'lbl.match'):
        assert key in warning, f'{key} missing from the valid-keys list'
