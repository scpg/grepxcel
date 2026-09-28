"""An ignored cell must come back ignored when a pattern is reloaded.

`cell:<ref> IGNORE` records a decision the author made: skip this cell. Preload
used to drop those steps entirely, so the cell reappeared unclassified — the
pattern said "skip this", the wizard showed nothing — and because the cell was
never marked as claimed, a later field could match it and quietly take it over.

Also covers the mini-table modifier list, whose missing 'not-null +
trim-whitespace' entry made that combination unreachable for a table column even
though the parser accepts it.
"""
import re
from pathlib import Path

import openpyxl
import pytest

from grepxcel.pattern_parser import VALID_CONSTRAINT_TOKENS
from grepxcel.wizard_core import _preload_from_pattern
from tests.unit.test_random_type_fixtures import write_pattern_csv


def _sheet(tmp_path):
    path = str(tmp_path / 'data.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = 'Invoice No'
    ws['B1'] = 'INV-001'
    ws['C1'] = 'internal-only'
    ws['A2'] = 'Total'
    ws['B2'] = 100
    wb.save(path)
    return openpyxl.load_workbook(path).active


def _preload(tmp_path, rows):
    pattern = str(tmp_path / 'p.csv')
    write_pattern_csv(pattern, rows)
    return _preload_from_pattern(_sheet(tmp_path), pattern)


class TestIgnoreIsRestored:
    def test_an_absolute_ignore_cell_loads_as_ignored(self, tmp_path):
        choices, _cfg, _warns = _preload(tmp_path, [
            ['lbl:', 'inv_label', 'string', 'Invoice No'],
            ['var:', 'inv', 'string', '.*'],
            [], ['START:'],
            ['cell:A1', 'inv_label'],
            ['cell:B1', 'inv'],
            ['cell:C1', 'IGNORE'],
            ['END:'],
        ])
        assert 'C1' in choices, f'ignored cell missing from preload: {sorted(choices)}'
        assert choices['C1'] == {'choice': 'I'}

    def test_the_other_cells_still_load(self, tmp_path):
        choices, _cfg, _warns = _preload(tmp_path, [
            ['lbl:', 'inv_label', 'string', 'Invoice No'],
            ['var:', 'inv', 'string', '.*'],
            [], ['START:'],
            ['cell:A1', 'inv_label'],
            ['cell:B1', 'inv'],
            ['cell:C1', 'IGNORE'],
            ['END:'],
        ])
        assert choices['A1']['choice'] == 'L'
        assert choices['B1']['choice'] == 'V'

    def test_an_ignored_cell_cannot_be_claimed_by_a_later_field(self, tmp_path):
        """The subtler half of the bug: an unclaimed ignored cell was still a
        candidate for a later label/value search, so a field could silently
        land on a cell the author had explicitly excluded."""
        choices, _cfg, _warns = _preload(tmp_path, [
            ['var:', 'anything', 'string', 'internal-only'],
            [], ['START:'],
            ['cell:C1', 'IGNORE'],
            ['cell:next', 'anything'],
            ['END:'],
        ])
        assert choices.get('C1') == {'choice': 'I'}, \
            'C1 should remain ignored, not be taken over by the var: field'

    def test_a_sequential_ignore_is_reported_not_dropped(self, tmp_path):
        """cell:next IGNORE depends on the live scan cursor, which preload does
        not simulate. It cannot be placed — but it must not vanish silently."""
        _choices, _cfg, warns = _preload(tmp_path, [
            ['lbl:', 'inv_label', 'string', 'Invoice No'],
            [], ['START:'],
            ['cell:A1', 'inv_label'],
            ['cell:next', 'IGNORE'],
            ['END:'],
        ])
        assert any('IGNORE' in w for w in warns), warns

    def test_an_unparsable_ignore_ref_is_reported(self, tmp_path):
        """A malformed ref must warn rather than silently skip the cell."""
        _choices, _cfg, warns = _preload(tmp_path, [
            ['lbl:', 'inv_label', 'string', 'Invoice No'],
            [], ['START:'],
            ['cell:A1', 'inv_label'],
            ['cell:next', 'IGNORE'],
            ['END:'],
        ])
        assert warns, 'expected at least one warning'


class TestMiniTableModifierList:
    """The mini-table column panel renders its own <select>; it must offer every
    combination the parser accepts."""

    JS = (Path(__file__).resolve().parents[2]
          / 'grepxcel' / 'static' / 'wizard.js').read_text(encoding='utf-8')

    def _offered(self):
        block = self.JS.split('const _TRM_MODIFIERS', 1)[1].split('];', 1)[0]
        return {v for v, _label in re.findall(r"\['([^']+)',\s*'([^']+)'\]", block)}

    def test_every_single_constraint_is_offered(self):
        offered = self._offered()
        # not-empty is an alias of not-null; the wizard standardises on not-null.
        expected = (set(VALID_CONSTRAINT_TOKENS) - {'not-empty'}) | {'none'}
        missing = expected - offered
        assert not missing, f'mini-table modifiers missing: {sorted(missing)}'

    def test_the_trim_combinations_are_both_offered(self):
        """'not-null:trim-whitespace' was absent, making it unreachable for a
        table column although the parser accepts it."""
        offered = self._offered()
        assert 'nullable:trim-whitespace' in offered
        assert 'not-null:trim-whitespace' in offered

    def test_no_contradictory_combination_is_offered(self):
        """nullable + not-null is rejected by the parser, so offering it would
        produce a pattern that cannot be re-read."""
        for value in self._offered():
            parts = value.split(':')
            assert not ('nullable' in parts and 'not-null' in parts), value

    def test_every_offered_token_is_one_the_parser_knows(self):
        for value in self._offered():
            if value == 'none':
                continue
            for token in value.split(':'):
                assert token in VALID_CONSTRAINT_TOKENS, \
                    f'{token!r} is not a parser constraint token'


class TestLoadedPatternNotice:
    """A pattern file can express more than the wizard models, so opening one
    raises a one-off notice. The flag that drives it must be set only when a
    pattern really was opened AND something was recognised — otherwise the
    notice would fire on a fresh session with nothing to lose."""

    HTML = (Path(__file__).resolve().parents[2]
            / 'grepxcel' / 'templates' / 'wizard.html').read_text(encoding='utf-8')
    JS = (Path(__file__).resolve().parents[2]
          / 'grepxcel' / 'static' / 'wizard.js').read_text(encoding='utf-8')

    def test_the_api_exposes_the_flag(self):
        assert "'pattern_was_loaded'" in (
            Path(__file__).resolve().parents[2] / 'grepxcel' / 'wizard_api.py'
        ).read_text(encoding='utf-8')

    def test_the_modal_exists_and_is_wired(self):
        assert 'id="loaded-pattern-modal"' in self.HTML
        assert 'closeLoadedPatternNotice()' in self.HTML
        assert 'function closeLoadedPatternNotice' in self.JS
        assert '_maybeShowLoadedPatternNotice' in self.JS

    def test_the_notice_names_the_concrete_limitation(self):
        """Generic 'may not work' wording would not tell anyone what to check.
        The glob case is the one that motivated it: a rule matching many labels
        cannot be represented as one cell, one classification."""
        lowered = self._modal_html().lower()
        assert 'glob' in lowered
        assert 'seek:' in lowered

    def test_the_notice_says_what_saving_does(self):
        assert 'save' in self._modal_html().lower()

    def _modal_html(self) -> str:
        """The modal block: from its id up to the toast that follows it."""
        after = self.HTML.split('id="loaded-pattern-modal"', 1)[1]
        return after.split('id="toast"', 1)[0]

    def test_the_notice_can_be_dismissed_permanently(self):
        assert 'lpm-dont-show' in self.HTML
        assert 'localStorage' in self.JS.split('_LPM_DISMISS_KEY', 1)[1][:800]

    def test_buttons_call_functions_that_exist(self):
        """A dead onclick fails silently in the browser and nothing in the test
        suite executes this JS, so the handler names are checked here."""
        for handler in ('closeLoadedPatternNotice', 'switchTab'):
            assert f'function {handler}' in self.JS, f'{handler} is not defined'
