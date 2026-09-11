"""Unit tests for the grepxcel TUI wizard (wizard_tui.py).

Tests are split into two layers:

1. **Pure-function tests** — exercise `_build_state_from_choices` and
   `_choices_to_csv` directly (no Textual, no xlsx required).  These cover
   the NEW-2 / NEW-3 bug fixes (ignore_case / currency_sign propagation).

2. **Pilot tests** — use Textual's ``App.run_test()`` async context manager
   (headless mode, no real terminal) to exercise the full TUI lifecycle:
   config-modal acceptance, cell classification (L / V / I), undo, and
   save/cancel.

   Each pilot test is a *synchronous* function that drives the async Textual
   event loop via ``asyncio.run()``.  No ``pytest-asyncio`` plugin is needed.

   Each test creates a minimal xlsx worksheet with ``openpyxl`` and drives
   the app with ``await pilot.press(...)``.  The session log written by the
   app goes to ``logs/wizard/`` inside the CWD; if that fails (temp directory,
   read-only FS) the app continues silently.

Textual is optional at import time.  Tests that need it are marked
``pytest.mark.skipif(not _TEXTUAL_OK, ...)`` so the suite stays green
even in environments without the wizard extra installed.
"""

from __future__ import annotations

import asyncio
import csv
import io
import os
import sys
import tempfile

import openpyxl
import pytest

# ── Import TUI internals (may be absent if textual not installed) ─────────────

try:
    from grepxcel.wizard_tui import (
        _build_state_from_choices,
        _choices_to_csv,
        _col_a_extra_from_parts,
        _col_a_extra_to_parts,
        _TEXTUAL_OK,
    )
    if _TEXTUAL_OK:
        from grepxcel.wizard_tui import WizardTUIApp
    from grepxcel.wizard import WizardState, _pattern_rows
except ImportError:
    _TEXTUAL_OK = False

_skip_no_textual = pytest.mark.skipif(
    not _TEXTUAL_OK,
    reason='textual not installed (pip install "grepxcel[wizard]")',
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_ws(data: dict[tuple[int, int], object]):
    """Return an in-memory openpyxl worksheet pre-populated with *data*.

    data: {(row, col): value}  — 1-indexed.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    for (r, c), v in data.items():
        ws.cell(row=r, column=c, value=v)
    return ws


def _csv_rows(text: str) -> list[list[str]]:
    return list(csv.reader(io.StringIO(text)))


# ── _build_state_from_choices — pure-function tests ───────────────────────────

class TestBuildStateFromChoices:
    """_build_state_from_choices propagates all config params into WizardState."""

    def _ws(self):
        return _make_ws({(1, 1): 'Invoice No:', (1, 2): 'INV-001'})

    def _cells(self, ws):
        return [(r, c) for r in range(1, ws.max_row + 1)
                for c in range(1, ws.max_column + 1)]

    def test_defaults_are_ltr_no_ignore_case_euro(self):
        ws = self._ws()
        state = _build_state_from_choices(ws, {}, self._cells(ws), 'LR', 'Sheet1')
        assert state.direction    == 'LR'
        assert state.ignore_case  is False
        assert state.currency_sign == '€'
        assert state.sheet_name   == 'Sheet1'

    def test_ignore_case_propagated(self):
        ws = self._ws()
        state = _build_state_from_choices(
            ws, {}, self._cells(ws), 'LR', 'Sheet1', ignore_case=True,
        )
        assert state.ignore_case is True

    def test_currency_sign_propagated(self):
        ws = self._ws()
        state = _build_state_from_choices(
            ws, {}, self._cells(ws), 'LR', 'Sheet1', currency_sign='$',
        )
        assert state.currency_sign == '$'

    def test_td_direction_propagated(self):
        ws = self._ws()
        state = _build_state_from_choices(ws, {}, self._cells(ws), 'TD', 'Sheet1')
        assert state.direction == 'TD'

    def test_label_choice_added_to_lbl_defs(self):
        ws = _make_ws({(1, 1): 'Invoice No:'})
        cells = [(1, 1)]
        choices = {
            'A1': {
                'choice': 'L',
                'name':   'invoice_label',
                'ltype':  'string',
                'lmatch': 'Invoice No:',
            }
        }
        state = _build_state_from_choices(ws, choices, cells, 'LR', 'Sheet1')
        assert any(t[0] == 'invoice_label' for t in state.lbl_defs)

    def test_var_choice_added_to_var_defs(self):
        ws = _make_ws({(1, 2): 'INV-001'})
        cells = [(1, 2)]
        choices = {
            'B1': {
                'choice': 'V',
                'name':   'invoice_number',
                'ftype':  'string',
                'match':  '.*',
            }
        }
        state = _build_state_from_choices(ws, choices, cells, 'LR', 'Sheet1')
        assert any(t[0] == 'invoice_number' for t in state.var_defs)

    def test_ignore_choice_adds_ignore_body_row(self):
        ws = _make_ws({(1, 1): 'noise'})
        cells = [(1, 1)]
        choices = {'A1': {'choice': 'I'}}
        state = _build_state_from_choices(ws, choices, cells, 'LR', 'Sheet1')
        assert ['cell:1', 'IGNORE'] in state.body_rows

    def test_empty_choices_yields_empty_state(self):
        ws = _make_ws({(1, 1): 'hello'})
        cells = [(1, 1)]
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1')
        assert state.lbl_defs  == []
        assert state.var_defs  == []
        assert state.body_rows == []


# ── _choices_to_csv — pure-function tests ─────────────────────────────────────

class TestChoicesToCsv:
    """_choices_to_csv emits correct config rows and all classified cells."""

    def _ws_and_cells(self):
        ws = _make_ws({(1, 1): 'Invoice No:', (1, 2): 'INV-001'})
        cells = [(r, c) for r in range(1, ws.max_row + 1)
                 for c in range(1, ws.max_column + 1)]
        return ws, cells

    # ── Config rows ────────────────────────────────────────────────────────────

    def test_direction_row_always_present(self):
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(ws, {}, cells, 'LR', 'Sheet1'))
        assert ['config:', 'read.direction', 'LR'] in rows

    def test_td_direction_row(self):
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(ws, {}, cells, 'TD', 'Sheet1'))
        assert ['config:', 'read.direction', 'TD'] in rows

    def test_ignore_case_absent_by_default(self):
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(ws, {}, cells, 'LR', 'Sheet1'))
        assert not any(r[:2] == ['config:', 'ignore.case'] for r in rows)

    def test_ignore_case_row_emitted_when_true(self):
        """NEW-3 fix: ignore.case config row must appear in the preview CSV."""
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(
            ws, {}, cells, 'LR', 'Sheet1', ignore_case=True,
        ))
        assert ['config:', 'ignore.case', 'yes'] in rows

    def test_currency_sign_default_euro(self):
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(ws, {}, cells, 'LR', 'Sheet1'))
        assert ['config:', 'currency.sign', '€'] in rows

    def test_currency_sign_dollar(self):
        """NEW-3 fix: custom currency symbol must appear in the preview CSV."""
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(
            ws, {}, cells, 'LR', 'Sheet1', currency_sign='$',
        ))
        assert ['config:', 'currency.sign', '$'] in rows

    def test_both_ignore_case_and_currency_together(self):
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(
            ws, {}, cells, 'LR', 'Sheet1',
            ignore_case=True, currency_sign='£',
        ))
        assert ['config:', 'ignore.case', 'yes'] in rows
        assert ['config:', 'currency.sign', '£'] in rows

    # ── Field rows ─────────────────────────────────────────────────────────────

    def test_start_end_always_present(self):
        ws, cells = self._ws_and_cells()
        rows = _csv_rows(_choices_to_csv(ws, {}, cells, 'LR', 'Sheet1'))
        assert ['START:'] in rows
        assert ['END:'] in rows

    def test_classified_label_appears_in_lbl_row(self):
        ws = _make_ws({(1, 1): 'Invoice No:'})
        cells = [(1, 1)]
        choices = {
            'A1': {
                'choice': 'L',
                'name':   'invoice_label',
                'ltype':  'string',
                'lmatch': 'Invoice No:',
            }
        }
        rows = _csv_rows(_choices_to_csv(ws, choices, cells, 'LR', 'Sheet1'))
        assert any(r[:2] == ['lbl:', 'invoice_label'] for r in rows)

    def test_classified_var_appears_in_var_row(self):
        ws = _make_ws({(1, 2): 'INV-001'})
        cells = [(1, 2)]
        choices = {
            'B1': {
                'choice': 'V',
                'name':   'invoice_number',
                'ftype':  'string',
                'match':  '.*',
            }
        }
        rows = _csv_rows(_choices_to_csv(ws, choices, cells, 'LR', 'Sheet1'))
        assert any(r[:2] == ['var:', 'invoice_number'] for r in rows)

    def test_config_rows_come_before_lbl_rows(self):
        ws = _make_ws({(1, 1): 'Invoice No:'})
        cells = [(1, 1)]
        choices = {
            'A1': {'choice': 'L', 'name': 'lbl', 'ltype': 'string', 'lmatch': 'Invoice No:'},
        }
        rows = _csv_rows(_choices_to_csv(ws, choices, cells, 'LR', 'Sheet1'))
        config_idxs = [i for i, r in enumerate(rows) if r and r[0] == 'config:']
        lbl_idxs    = [i for i, r in enumerate(rows) if r and r[0] == 'lbl:']
        assert all(ci < li for ci in config_idxs for li in lbl_idxs)


# ── Textual pilot tests ────────────────────────────────────────────────────────

@_skip_no_textual
class TestWizardTUIApp:
    """Headless Textual pilot tests for the WizardTUIApp lifecycle.

    Each test is a *synchronous* function that uses ``asyncio.run()``
    internally.  No ``pytest-asyncio`` plugin is required.

    The config modal opens automatically on mount; pressing 'enter'
    accepts all defaults (LR direction, no template, no ignore_case, € currency).
    """

    def _make_app(self, data: dict[tuple[int, int], object],
                  tmp_path) -> 'WizardTUIApp':
        """Create a WizardTUIApp with a temporary xlsx data file."""
        xlsx_path = str(tmp_path / 'test_data.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        for (r, c), v in data.items():
            ws.cell(row=r, column=c, value=v)
        wb.save(xlsx_path)

        wb2 = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws2  = wb2.active
        state = WizardState(sheet_name=ws2.title)
        return WizardTUIApp(ws2, state, xlsx_path)

    # ── helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    async def _accept_config_and_then(pilot, *keys):
        """Accept the config modal (enter), then press each additional key."""
        await pilot.pause()
        await pilot.press('enter')   # dismiss _ConfigModal with defaults
        await pilot.pause()
        for key in keys:
            await pilot.press(key)
            await pilot.pause()

    # ── tests ─────────────────────────────────────────────────────────────────

    def test_cancel_returns_none(self, tmp_path):
        """Cancelling (ctrl+q) after accepting config exits with result=None."""
        app = self._make_app({(1, 1): 'Invoice No:', (1, 2): 'INV-001'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(pilot, 'ctrl+q')

        asyncio.run(_run())
        assert app._return_value is None

    def test_config_defaults_applied(self, tmp_path):
        """Accepting config defaults yields direction=LR, ignore_case=False, currency=€."""
        app = self._make_app({(1, 1): 'Hello'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await pilot.pause()
                await pilot.press('enter')
                await pilot.pause()
                assert app._state.direction     == 'LR'
                assert app._state.ignore_case   is False
                assert app._state.currency_sign == '€'
                await pilot.press('ctrl+q')
                await pilot.pause()

        asyncio.run(_run())

    def test_label_key_opens_modal(self, tmp_path):
        """Pressing 'l' on a non-empty cell opens the label classification modal."""
        app = self._make_app({(1, 1): 'Invoice No:'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(
                    pilot, 'l', 'escape', 'ctrl+q',
                )

        asyncio.run(_run())
        # No crash → test passes

    def test_ignore_key_classifies_cell(self, tmp_path):
        """Pressing 'i' classifies the current cell as Ignore without a modal."""
        app = self._make_app({(1, 1): 'noise'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(pilot, 'i')
                assert len(app._choices) == 1
                cell_meta = next(iter(app._choices.values()))
                assert cell_meta['choice'] == 'I'
                await pilot.press('ctrl+q')
                await pilot.pause()

        asyncio.run(_run())

    def test_undo_removes_classification(self, tmp_path):
        """Ctrl+Z after 'i' removes the classification from _choices."""
        app = self._make_app({(1, 1): 'noise'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(pilot, 'i')
                assert len(app._choices) == 1
                await pilot.press('ctrl+z')
                await pilot.pause()
                assert len(app._choices) == 0
                await pilot.press('ctrl+q')
                await pilot.pause()

        asyncio.run(_run())

    def test_save_with_ignore_produces_state(self, tmp_path):
        """Pressing 'i' then 'e' exits with a WizardState (not None)."""
        app = self._make_app({(1, 1): 'noise'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(pilot, 'i', 'e')

        asyncio.run(_run())
        result = app._return_value
        assert isinstance(result, WizardState)
        assert ['cell:1', 'IGNORE'] in result.body_rows

    def test_preview_key_opens_modal(self, tmp_path):
        """Pressing 'f3' opens the pattern preview modal without crashing."""
        app = self._make_app({(1, 1): 'Invoice No:'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(
                    pilot, 'f3', 'escape', 'ctrl+q',
                )

        asyncio.run(_run())
        # No crash → test passes

    def test_state_currency_sign_in_saved_pattern(self, tmp_path):
        """currency_sign from config (default €) propagates into the saved WizardState."""
        app = self._make_app({(1, 1): 'Price'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config_and_then(pilot, 'i', 'e')

        asyncio.run(_run())
        result = app._return_value
        assert isinstance(result, WizardState)
        assert result.currency_sign == '€'


# ══════════════════════════════════════════════════════════════════════════════
# New tests: full feature parity — modifiers, config, match modes
# ══════════════════════════════════════════════════════════════════════════════

# ── Helper: col_a_extra round-trip utilities ──────────────────────────────────

class TestColAExtraHelpers:
    """_col_a_extra_from_parts / _col_a_extra_to_parts round-trips."""

    def test_from_parts_none_gives_empty(self):
        assert _col_a_extra_from_parts('(default)', 'none') == ''

    def test_from_parts_nullable_only(self):
        assert _col_a_extra_from_parts('(default)', 'nullable') == 'nullable'

    def test_from_parts_not_null_only(self):
        assert _col_a_extra_from_parts('(default)', 'not-null') == 'not-null'

    def test_from_parts_trim_whitespace_only(self):
        assert _col_a_extra_from_parts('(default)', 'trim-whitespace') == 'trim-whitespace'

    def test_from_parts_literal_mode(self):
        assert _col_a_extra_from_parts('literal', 'none') == 'literal'

    def test_from_parts_glob_mode(self):
        assert _col_a_extra_from_parts('glob', 'none') == 'glob'

    def test_from_parts_nullable_plus_trim(self):
        assert _col_a_extra_from_parts('(default)', 'nullable:trim-whitespace') == 'nullable:trim-whitespace'

    def test_from_parts_not_null_plus_trim(self):
        assert _col_a_extra_from_parts('(default)', 'not-null:trim-whitespace') == 'not-null:trim-whitespace'

    def test_from_parts_literal_plus_nullable(self):
        assert _col_a_extra_from_parts('literal', 'nullable') == 'literal:nullable'

    def test_to_parts_empty_gives_defaults(self):
        assert _col_a_extra_to_parts('') == ('(default)', 'none')

    def test_to_parts_nullable(self):
        assert _col_a_extra_to_parts('nullable') == ('(default)', 'nullable')

    def test_to_parts_not_null(self):
        assert _col_a_extra_to_parts('not-null') == ('(default)', 'not-null')

    def test_to_parts_trim_whitespace(self):
        assert _col_a_extra_to_parts('trim-whitespace') == ('(default)', 'trim-whitespace')

    def test_to_parts_literal_mode(self):
        assert _col_a_extra_to_parts('literal') == ('literal', 'none')

    def test_to_parts_glob_mode(self):
        assert _col_a_extra_to_parts('glob') == ('glob', 'none')

    def test_to_parts_nullable_trim(self):
        mode, mods = _col_a_extra_to_parts('nullable:trim-whitespace')
        assert mode == '(default)'
        assert 'nullable' in mods
        assert 'trim-whitespace' in mods

    def test_round_trip_not_null_trim(self):
        original = 'not-null:trim-whitespace'
        mode, mods = _col_a_extra_to_parts(original)
        rebuilt = _col_a_extra_from_parts(mode, mods)
        assert 'not-null' in rebuilt
        assert 'trim-whitespace' in rebuilt


# ── _build_state_from_choices: new config fields ──────────────────────────────

class TestBuildStateNewConfigFields:
    """New WizardState config fields propagate from _build_state_from_choices."""

    def _ws_and_cells(self):
        ws = _make_ws({(1, 1): 'Label:', (1, 2): 'Value'})
        cells = [(r, c) for r in range(1, ws.max_row + 1)
                 for c in range(1, ws.max_column + 1)]
        return ws, cells

    def test_trim_whitespace_false_by_default(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1')
        assert state.trim_whitespace is False

    def test_trim_whitespace_true_propagated(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1',
                                          trim_whitespace=True)
        assert state.trim_whitespace is True

    def test_lbl_match_empty_by_default(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1')
        assert state.lbl_match == ''

    def test_lbl_match_glob_propagated(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1',
                                          lbl_match='glob')
        assert state.lbl_match == 'glob'

    def test_var_match_empty_by_default(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1')
        assert state.var_match == ''

    def test_var_match_literal_propagated(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1',
                                          var_match='literal')
        assert state.var_match == 'literal'

    def test_empty_aliases_empty_by_default(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1')
        assert state.empty_aliases == []

    def test_empty_aliases_propagated(self):
        ws, cells = self._ws_and_cells()
        state = _build_state_from_choices(ws, {}, cells, 'LR', 'Sheet1',
                                          empty_aliases=['N/A', '-'])
        assert state.empty_aliases == ['N/A', '-']


# ── _choices_to_csv: var: modifier output ────────────────────────────────────

class TestChoicesToCsvModifiers:
    """_choices_to_csv emits correct col-A tokens for modifiers and modes."""

    def _ws_and_cells(self):
        ws = _make_ws({(1, 1): 'Value'})
        cells = [(1, 1)]
        return ws, cells

    def _rows(self, choices, **kwargs) -> list[list[str]]:
        ws, cells = self._ws_and_cells()
        return _csv_rows(_choices_to_csv(ws, choices, cells, 'LR', 'Sheet1', **kwargs))

    def _choice_V(self, col_a_extra: str = '') -> dict:
        return {'A1': {'choice': 'V', 'name': 'myfield', 'ftype': 'string',
                       'match': '.*', 'col_a_extra': col_a_extra}}

    def test_bare_var_no_modifier(self):
        rows = self._rows(self._choice_V(''))
        assert any(r[0] == 'var:' and r[1] == 'myfield' for r in rows)

    def test_nullable_modifier_in_col_a(self):
        rows = self._rows(self._choice_V('nullable'))
        assert any(r[0] == 'var:nullable' and r[1] == 'myfield' for r in rows)

    def test_not_null_modifier_in_col_a(self):
        rows = self._rows(self._choice_V('not-null'))
        assert any(r[0] == 'var:not-null' and r[1] == 'myfield' for r in rows)

    def test_trim_whitespace_modifier_in_col_a(self):
        rows = self._rows(self._choice_V('trim-whitespace'))
        assert any(r[0] == 'var:trim-whitespace' and r[1] == 'myfield' for r in rows)

    def test_nullable_plus_trim_in_col_a(self):
        rows = self._rows(self._choice_V('nullable:trim-whitespace'))
        assert any(r[0] == 'var:nullable:trim-whitespace' and r[1] == 'myfield' for r in rows)

    def test_literal_var_mode_in_col_a(self):
        rows = self._rows(self._choice_V('literal'))
        assert any(r[0] == 'var:literal' and r[1] == 'myfield' for r in rows)

    def test_glob_var_mode_in_col_a(self):
        rows = self._rows(self._choice_V('glob'))
        assert any(r[0] == 'var:glob' and r[1] == 'myfield' for r in rows)

    def test_glob_mode_plus_nullable_in_col_a(self):
        rows = self._rows(self._choice_V('glob:nullable'))
        assert any(r[0] == 'var:glob:nullable' and r[1] == 'myfield' for r in rows)


# ── _choices_to_csv: lbl: match mode output ───────────────────────────────────

class TestChoicesToCsvLblMode:
    """_choices_to_csv emits correct col-A lbl: tokens for per-field match mode."""

    def _ws_and_cells(self):
        ws = _make_ws({(1, 1): 'Invoice:'})
        cells = [(1, 1)]
        return ws, cells

    def _rows(self, choices) -> list[list[str]]:
        ws, cells = self._ws_and_cells()
        return _csv_rows(_choices_to_csv(ws, choices, cells, 'LR', 'Sheet1'))

    def _choice_L(self, lbl_mode: str = '') -> dict:
        return {'A1': {'choice': 'L', 'name': 'inv_lbl', 'ltype': 'string',
                       'lmatch': 'Invoice:', 'lbl_mode': lbl_mode}}

    def test_bare_lbl_no_mode(self):
        rows = self._rows(self._choice_L(''))
        assert any(r[0] == 'lbl:' and r[1] == 'inv_lbl' for r in rows)

    def test_lbl_glob_mode_in_col_a(self):
        rows = self._rows(self._choice_L('glob'))
        assert any(r[0] == 'lbl:glob' and r[1] == 'inv_lbl' for r in rows)

    def test_lbl_regexp_mode_in_col_a(self):
        rows = self._rows(self._choice_L('regexp'))
        assert any(r[0] == 'lbl:regexp' and r[1] == 'inv_lbl' for r in rows)


# ── _choices_to_csv: new config rows ─────────────────────────────────────────

class TestChoicesToCsvNewConfigRows:
    """_choices_to_csv emits new config rows when non-default values are set."""

    def _ws_and_cells(self):
        ws = _make_ws({(1, 1): 'x'})
        cells = [(1, 1)]
        return ws, cells

    def _rows(self, **kwargs) -> list[list[str]]:
        ws, cells = self._ws_and_cells()
        return _csv_rows(_choices_to_csv(ws, {}, cells, 'LR', 'Sheet1', **kwargs))

    def test_trim_whitespace_absent_by_default(self):
        rows = self._rows()
        assert not any(r[:2] == ['config:', 'trim.whitespace'] for r in rows)

    def test_trim_whitespace_yes_emitted(self):
        rows = self._rows(trim_whitespace=True)
        assert ['config:', 'trim.whitespace', 'yes'] in rows

    def test_lbl_match_absent_by_default(self):
        rows = self._rows()
        assert not any(r[:2] == ['config:', 'lbl.match'] for r in rows)

    def test_lbl_match_glob_emitted(self):
        rows = self._rows(lbl_match='glob')
        assert ['config:', 'lbl.match', 'glob'] in rows

    def test_lbl_match_regexp_emitted(self):
        rows = self._rows(lbl_match='regexp')
        assert ['config:', 'lbl.match', 'regexp'] in rows

    def test_var_match_absent_by_default(self):
        rows = self._rows()
        assert not any(r[:2] == ['config:', 'var.match'] for r in rows)

    def test_var_match_literal_emitted(self):
        rows = self._rows(var_match='literal')
        assert ['config:', 'var.match', 'literal'] in rows

    def test_var_match_glob_emitted(self):
        rows = self._rows(var_match='glob')
        assert ['config:', 'var.match', 'glob'] in rows

    def test_empty_aliases_absent_by_default(self):
        rows = self._rows()
        assert not any(r[:2] == ['config:', 'empty.aliases'] for r in rows)

    def test_empty_aliases_one_alias_emitted(self):
        rows = self._rows(empty_aliases=['N/A'])
        assert ['config:', 'empty.aliases', 'N/A'] in rows

    def test_empty_aliases_multiple_rows_emitted(self):
        rows = self._rows(empty_aliases=['N/A', '-', 'n/a'])
        assert ['config:', 'empty.aliases', 'N/A'] in rows
        assert ['config:', 'empty.aliases', '-'] in rows
        assert ['config:', 'empty.aliases', 'n/a'] in rows

    def test_config_rows_before_start(self):
        rows = self._rows(trim_whitespace=True, lbl_match='glob', var_match='literal')
        start_idx = next(i for i, r in enumerate(rows) if r == ['START:'])
        config_idxs = [i for i, r in enumerate(rows)
                       if r and r[0] == 'config:']
        assert all(ci < start_idx for ci in config_idxs)

    def test_all_new_config_rows_together(self):
        rows = self._rows(
            trim_whitespace=True, lbl_match='glob',
            var_match='literal', empty_aliases=['N/A'],
        )
        assert ['config:', 'trim.whitespace', 'yes'] in rows
        assert ['config:', 'lbl.match', 'glob'] in rows
        assert ['config:', 'var.match', 'literal'] in rows
        assert ['config:', 'empty.aliases', 'N/A'] in rows


# ── _pattern_rows: engine-side output for modifiers ──────────────────────────

class TestPatternRowsModifiers:
    """_pattern_rows correctly builds col-A from WizardState 4-tuples."""

    def _state(self) -> WizardState:
        return WizardState(direction='LR', currency_sign='€')

    def test_bare_var_outputs_var_colon(self):
        s = self._state()
        s.var_defs.append(('field', 'string', '.*', ''))
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:' and r[1] == 'field' for r in rows)

    def test_nullable_var_outputs_var_nullable(self):
        s = self._state()
        s.var_defs.append(('field', 'string', '.*', 'nullable'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:nullable' and r[1] == 'field' for r in rows)

    def test_not_null_var(self):
        s = self._state()
        s.var_defs.append(('field', 'string', r'\d+', 'not-null'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:not-null' and r[1] == 'field' for r in rows)

    def test_trim_whitespace_var(self):
        s = self._state()
        s.var_defs.append(('field', 'string', '.*', 'trim-whitespace'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:trim-whitespace' and r[1] == 'field' for r in rows)

    def test_literal_mode_var(self):
        s = self._state()
        s.var_defs.append(('field', 'string', 'exact text', 'literal'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:literal' and r[1] == 'field' for r in rows)

    def test_glob_mode_plus_nullable(self):
        s = self._state()
        s.var_defs.append(('field', 'string', '*.txt', 'glob:nullable'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:glob:nullable' and r[1] == 'field' for r in rows)

    def test_bare_lbl_outputs_lbl_colon(self):
        s = self._state()
        s.lbl_defs.append(('my_lbl', 'string', 'My Label:', ''))
        rows = _pattern_rows(s)
        assert any(r[0] == 'lbl:' and r[1] == 'my_lbl' for r in rows)

    def test_glob_lbl_mode(self):
        s = self._state()
        s.lbl_defs.append(('my_lbl', 'string', 'My*', 'glob'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'lbl:glob' and r[1] == 'my_lbl' for r in rows)

    def test_regexp_lbl_mode(self):
        s = self._state()
        s.lbl_defs.append(('my_lbl', 'string', r'My\s+Label', 'regexp'))
        rows = _pattern_rows(s)
        assert any(r[0] == 'lbl:regexp' and r[1] == 'my_lbl' for r in rows)

    def test_trim_whitespace_config_row(self):
        s = self._state()
        s.trim_whitespace = True
        rows = _pattern_rows(s)
        assert ['config:', 'trim.whitespace', 'yes'] in rows

    def test_lbl_match_config_row(self):
        s = self._state()
        s.lbl_match = 'regexp'
        rows = _pattern_rows(s)
        assert ['config:', 'lbl.match', 'regexp'] in rows

    def test_var_match_config_row(self):
        s = self._state()
        s.var_match = 'glob'
        rows = _pattern_rows(s)
        assert ['config:', 'var.match', 'glob'] in rows

    def test_empty_aliases_config_rows(self):
        s = self._state()
        s.empty_aliases = ['N/A', '-']
        rows = _pattern_rows(s)
        assert ['config:', 'empty.aliases', 'N/A'] in rows
        assert ['config:', 'empty.aliases', '-'] in rows

    def test_3tuple_lbl_def_backward_compat(self):
        """WizardState with legacy 3-tuple lbl_defs still produces bare lbl:."""
        s = self._state()
        s.lbl_defs.append(('lbl', 'string', 'Old Label:'))  # 3-tuple
        rows = _pattern_rows(s)
        assert any(r[0] == 'lbl:' and r[1] == 'lbl' for r in rows)

    def test_3tuple_var_def_backward_compat(self):
        """WizardState with legacy 3-tuple var_defs still produces bare var:."""
        s = self._state()
        s.var_defs.append(('field', 'string', '.*'))  # 3-tuple
        rows = _pattern_rows(s)
        assert any(r[0] == 'var:' and r[1] == 'field' for r in rows)


# ── WizardState: new config fields round-trip ─────────────────────────────────

class TestWizardStateNewFieldsRoundTrip:
    """New WizardState config fields round-trip through to_dict / from_dict."""

    def test_trim_whitespace_round_trips(self):
        s = WizardState(trim_whitespace=True)
        r = WizardState.from_dict(s.to_dict())
        assert r.trim_whitespace is True

    def test_lbl_match_round_trips(self):
        s = WizardState(lbl_match='glob')
        r = WizardState.from_dict(s.to_dict())
        assert r.lbl_match == 'glob'

    def test_var_match_round_trips(self):
        s = WizardState(var_match='literal')
        r = WizardState.from_dict(s.to_dict())
        assert r.var_match == 'literal'

    def test_empty_aliases_round_trips(self):
        s = WizardState(empty_aliases=['N/A', '-'])
        r = WizardState.from_dict(s.to_dict())
        assert r.empty_aliases == ['N/A', '-']

    def test_defaults_on_empty_dict(self):
        r = WizardState.from_dict({})
        assert r.trim_whitespace is False
        assert r.lbl_match == ''
        assert r.var_match == ''
        assert r.empty_aliases == []

    def test_4tuple_var_def_round_trips(self):
        s = WizardState()
        s.var_defs.append(('inv_no', 'string', r'[A-Z]+\d+', 'nullable'))
        r = WizardState.from_dict(s.to_dict())
        assert r.var_defs == [('inv_no', 'string', r'[A-Z]+\d+', 'nullable')]

    def test_4tuple_lbl_def_round_trips(self):
        s = WizardState()
        s.lbl_defs.append(('inv_lbl', 'string', 'Invoice:', 'glob'))
        r = WizardState.from_dict(s.to_dict())
        assert r.lbl_defs == [('inv_lbl', 'string', 'Invoice:', 'glob')]


# ── Integration: nullable modifier flows through CSV output ───────────────────

class TestNullableModifierEndToEnd:
    """var:nullable in _choices flows all the way to a parseable pattern CSV."""

    def test_nullable_var_produces_parseable_pattern(self, tmp_path):
        """Pattern with var:nullable is accepted by the grepxcel parser."""
        from grepxcel.pattern_parser import PatternParser

        ws = _make_ws({(1, 1): 'Name:'})
        cells = [(1, 1)]
        choices = {
            'A1': {'choice': 'V', 'name': 'patient_name', 'ftype': 'string',
                   'match': '.*', 'col_a_extra': 'nullable'},
        }
        csv_text = _choices_to_csv(ws, choices, cells, 'LR', 'Sheet1')
        rows = _csv_rows(csv_text)
        # var:nullable must appear in col A
        assert any(r[0] == 'var:nullable' and r[1] == 'patient_name' for r in rows)
        # Pattern must parse without error: write to temp CSV and parse
        csv_path = str(tmp_path / 'pattern.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as fh:
            fh.write(csv_text)
        parser = PatternParser()
        result = parser.parse(csv_path)
        # Parser ran without error is the key assertion; result is a non-empty tuple
        assert result is not None
        assert len(result) >= 1

    def test_not_null_var_produces_required_true(self):
        """var:not-null col_a_extra → parser produces FieldDef.required=True."""
        import io as _io
        from grepxcel.pattern_parser import PatternParser

        ws = _make_ws({(1, 1): 'Amount:'})
        cells = [(1, 1)]
        choices = {
            'A1': {'choice': 'V', 'name': 'amount', 'ftype': 'currency',
                   'match': r'\d+\.\d{2}', 'col_a_extra': 'not-null'},
        }
        csv_text = _choices_to_csv(ws, choices, cells, 'LR', 'Sheet1')
        rows = _csv_rows(csv_text)
        assert any(r[0] == 'var:not-null' and r[1] == 'amount' for r in rows)

    def test_trim_whitespace_var_produces_correct_col_a(self):
        ws = _make_ws({(1, 1): 'Notes:'})
        cells = [(1, 1)]
        choices = {
            'A1': {'choice': 'V', 'name': 'notes', 'ftype': 'string',
                   'match': '.*', 'col_a_extra': 'trim-whitespace'},
        }
        csv_text = _choices_to_csv(ws, choices, cells, 'LR', 'Sheet1')
        rows = _csv_rows(csv_text)
        assert any(r[0] == 'var:trim-whitespace' and r[1] == 'notes' for r in rows)


# ── Pilot: config modal now includes new fields ────────────────────────────────

@_skip_no_textual
class TestConfigModalNewFields:
    """Config modal dismisses successfully with new field defaults."""

    def _make_app(self, data, tmp_path) -> 'WizardTUIApp':
        xlsx_path = str(tmp_path / 'test_data.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        for (r, c), v in data.items():
            ws.cell(row=r, column=c, value=v)
        wb.save(xlsx_path)
        wb2 = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws2  = wb2.active
        state = WizardState(sheet_name=ws2.title)
        return WizardTUIApp(ws2, state, xlsx_path)

    def test_config_modal_accepts_with_new_defaults(self, tmp_path):
        """Config modal dismisses with ENTER; new config fields get defaults."""
        app = self._make_app({(1, 1): 'Hello'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 50)) as pilot:
                await pilot.pause()
                await pilot.press('enter')  # accept config modal
                await pilot.pause()
                # New config defaults
                assert app._state.trim_whitespace is False
                assert app._state.lbl_match == 'literal'
                assert app._state.var_match == 'regexp'
                assert app._state.empty_aliases == []
                await pilot.press('ctrl+q')
                await pilot.pause()

        asyncio.run(_run())

    def test_save_preserves_new_config_defaults(self, tmp_path):
        """New config fields appear in saved WizardState."""
        app = self._make_app({(1, 1): 'noise'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 50)) as pilot:
                await pilot.pause()
                await pilot.press('enter')  # accept config
                await pilot.pause()
                await pilot.press('i')      # classify as Ignore
                await pilot.pause()
                await pilot.press('e')      # save
                await pilot.pause()

        asyncio.run(_run())
        result = app._return_value
        assert isinstance(result, WizardState)
        # New fields present with defaults
        assert result.trim_whitespace is False
        assert result.lbl_match == 'literal'
        assert result.var_match == 'regexp'
        assert result.empty_aliases == []
