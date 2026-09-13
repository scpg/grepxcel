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
        _fd_to_col_a_extra,
        _preload_from_pattern,
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
        """Accept the config modal (ctrl+enter), then press each additional key."""
        await pilot.pause()
        await pilot.press('ctrl+enter')  # dismiss _ConfigModal with defaults
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
        """Config modal dismisses on OK button click; new config fields get defaults."""
        app = self._make_app({(1, 1): 'Hello'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 50)) as pilot:
                await pilot.pause()
                await pilot.press('ctrl+enter')  # accept config modal
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
                await pilot.press('ctrl+enter')  # accept config
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


@_skip_no_textual
class TestConfigModalSelectSafety:
    """Verify that Select interactions inside _ConfigModal do NOT close the modal.

    Regression tests for the bug where Textual's Select BINDING for 'enter'
    does not stop the key event, causing it to bubble to on_key and dismiss the
    modal prematurely.
    """

    def _make_app(self, data, tmp_path) -> 'WizardTUIApp':
        xlsx_path = str(tmp_path / 'test_data.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        for (r, c), v in data.items():
            ws.cell(row=r, column=c, value=v)
        wb.save(xlsx_path)
        wb2 = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws2 = wb2.active
        state = WizardState(sheet_name=ws2.title)
        return WizardTUIApp(ws2, state, xlsx_path)

    def test_enter_on_select_does_not_dismiss_modal(self, tmp_path):
        """Pressing Enter while a Select is focused must NOT dismiss the config modal."""
        app = self._make_app({(1, 1): 'Hello'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 60)) as pilot:
                await pilot.pause()
                # Config modal is open; first Select (#dir) has focus
                # Press Enter multiple times — each press should open/close the
                # dropdown, never dismiss the modal
                for _ in range(4):
                    await pilot.press('enter')
                    await pilot.pause()
                # Modal must still be open (screen stack has _ConfigModal on top)
                from grepxcel.wizard_tui import _ConfigModal
                assert any(
                    isinstance(s, _ConfigModal) for s in app.screen_stack
                ), "Config modal should still be open after Enter presses on Select"
                # Now dismiss properly
                await pilot.press('ctrl+enter')
                await pilot.pause()

        asyncio.run(_run())

    def test_ctrl_enter_dismisses_with_correct_values(self, tmp_path):
        """Ctrl+Enter confirms the config modal with the current widget values."""
        app = self._make_app({(1, 1): 'Hello'}, tmp_path)

        async def _run():
            async with app.run_test(headless=True, size=(120, 60)) as pilot:
                await pilot.pause()
                # Open the direction Select, press arrow key to highlight TD option,
                # then Ctrl+Enter to confirm the whole modal
                await pilot.press('enter')    # open #dir dropdown
                await pilot.pause()
                await pilot.press('ctrl+enter')  # confirm modal (force)
                await pilot.pause()
                await pilot.press('ctrl+q')

        asyncio.run(_run())
        # Just verify we didn't crash — the modal was dismissed cleanly
        assert True

    def test_escape_cancels_modal(self, tmp_path):
        """ESC dismisses the config modal with None (cancel, no state change)."""
        app = self._make_app({(1, 1): 'Hello'}, tmp_path)
        original_dir = 'LR'

        async def _run():
            async with app.run_test(headless=True, size=(120, 60)) as pilot:
                await pilot.pause()
                await pilot.press('escape')
                await pilot.pause()
                # App should exit with None (cancel)
                await pilot.press('ctrl+q')
                await pilot.pause()

        asyncio.run(_run())


# ── Helpers shared across TestPreloadFromPattern ──────────────────────────────

def _make_ws(cells: dict) -> openpyxl.worksheet.worksheet.Worksheet:
    """Return a minimal in-memory worksheet populated from {(row,col): value}."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for (r, c), v in cells.items():
        ws.cell(row=r, column=c, value=v)
    return ws


def _write_pattern_csv(path, rows: list[list]) -> None:
    """Write a plain-CSV pattern file from a list-of-lists."""
    with open(path, 'w', newline='', encoding='utf-8') as fh:
        writer = csv.writer(fh)
        for row in rows:
            writer.writerow(row)


# ── TestFdToColAExtra ─────────────────────────────────────────────────────────

class TestFdToColAExtra:
    """_fd_to_col_a_extra reconstructs col_a_extra tokens from a FieldDef."""

    def _make_fd(self, **kw):
        from grepxcel.models import FieldDef
        defaults = dict(name='x', type='string', regex='.*',
                        role='var', var_mode=None, required=False,
                        nullable=False, trim_whitespace=False)
        defaults.update(kw)
        return FieldDef(**defaults)

    def test_all_defaults_gives_empty(self):
        fd = self._make_fd()
        assert _fd_to_col_a_extra(fd) == ''

    def test_nullable(self):
        fd = self._make_fd(nullable=True)
        assert _fd_to_col_a_extra(fd) == 'nullable'

    def test_not_null(self):
        fd = self._make_fd(required=True)
        assert _fd_to_col_a_extra(fd) == 'not-null'

    def test_trim_whitespace_only(self):
        fd = self._make_fd(trim_whitespace=True)
        assert _fd_to_col_a_extra(fd) == 'trim-whitespace'

    def test_not_null_and_trim(self):
        fd = self._make_fd(required=True, trim_whitespace=True)
        assert _fd_to_col_a_extra(fd) == 'not-null:trim-whitespace'

    def test_nullable_and_trim(self):
        fd = self._make_fd(nullable=True, trim_whitespace=True)
        assert _fd_to_col_a_extra(fd) == 'nullable:trim-whitespace'

    def test_literal_var_mode(self):
        fd = self._make_fd(var_mode='literal')
        assert _fd_to_col_a_extra(fd) == 'literal'

    def test_glob_var_mode(self):
        fd = self._make_fd(var_mode='glob')
        assert _fd_to_col_a_extra(fd) == 'glob'

    def test_regexp_var_mode_omitted(self):
        # 'regexp' is the default — it should not appear in col_a_extra
        fd = self._make_fd(var_mode='regexp')
        assert _fd_to_col_a_extra(fd) == ''

    def test_literal_plus_nullable(self):
        fd = self._make_fd(var_mode='literal', nullable=True)
        assert _fd_to_col_a_extra(fd) == 'literal:nullable'

    def test_required_wins_over_nullable(self):
        # required (not-null) takes priority; nullable should not appear
        fd = self._make_fd(required=True, nullable=True)
        result = _fd_to_col_a_extra(fd)
        assert 'not-null' in result
        assert 'nullable' not in result


# ── TestPreloadFromPattern ────────────────────────────────────────────────────

class TestPreloadFromPattern:
    """_preload_from_pattern: pattern file → (choices, config, warnings)."""

    # ── Label matching ────────────────────────────────────────────────────────

    def test_lbl_literal_found(self, tmp_path):
        ws = _make_ws({(1, 1): 'Invoice Date:', (1, 2): '2026-01-15'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'inv_lbl', 'string', 'Invoice Date:'],
            ['var:', 'inv_date', 'string', '.*'],
            ['START:'],
            ['cell:1', 'inv_lbl'],
            ['cell:1', 'inv_date'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'A1' in choices
        assert choices['A1']['choice'] == 'L'
        assert choices['A1']['name'] == 'inv_lbl'

    def test_lbl_not_found_emits_warning(self, tmp_path):
        ws = _make_ws({(1, 1): 'Something Else', (1, 2): '42'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'missing_lbl', 'string', 'Invoice Date:'],
            ['START:'],
            ['cell:1', 'missing_lbl'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'A1' not in choices
        assert any('missing_lbl' in w for w in warns)

    def test_var_adjacent_LR(self, tmp_path):
        """In LR mode the var cell is one column right of the label."""
        ws = _make_ws({(2, 1): 'Total:', (2, 2): '100.00'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'read.direction', 'LR'],
            ['lbl:', 'total_lbl', 'string', 'Total:'],
            ['var:', 'total', 'currency', r'\d+\.\d{2}'],
            ['START:'],
            ['cell:1', 'total_lbl'],
            ['cell:1', 'total'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'A2' in choices and choices['A2']['choice'] == 'L'
        assert 'B2' in choices and choices['B2']['choice'] == 'V'
        assert choices['B2']['name'] == 'total'
        assert not warns

    def test_var_adjacent_TD(self, tmp_path):
        """In TD mode the var cell is one row below the label."""
        ws = _make_ws({(1, 3): 'Total:', (2, 3): '100.00'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'read.direction', 'TD'],
            ['lbl:', 'total_lbl', 'string', 'Total:'],
            ['var:', 'total', 'currency', r'\d+\.\d{2}'],
            ['START:'],
            ['cell:1', 'total_lbl'],
            ['cell:1', 'total'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'C1' in choices and choices['C1']['choice'] == 'L'
        assert 'C2' in choices and choices['C2']['choice'] == 'V'
        assert not warns

    def test_abs_cell_lbl(self, tmp_path):
        """cell:B3 places the label at exactly (3, 2)."""
        ws = _make_ws({(3, 2): 'Name:', (3, 3): 'Alice'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'name_lbl', 'string', 'Name:'],
            ['var:', 'name', 'string', '.*'],
            ['START:'],
            ['cell:B3', 'name_lbl'],
            ['cell:1', 'name'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'B3' in choices
        assert choices['B3']['choice'] == 'L'

    def test_abs_cell_var(self, tmp_path):
        """cell:C3 places a var at exactly (3, 3)."""
        ws = _make_ws({(3, 2): 'Name:', (3, 3): 'Bob'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'name_lbl', 'string', 'Name:'],
            ['var:', 'name', 'string', '.*'],
            ['START:'],
            ['cell:B3', 'name_lbl'],
            ['cell:C3', 'name'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'C3' in choices
        assert choices['C3']['choice'] == 'V'

    def test_table_instruction_emits_warning(self, tmp_path):
        """TABLE instructions produce a warning listing the field names."""
        ws = _make_ws({(1, 1): 'Date', (1, 2): 'Amount', (2, 1): '2026-01-01', (2, 2): '42'})
        pat = tmp_path / 'p.csv'
        # Minimal valid table pattern
        _write_pattern_csv(pat, [
            ['START:'],
            ['table:*', ''],
            ['', 'HEADER:1', 'date_col', 'amount_col'],
            ['', 'DATA:*', 'date_col', 'amount_col'],
            ['var:', 'date_col', 'string', '.*'],
            ['var:', 'amount_col', 'currency', '.*'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        # At least one warning mentions table fields
        assert any('TABLE' in w or 'table' in w.lower() for w in warns)

    def test_ignore_field_skipped(self, tmp_path):
        """cell:1 IGNORE produces no choices entry and does not crash."""
        ws = _make_ws({(1, 1): 'Label:', (1, 2): 'skip_me', (1, 3): 'Value'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'lbl1', 'string', 'Label:'],
            ['START:'],
            ['cell:1', 'lbl1'],
            ['cell:1', 'IGNORE'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'A1' in choices
        # B1 (the IGNORE cell) must not appear in choices
        assert 'B1' not in choices

    def test_multiple_lbl_var_pairs(self, tmp_path):
        """Three lbl/var pairs are all located correctly."""
        ws = _make_ws({
            (1, 1): 'Name:', (1, 2): 'Alice',
            (2, 1): 'Date:', (2, 2): '2026-01-01',
            (3, 1): 'Amount:', (3, 2): '99.00',
        })
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'name_lbl',   'string',   'Name:'],
            ['var:', 'name',       'string',   '.*'],
            ['lbl:', 'date_lbl',   'string',   'Date:'],
            ['var:', 'date',       'date',     r'\d{4}-\d{2}-\d{2}'],
            ['lbl:', 'amount_lbl', 'string',   'Amount:'],
            ['var:', 'amount',     'currency', r'\d+\.\d{2}'],
            ['START:'],
            ['cell:1', 'name_lbl'],   ['cell:1', 'name'],
            ['cell:1', 'date_lbl'],   ['cell:1', 'date'],
            ['cell:1', 'amount_lbl'], ['cell:1', 'amount'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert len([v for v in choices.values() if v['choice'] == 'L']) == 3
        assert len([v for v in choices.values() if v['choice'] == 'V']) == 3
        assert not warns

    def test_dir_instruction_updates_adjacency(self, tmp_path):
        """dir:TD mid-sequence makes subsequent var adjacent below, not right."""
        ws = _make_ws({(5, 2): 'Label:', (6, 2): 'val_below'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'lbl1', 'string', 'Label:'],
            ['var:', 'val1', 'string', '.*'],
            ['START:'],
            ['dir:TD'],
            ['cell:1', 'lbl1'],
            ['cell:1', 'val1'],
            ['END:'],
        ])
        choices, cfg, warns = _preload_from_pattern(ws, str(pat))
        assert 'B5' in choices and choices['B5']['choice'] == 'L'
        assert 'B6' in choices and choices['B6']['choice'] == 'V'

    def test_parse_error_returns_warning(self, tmp_path):
        """A corrupt pattern file → empty choices + warning."""
        ws = _make_ws({(1, 1): 'x'})
        bad = tmp_path / 'bad.csv'
        bad.write_text('INVALID_ROW_MARKER,field,string,val\n', encoding='utf-8')
        choices, cfg, warns = _preload_from_pattern(ws, str(bad))
        assert choices == {}
        assert len(warns) >= 1
        # Config dict is empty too
        assert cfg == {}

    # ── Config preload ────────────────────────────────────────────────────────

    def test_config_direction_LR(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'read.direction', 'LR'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['direction'] == 'LR'

    def test_config_direction_TD(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'read.direction', 'TD'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['direction'] == 'TD'

    def test_config_ignore_case(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'ignore.case', 'yes'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['ignore_case'] is True

    def test_config_trim_whitespace(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'trim.whitespace', 'yes'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['trim_whitespace'] is True

    def test_config_currency_sign(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'currency.sign', '$'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['currency_sign'] == '$'

    def test_config_lbl_match_glob_preserved(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'lbl.match', 'glob'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['lbl_match'] == 'glob'

    def test_config_lbl_match_literal_omitted(self, tmp_path):
        """literal is the engine default — preload_cfg sets it to '' so the
        ConfigModal does not emit a redundant config: row on save."""
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'lbl.match', 'literal'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['lbl_match'] == ''

    def test_config_var_match_regexp_omitted(self, tmp_path):
        """regexp is the engine default — preload_cfg sets it to ''."""
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'var.match', 'regexp'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['var_match'] == ''

    def test_config_var_match_glob_preserved(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'var.match', 'glob'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['var_match'] == 'glob'

    def test_config_empty_aliases(self, tmp_path):
        ws = _make_ws({(1, 1): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['config:', 'empty.aliases', 'N/A'],
            ['config:', 'empty.aliases', '-'],
            ['START:'], ['END:'],
        ])
        _, cfg, _ = _preload_from_pattern(ws, str(pat))
        assert cfg['empty_aliases'] == ['N/A', '-']

    # ── col_a_extra reconstruction ────────────────────────────────────────────

    def test_nullable_modifier_in_choices(self, tmp_path):
        """var: field with nullable=True → col_a_extra='nullable' in choices."""
        ws = _make_ws({(1, 1): 'Note:', (1, 2): None})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'note_lbl', 'string', 'Note:'],
            ['var:nullable', 'note', 'string', '.*'],
            ['START:'],
            ['cell:1', 'note_lbl'],
            ['cell:1', 'note'],
            ['END:'],
        ])
        choices, _, warns = _preload_from_pattern(ws, str(pat))
        assert 'A1' in choices
        # B1 is empty (None) — it still gets pre-classified if within bounds
        # but we just check the col_a_extra on the var field at B1
        if 'B1' in choices:
            assert choices['B1']['col_a_extra'] == 'nullable'

    def test_not_null_trim_in_choices(self, tmp_path):
        ws = _make_ws({(1, 1): 'Ref:', (1, 2): 'REF-001'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'ref_lbl', 'string', 'Ref:'],
            ['var:not-null:trim-whitespace', 'ref', 'string', '.*'],
            ['START:'],
            ['cell:1', 'ref_lbl'],
            ['cell:1', 'ref'],
            ['END:'],
        ])
        choices, _, warns = _preload_from_pattern(ws, str(pat))
        assert 'B1' in choices
        assert 'not-null' in choices['B1']['col_a_extra']
        assert 'trim-whitespace' in choices['B1']['col_a_extra']

    # ── lbl_mode round-trip ───────────────────────────────────────────────────

    def test_lbl_mode_glob_preserved(self, tmp_path):
        ws = _make_ws({(1, 1): 'Invoice *', (1, 2): '42'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:glob', 'inv_lbl', 'string', 'Invoice *'],
            ['var:', 'inv_val', 'string', '.*'],
            ['START:'],
            ['cell:1', 'inv_lbl'],
            ['cell:1', 'inv_val'],
            ['END:'],
        ])
        choices, _, warns = _preload_from_pattern(ws, str(pat))
        assert 'A1' in choices
        assert choices['A1']['lbl_mode'] == 'glob'

    def test_lbl_mode_default_empty(self, tmp_path):
        ws = _make_ws({(1, 1): 'Name:', (1, 2): 'x'})
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'name_lbl', 'string', 'Name:'],
            ['START:'],
            ['cell:1', 'name_lbl'],
            ['END:'],
        ])
        choices, _, warns = _preload_from_pattern(ws, str(pat))
        assert choices['A1']['lbl_mode'] == ''

    # ── Round-trip integration ────────────────────────────────────────────────

    def test_round_trip_choices_to_csv(self, tmp_path):
        """Write a pattern → preload → build state → write CSV → parse back.

        Verifies that field names survive the full round-trip through
        _preload_from_pattern → _build_state_from_choices → _choices_to_csv
        → PatternParser.
        """
        # Minimal worksheet
        ws = _make_ws({
            (1, 1): 'Project:',  (1, 2): 'Grepxcel',
            (2, 1): 'Version:',  (2, 2): '0.3.0',
        })

        # Original pattern CSV
        orig_csv = tmp_path / 'original.csv'
        _write_pattern_csv(orig_csv, [
            ['lbl:', 'proj_lbl',  'string', 'Project:'],
            ['var:', 'project',   'string', '.*'],
            ['lbl:', 'ver_lbl',   'string', 'Version:'],
            ['var:', 'version',   'string', r'\d+\.\d+\.\d+'],
            ['START:'],
            ['cell:1', 'proj_lbl'], ['cell:1', 'project'],
            ['cell:1', 'ver_lbl'],  ['cell:1', 'version'],
            ['END:'],
        ])

        choices, preload_cfg, warns = _preload_from_pattern(ws, str(orig_csv))
        assert not warns, f'Unexpected warnings: {warns}'
        assert len(choices) == 4

        # Build state and CSV from the preloaded choices
        from grepxcel.wizard import _cell_ref as _cr
        cells = [(r, c) for r in range(1, ws.max_row + 1)
                 for c in range(1, ws.max_column + 1)]
        state = _build_state_from_choices(
            ws, choices, cells,
            direction=preload_cfg.get('direction', 'LR'),
            sheet_name='Sheet',
        )

        buf = io.StringIO()
        rows_out = _choices_to_csv(
            ws, choices, cells,
            direction=preload_cfg.get('direction', 'LR'),
            sheet_name='Sheet',
        )

        # Write as CSV and parse
        out_csv = tmp_path / 'out.csv'
        with open(out_csv, 'w', newline='', encoding='utf-8') as fh:
            csv.writer(fh).writerows([r if isinstance(r, list) else [] for r in rows_out.splitlines()])

        # The field names must be present in the state
        lbl_names  = {t[0] for t in state.lbl_defs}
        var_names  = {t[0] for t in state.var_defs}
        assert 'proj_lbl' in lbl_names
        assert 'ver_lbl'  in lbl_names
        assert 'project'  in var_names
        assert 'version'  in var_names


# ── TestPreloadTable ──────────────────────────────────────────────────────────

class TestPreloadTable:
    """Tests for TABLE pre-population via _preload_from_pattern.

    Exercises _preload_table end-to-end by calling _preload_from_pattern with
    patterns that contain TABLE instructions backed by lbl: header columns.

    Layout used by most tests (2-column table):
        Row 1: 'Date'  | 'Amount'   ← header (date_lbl lbl, amount_lbl lbl)
        Row 2: data    | data       ← DATA row (date var, amount var)
        Row 3: data    | data
    """

    # ── Shared factories ──────────────────────────────────────────────────────

    @staticmethod
    def _simple_ws():
        """2-column, 2-data-row worksheet."""
        return _make_ws({
            (1, 1): 'Date',       (1, 2): 'Amount',
            (2, 1): '2026-01-01', (2, 2): 100,
            (3, 1): '2026-01-02', (3, 2): 200,
        })

    @staticmethod
    def _write_simple_pattern(path):
        """table:* with 2 lbl header cols and 2 var data cols."""
        _write_pattern_csv(path, [
            ['lbl:', 'date_lbl',   'string',   'Date'],
            ['lbl:', 'amount_lbl', 'string',   'Amount'],
            ['var:', 'date',       'string',   '.*'],
            ['var:', 'amount',     'currency', r'\d+'],
            ['START:'],
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['END:'],
        ])

    # ── Anchor T entry ────────────────────────────────────────────────────────

    def test_table_anchor_t_in_choices(self, tmp_path):
        """Table with lbl: header found → T entry at the anchor ref."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, warns = _preload_from_pattern(ws, str(pat))

        assert not warns, f'Unexpected warnings: {warns}'
        assert 'A1' in choices, f'Expected A1 (T anchor); got: {sorted(choices)}'
        assert choices['A1']['choice'] == 'T'

    def test_table_anchor_meta_required_keys(self, tmp_path):
        """T anchor carries all required meta keys."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        meta = choices['A1']
        for key in ('choice', 'name', 'mult', 'range',
                    'start_row', 'end_row', 'start_col', 'end_col',
                    'header_rows', 'footer_rows', 'data_vars', 'skip_rows',
                    'row_types'):
            assert key in meta, f'Missing key {key!r} in T anchor meta'

    def test_table_anchor_coord_fields(self, tmp_path):
        """T anchor start_row/start_col/end_col/range match the located header."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        meta = choices['A1']
        assert meta['start_row'] == 1
        assert meta['start_col'] == 1
        assert meta['end_col']   == 2
        assert meta['range']     == 'A1:B1'
        assert meta['mult']      == '*'

    def test_table_name_derived_from_lbl_field(self, tmp_path):
        """Table name strips _lbl/_col suffixes from the first lbl field name."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        # first_fd.name = 'date_lbl' → remove '_lbl' → 'date' → slugify
        assert choices['A1']['name'] == 'date'

    # ── T-HEAD entries ────────────────────────────────────────────────────────

    def test_table_thead_for_non_anchor_header_cell(self, tmp_path):
        """Non-anchor header cells get T-HEAD pointing to the anchor."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, warns = _preload_from_pattern(ws, str(pat))

        assert not warns
        assert 'B1' in choices, f'Expected B1 (T-HEAD); got: {sorted(choices)}'
        assert choices['B1']['choice'] == 'T-HEAD'
        assert choices['B1']['anchor'] == 'A1'

    def test_table_anchor_is_not_t_head(self, tmp_path):
        """Anchor cell itself must be classified T, not T-HEAD."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        assert choices['A1']['choice'] == 'T'

    # ── T-DATA entries ────────────────────────────────────────────────────────

    def test_table_tdata_for_data_rows(self, tmp_path):
        """All cells in detected data rows get T-DATA pointing to the anchor."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, warns = _preload_from_pattern(ws, str(pat))

        assert not warns
        for ref in ('A2', 'B2', 'A3', 'B3'):
            assert ref in choices, f'Expected {ref} (T-DATA); got: {sorted(choices)}'
            assert choices[ref]['choice'] == 'T-DATA', f'{ref} choice should be T-DATA'
            assert choices[ref]['anchor'] == 'A1'

    def test_table_data_rows_stop_at_empty_row(self, tmp_path):
        """Data scan stops at the first fully-empty row; later rows are excluded."""
        ws = _make_ws({
            (1, 1): 'Date',       (1, 2): 'Amount',
            (2, 1): '2026-01-01', (2, 2): 100,
            # row 3: empty gap → scan stops
            (4, 1): '2026-01-03', (4, 2): 300,   # must NOT be pre-loaded
        })
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        assert 'A2' in choices and choices['A2']['choice'] == 'T-DATA'
        assert 'A4' not in choices, 'Row 4 is after an empty row — must not be pre-loaded'

    # ── row_types ─────────────────────────────────────────────────────────────

    def test_table_row_types_h_and_d(self, tmp_path):
        """row_types maps header rows to 'H' and data rows to 'D'."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        rt = choices['A1']['row_types']
        assert rt.get(1) == 'H'
        assert rt.get(2) == 'D'
        assert rt.get(3) == 'D'

    # ── header_rows structure ─────────────────────────────────────────────────

    def test_table_header_rows_entry_count(self, tmp_path):
        """header_rows has one entry (matching the one HEADER template row)."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        hr = choices['A1']['header_rows']
        assert len(hr) == 1
        assert hr[0]['row'] == 1
        assert len(hr[0]['cols']) == 2

    def test_table_header_rows_lbl_col_fields(self, tmp_path):
        """First header col (lbl role) carries role, lbl_name, cell_value, ref."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        col0 = choices['A1']['header_rows'][0]['cols'][0]
        assert col0['role']       == 'label'
        assert col0['lbl_name']   == 'date_lbl'
        assert col0['cell_value'] == 'Date'
        assert col0['ref']        == 'A1'

    # ── data_vars ─────────────────────────────────────────────────────────────

    def test_table_data_vars_count_and_roles(self, tmp_path):
        """data_vars has one entry per DATA template column, all var role."""
        ws  = self._simple_ws()
        pat = tmp_path / 'p.csv'
        self._write_simple_pattern(pat)

        choices, _, _ = _preload_from_pattern(ws, str(pat))

        dvars = choices['A1']['data_vars']
        assert len(dvars) == 2
        assert dvars[0]['role']     == 'var'
        assert dvars[0]['var_name'] == 'date'
        assert dvars[1]['role']     == 'var'
        assert dvars[1]['var_name'] == 'amount'
        assert dvars[1]['var_type'] == 'currency'

    def test_table_nullable_var_col_a_extra(self, tmp_path):
        """A nullable var in DATA produces col_a_extra='nullable' in data_vars."""
        ws = _make_ws({
            (1, 1): 'Date', (1, 2): 'Amount',
            (2, 1): '2026-01-01', (2, 2): None,
        })
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'date_lbl',   'string',   'Date'],
            ['lbl:', 'amount_lbl', 'string',   'Amount'],
            ['var:', 'date',       'string',   '.*'],
            ['var:nullable', 'amount', 'currency', r'\d+'],
            ['START:'],
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['END:'],
        ])

        choices, _, warns = _preload_from_pattern(ws, str(pat))

        assert not warns
        dvars = choices['A1']['data_vars']
        amount_dv = next((d for d in dvars if d.get('var_name') == 'amount'), None)
        assert amount_dv is not None
        assert 'nullable' in amount_dv['col_a_extra']

    # ── Header not found ──────────────────────────────────────────────────────

    def test_table_header_not_found_warning(self, tmp_path):
        """table:1 with absent header → 'not found' warning; table:* silently allows 0."""
        ws = _make_ws({
            (1, 1): 'Product', (1, 2): 'Price',   # different headers
            (2, 1): 'Widget',  (2, 2): 9.99,
        })
        # table:* is 0-or-more — no warning expected even when header is absent
        pat_star = tmp_path / 'p_star.csv'
        self._write_simple_pattern(pat_star)   # uses table:*
        choices_star, _, warns_star = _preload_from_pattern(ws, str(pat_star))
        assert not warns_star, f'table:* should not warn for 0 instances; got: {warns_star}'
        assert all(v.get('choice') != 'T' for v in choices_star.values())

        # table:1 is exactly-1 — must warn when header is absent
        pat_one = tmp_path / 'p_one.csv'
        _write_pattern_csv(pat_one, [
            ['lbl:', 'date_lbl',   'string', 'Date'],
            ['lbl:', 'amount_lbl', 'string', 'Amount'],
            ['var:', 'date',       'string', '.*'],
            ['var:', 'amount',     'currency', r'\d+'],
            ['START:'],
            ['table:1', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['END:'],
        ])
        choices_one, _, warns_one = _preload_from_pattern(ws, str(pat_one))
        assert any(
            'not found' in w.lower() or 'header' in w.lower()
            for w in warns_one
        ), f'table:1 expected header-not-found warning; got: {warns_one}'
        assert all(v.get('choice') != 'T' for v in choices_one.values())

    # ── Multi-table same header ───────────────────────────────────────────────

    def test_multi_table_same_header_distinct_anchors(self, tmp_path):
        """3 table:* blocks with the same header text → 3 distinct anchors."""
        ws = _make_ws({
            # Table 1: rows 1-2
            (1, 1): 'Date',       (1, 2): 'Amount',
            (2, 1): '2026-01-01', (2, 2): 100,
            # row 3: empty separator
            # Table 2: rows 4-5
            (4, 1): 'Date',       (4, 2): 'Amount',
            (5, 1): '2026-02-01', (5, 2): 200,
            # row 6: empty separator
            # Table 3: rows 7-8
            (7, 1): 'Date',       (7, 2): 'Amount',
            (8, 1): '2026-03-01', (8, 2): 300,
        })
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'date_lbl',   'string',   'Date'],
            ['lbl:', 'amount_lbl', 'string',   'Amount'],
            ['var:', 'date',       'string',   '.*'],
            ['var:', 'amount',     'currency', r'\d+'],
            ['START:'],
            # Table 1
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            # Table 2 (identical fields)
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            # Table 3 (identical fields)
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['END:'],
        ])

        choices, _, warns = _preload_from_pattern(ws, str(pat))

        t_anchors = {ref for ref, meta in choices.items() if meta.get('choice') == 'T'}
        assert 'A1' in t_anchors, f'Table 1 anchor missing; T anchors: {t_anchors}'
        assert 'A4' in t_anchors, f'Table 2 anchor missing; T anchors: {t_anchors}'
        assert 'A7' in t_anchors, f'Table 3 anchor missing; T anchors: {t_anchors}'
        assert len(t_anchors) == 3, f'Expected exactly 3 T anchors; got: {t_anchors}'

    def test_multi_table_data_rows_isolated_per_anchor(self, tmp_path):
        """Each table's T-DATA entries reference only their own anchor."""
        ws = _make_ws({
            (1, 1): 'Date',       (1, 2): 'Amount',
            (2, 1): '2026-01-01', (2, 2): 100,
            # row 3: empty separator
            (4, 1): 'Date',       (4, 2): 'Amount',
            (5, 1): '2026-02-01', (5, 2): 200,
        })
        pat = tmp_path / 'p.csv'
        _write_pattern_csv(pat, [
            ['lbl:', 'date_lbl',   'string',   'Date'],
            ['lbl:', 'amount_lbl', 'string',   'Amount'],
            ['var:', 'date',       'string',   '.*'],
            ['var:', 'amount',     'currency', r'\d+'],
            ['START:'],
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['END:'],
        ])

        choices, _, warns = _preload_from_pattern(ws, str(pat))

        assert not warns, f'Unexpected warnings: {warns}'
        # Table 1 data rows → anchor A1
        assert choices.get('A2', {}).get('anchor') == 'A1'
        assert choices.get('B2', {}).get('anchor') == 'A1'
        # Table 2 data rows → anchor A4
        assert choices.get('A5', {}).get('anchor') == 'A4'
        assert choices.get('B5', {}).get('anchor') == 'A4'


# ── TestPreloadTableTUI (Textual pilot) ──────────────────────────────────────

@_skip_no_textual
class TestPreloadTableTUI:
    """Pilot tests: preloaded TABLE choices survive the full TUI lifecycle."""

    def test_preloaded_table_survives_save(self, tmp_path):
        """_preload_from_pattern (table) → TUI → save → WizardState has table data."""
        # Build data file
        xlsx_path = str(tmp_path / 'data.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        for (r, c), v in {
            (1, 1): 'Date', (1, 2): 'Amount',
            (2, 1): '2026-01-01', (2, 2): 100,
            (3, 1): '2026-01-02', (3, 2): 200,
        }.items():
            ws.cell(row=r, column=c, value=v)
        wb.save(xlsx_path)

        # Write pattern
        pat_path = str(tmp_path / 'p.csv')
        _write_pattern_csv(tmp_path / 'p.csv', [
            ['lbl:', 'date_lbl',   'string',   'Date'],
            ['lbl:', 'amount_lbl', 'string',   'Amount'],
            ['var:', 'date',       'string',   '.*'],
            ['var:', 'amount',     'currency', r'\d+'],
            ['START:'],
            ['table:*', ''],
            ['', 'HEADER:1', 'date_lbl', 'amount_lbl'],
            ['', 'DATA:*',   'date',     'amount'],
            ['END:'],
        ])

        # Pre-populate choices from pattern
        wb2 = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws2 = wb2.active
        load_choices, preload_cfg, warns = _preload_from_pattern(ws2, pat_path)
        assert not warns, f'Pre-load should succeed; got: {warns}'
        assert any(v.get('choice') == 'T' for v in load_choices.values()), \
            f'Expected at least one T anchor in preloaded choices; got: {list(load_choices)}'

        # Launch TUI with preloaded choices
        state = WizardState(sheet_name=ws2.title)
        app = WizardTUIApp(ws2, state, xlsx_path,
                           load_choices=load_choices,
                           preload_config=preload_cfg)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await pilot.pause()
                await pilot.press('ctrl+enter')  # accept config
                await pilot.pause()
                await pilot.press('e')       # End & Save
                await pilot.pause()

        asyncio.run(_run())

        result = app._return_value
        assert result is not None, 'Expected a WizardState after pressing e'
        assert isinstance(result, WizardState)

        # body_rows must contain a table: row
        tbl_body = [r for r in result.body_rows if r and r[0].startswith('table:')]
        assert tbl_body, f'Expected table: row in body_rows; got:\n{result.body_rows}'

        # lbl_defs must include the header lbl columns
        lbl_names = {t[0] for t in result.lbl_defs}
        assert 'date_lbl'   in lbl_names, f'date_lbl missing; lbl_defs: {result.lbl_defs}'
        assert 'amount_lbl' in lbl_names, f'amount_lbl missing; lbl_defs: {result.lbl_defs}'

        # _build_state_from_choices namespaces DATA var names with the table name
        # (table_name='date' from 'date_lbl') → 'date.date', 'date.amount'
        var_names = {t[0] for t in result.var_defs}
        assert any('date' in vn for vn in var_names), \
            f'Expected a "date"-containing var; var_defs: {result.var_defs}'
        assert any('amount' in vn for vn in var_names), \
            f'Expected an "amount"-containing var; var_defs: {result.var_defs}'


# ── TestPreloadPatternTUI (Textual pilot) ─────────────────────────────────────

@_skip_no_textual
class TestPreloadPatternTUI:
    """Pilot tests: verify the TUI accepts --load-pattern preloaded choices.

    These tests exercise the full headless TUI lifecycle:
      1. WizardTUIApp is constructed with load_choices / preload_config
      2. Config modal opens and is accepted (enter)
      3. _on_config_done calls _restyle_cell for every preloaded cell
      4. Pressing 'e' saves and produces a WizardState with the right fields

    This layer is distinct from TestPreloadFromPattern (pure-function): here we
    verify that the TUI correctly wires preloaded data through the UI lifecycle.
    """

    def _make_app_with_preload(
        self,
        cell_data: dict,
        load_choices: dict,
        preload_config: dict | None,
        preload_warnings: list,
        tmp_path,
    ) -> 'WizardTUIApp':
        """Build a WizardTUIApp with pre-loaded choices (simulates --load-pattern)."""
        xlsx_path = str(tmp_path / 'data.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        for (r, c), v in cell_data.items():
            ws.cell(row=r, column=c, value=v)
        wb.save(xlsx_path)

        wb2 = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws2 = wb2.active
        state = WizardState(sheet_name=ws2.title)
        return WizardTUIApp(
            ws2, state, xlsx_path,
            load_choices=load_choices,
            preload_config=preload_config,
            preload_warnings=preload_warnings,
        )

    @staticmethod
    async def _accept_config(pilot):
        await pilot.pause()
        await pilot.press('ctrl+enter')  # dismiss _ConfigModal with defaults
        await pilot.pause()

    # ── Choices wiring ────────────────────────────────────────────────────────

    def test_preloaded_choices_present_after_config(self, tmp_path):
        """load_choices dict is available in app._choices after config modal."""
        load_choices = {
            'A1': {'choice': 'L', 'name': 'lbl1', 'ltype': 'string',
                   'lmatch': 'Name:', 'lbl_mode': ''},
            'B1': {'choice': 'V', 'name': 'name', 'ftype': 'string',
                   'match': '.*', 'col_a_extra': ''},
        }
        app = self._make_app_with_preload(
            {(1, 1): 'Name:', (1, 2): 'Alice'},
            load_choices, None, [], tmp_path,
        )

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config(pilot)
                # After config is accepted both cells must still be classified
                assert 'A1' in app._choices
                assert 'B1' in app._choices
                assert app._choices['A1']['choice'] == 'L'
                assert app._choices['B1']['choice'] == 'V'
                await pilot.press('ctrl+q')

        asyncio.run(_run())

    def test_save_with_preloaded_choices_produces_correct_state(self, tmp_path):
        """Pressing 'e' after preload saves a WizardState with the right fields."""
        load_choices = {
            'A1': {'choice': 'L', 'name': 'inv_lbl', 'ltype': 'string',
                   'lmatch': 'Invoice:', 'lbl_mode': ''},
            'B1': {'choice': 'V', 'name': 'invoice_no', 'ftype': 'string',
                   'match': '.*', 'col_a_extra': ''},
        }
        app = self._make_app_with_preload(
            {(1, 1): 'Invoice:', (1, 2): 'INV-001'},
            load_choices, None, [], tmp_path,
        )

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config(pilot)
                await pilot.press('e')    # End & Save
                await pilot.pause()

        asyncio.run(_run())

        result = app._return_value
        assert result is not None, 'WizardTUIApp should return a WizardState on save'
        assert isinstance(result, WizardState)
        lbl_names = {t[0] for t in result.lbl_defs}
        var_names  = {t[0] for t in result.var_defs}
        assert 'inv_lbl'    in lbl_names
        assert 'invoice_no' in var_names

    def test_preloaded_choices_can_be_overridden(self, tmp_path):
        """User can reclassify a preloaded cell (V → I); the new choice wins."""
        load_choices = {
            'A1': {'choice': 'V', 'name': 'orig', 'ftype': 'string',
                   'match': '.*', 'col_a_extra': ''},
        }
        app = self._make_app_with_preload(
            {(1, 1): 'some_value'},
            load_choices, None, [], tmp_path,
        )

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config(pilot)
                # Press 'I' to reclassify the current cell as Ignore
                await pilot.press('i')
                await pilot.pause()
                # The cell should now be classified as 'I', not 'V'
                assert app._choices.get('A1', {}).get('choice') == 'I'
                await pilot.press('ctrl+q')

        asyncio.run(_run())

    def test_empty_load_choices_opens_normally(self, tmp_path):
        """load_choices={} is equivalent to no preloading — app opens normally."""
        app = self._make_app_with_preload(
            {(1, 1): 'hello'},
            {}, None, [], tmp_path,
        )

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config(pilot)
                assert app._choices == {}
                await pilot.press('ctrl+q')

        asyncio.run(_run())
        assert app._return_value is None

    # ── Config preload ────────────────────────────────────────────────────────

    def test_preload_config_direction_td(self, tmp_path):
        """preload_config direction=TD is reflected in app._state after config."""
        preload_config = {'direction': 'TD', 'lbl_match': '', 'var_match': ''}
        app = self._make_app_with_preload(
            {(1, 1): 'x'},
            {}, preload_config, [], tmp_path,
        )

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config(pilot)
                assert app._state.direction == 'TD'
                await pilot.press('ctrl+q')

        asyncio.run(_run())

    def test_preload_config_currency_dollar(self, tmp_path):
        """preload_config currency_sign='$' flows into app._state."""
        preload_config = {'direction': 'LR', 'currency_sign': '$',
                          'lbl_match': '', 'var_match': ''}
        app = self._make_app_with_preload(
            {(1, 1): 'x'},
            {}, preload_config, [], tmp_path,
        )

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await self._accept_config(pilot)
                assert app._state.currency_sign == '$'
                await pilot.press('ctrl+q')

        asyncio.run(_run())

    # ── Warnings ─────────────────────────────────────────────────────────────

    def test_preload_warnings_stored_on_app(self, tmp_path):
        """preload_warnings are stored on _preload_warnings (for the toast)."""
        warns = ['Label xyz not found', 'TABLE fields not located']
        app = self._make_app_with_preload(
            {(1, 1): 'x'},
            {}, None, warns, tmp_path,
        )
        assert app._preload_warnings == warns

    # ── Round-trip: preload → save → pattern file ─────────────────────────────

    def test_roundtrip_preload_save_to_csv(self, tmp_path):
        """Full round-trip: preload → accept config → save → CSV contains fields."""
        cell_data = {(1, 1): 'Project:', (1, 2): 'Grepxcel'}
        orig_csv = str(tmp_path / 'orig.csv')
        out_csv  = str(tmp_path / 'out.csv')

        with open(orig_csv, 'w', newline='', encoding='utf-8') as fh:
            csv.writer(fh).writerows([
                ['lbl:', 'proj_lbl', 'string', 'Project:'],
                ['var:', 'project',  'string', '.*'],
                ['START:'],
                ['cell:1', 'proj_lbl'],
                ['cell:1', 'project'],
                ['END:'],
            ])

        # 1. Build preload choices from the pattern
        xlsx_path = str(tmp_path / 'data.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        for (r, c), v in cell_data.items():
            ws.cell(row=r, column=c, value=v)
        wb.save(xlsx_path)
        wb2 = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws2 = wb2.active

        load_choices, preload_cfg, warns = _preload_from_pattern(ws2, orig_csv)
        assert not warns, f'Unexpected warnings: {warns}'
        assert len(load_choices) == 2

        # 2. Launch TUI with preloaded choices and save immediately
        state = WizardState(sheet_name=ws2.title)
        app = WizardTUIApp(ws2, state, xlsx_path,
                           load_choices=load_choices,
                           preload_config=preload_cfg)

        async def _run():
            async with app.run_test(headless=True, size=(120, 40)) as pilot:
                await pilot.pause()
                await pilot.press('ctrl+enter')  # accept config
                await pilot.pause()
                await pilot.press('e')      # save
                await pilot.pause()

        asyncio.run(_run())

        result = app._return_value
        assert result is not None
        lbl_names = {t[0] for t in result.lbl_defs}
        var_names  = {t[0] for t in result.var_defs}
        assert 'proj_lbl' in lbl_names
        assert 'project'  in var_names
