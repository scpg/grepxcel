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
        _TEXTUAL_OK,
    )
    if _TEXTUAL_OK:
        from grepxcel.wizard_tui import WizardTUIApp
    from grepxcel.wizard import WizardState
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
        assert any(n == 'invoice_label' for n, _, _ in state.lbl_defs)

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
        assert any(n == 'invoice_number' for n, _, _ in state.var_defs)

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
