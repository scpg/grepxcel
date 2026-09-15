"""
Unit tests for grepxcel.watcher — watch mode.

All tests run without the watchdog package installed: they exercise the
non-watchdog code paths (helpers, validation, error handling) by patching
the observer and using the handler directly.
"""

import json
import os
import sys
import threading
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest


# ── _is_temp_file ─────────────────────────────────────────────────────────────

from grepxcel.watcher import _is_temp_file


class TestIsTempFile:
    def test_excel_lock_file(self):
        assert _is_temp_file('~$invoice.xlsx') is True

    def test_excel_lock_full_path(self):
        assert _is_temp_file('/tmp/inbox/~$report.xlsx') is True

    def test_dotfile(self):
        assert _is_temp_file('.hidden.xlsx') is True

    def test_normal_file(self):
        assert _is_temp_file('invoice.xlsx') is False

    def test_normal_file_full_path(self):
        assert _is_temp_file('/home/user/inbox/invoice.xlsx') is False

    def test_tilde_in_middle(self):
        # tilde not at the start is not a temp file
        assert _is_temp_file('report~backup.xlsx') is False


# ── watch() argument validation ───────────────────────────────────────────────

class TestWatchValidation:
    """Tests that watch() raises the right errors before starting the observer."""

    def test_missing_pattern_file(self, tmp_path):
        from grepxcel.watcher import watch
        with pytest.raises(FileNotFoundError, match='Pattern file not found'):
            watch(
                pattern_path=str(tmp_path / 'nonexistent.xlsx'),
                directory=str(tmp_path),
            )

    def test_missing_directory(self, tmp_path):
        from grepxcel.watcher import watch
        # Create a dummy pattern file so the first check passes
        p = tmp_path / 'pattern.xlsx'
        p.write_bytes(b'dummy')
        with pytest.raises(FileNotFoundError, match='Watch directory not found'):
            watch(
                pattern_path=str(p),
                directory=str(tmp_path / 'nonexistent'),
            )

    def test_invalid_on_error(self, tmp_path):
        from grepxcel.watcher import watch
        p = tmp_path / 'pattern.xlsx'
        p.write_bytes(b'dummy')
        d = tmp_path / 'inbox'
        d.mkdir()
        with pytest.raises(ValueError, match="on_error must be"):
            watch(
                pattern_path=str(p),
                directory=str(d),
                on_error='explode',
            )

    def test_no_watchdog_raises_import_error(self, tmp_path):
        """If watchdog is not installed watch() raises ImportError with install hint."""
        from grepxcel.watcher import watch
        p = tmp_path / 'pattern.xlsx'
        p.write_bytes(b'dummy')
        d = tmp_path / 'inbox'
        d.mkdir()
        with patch.dict(sys.modules, {'watchdog': None,
                                       'watchdog.observers': None,
                                       'watchdog.events': None}):
            with pytest.raises(ImportError, match="grepxcel\\[watch\\]"):
                watch(pattern_path=str(p), directory=str(d))


# ── _XlsxHandler ─────────────────────────────────────────────────────────────

class TestXlsxHandler:
    """Tests for the extraction handler, mocking Engine and watchdog events."""

    def _make_handler(self, output_dir=None, on_error='continue',
                      quiet=True, callback=None):
        from grepxcel.watcher import _XlsxHandler
        return _XlsxHandler(
            pattern_path='/fake/pattern.xlsx',
            output_dir=str(output_dir) if output_dir else None,
            fmt='nested',
            sheet=None,
            on_error=on_error,
            quiet=quiet,
            engine_kwargs={},
            callback=callback,
        )

    def _make_event(self, path, is_dir=False, event_type='created'):
        """Create a mock watchdog event."""
        evt = MagicMock()
        evt.is_directory = is_dir
        if event_type == 'moved':
            evt.dest_path = path
            evt.src_path = '/old/path.xlsx'
            # make isinstance checks work
            from watchdog.events import FileMovedEvent
            evt.__class__ = FileMovedEvent
        else:
            evt.src_path = path
            from watchdog.events import FileCreatedEvent
            evt.__class__ = FileCreatedEvent
        return evt

    @pytest.fixture(autouse=True)
    def _require_watchdog(self):
        pytest.importorskip('watchdog')

    def test_skips_temp_files(self):
        handler = self._make_handler()
        evt = self._make_event('/inbox/~$invoice.xlsx')
        with patch.object(handler, '_process') as mock_proc:
            handler.dispatch(evt)
            mock_proc.assert_not_called()

    def test_skips_directories(self):
        handler = self._make_handler()
        evt = self._make_event('/inbox/subdir', is_dir=True)
        with patch.object(handler, '_process') as mock_proc:
            handler.dispatch(evt)
            mock_proc.assert_not_called()

    def test_skips_non_xlsx(self):
        handler = self._make_handler()
        evt = self._make_event('/inbox/file.csv')
        with patch.object(handler, '_process') as mock_proc:
            handler.dispatch(evt)
            mock_proc.assert_not_called()

    def test_processes_xlsx_on_create(self):
        handler = self._make_handler()
        evt = self._make_event('/inbox/invoice.xlsx')
        with patch.object(handler, '_process') as mock_proc:
            handler.dispatch(evt)
            mock_proc.assert_called_once_with('/inbox/invoice.xlsx')

    def test_processes_xlsx_on_move(self):
        handler = self._make_handler()
        evt = self._make_event('/inbox/invoice.xlsx', event_type='moved')
        with patch.object(handler, '_process') as mock_proc:
            handler.dispatch(evt)
            mock_proc.assert_called_once_with('/inbox/invoice.xlsx')

    def test_calls_callback_on_success(self, tmp_path):
        results = []
        def cb(path, result, error):
            results.append((path, result, error))

        handler = self._make_handler(callback=cb)
        fake_result = {'inv': {'number': 'INV-001'}}

        with patch('grepxcel.watcher._XlsxHandler._Engine') as MockEngine:
            instance = MockEngine.return_value
            instance.process.return_value = fake_result
            handler._process('/inbox/invoice.xlsx')

        assert len(results) == 1
        assert results[0][0] == '/inbox/invoice.xlsx'
        assert results[0][1] == fake_result
        assert results[0][2] is None

    def test_calls_callback_on_error(self, tmp_path):
        results = []
        def cb(path, result, error):
            results.append((path, result, error))

        handler = self._make_handler(callback=cb, on_error='continue')
        boom = RuntimeError('bad file')

        with patch('grepxcel.watcher._XlsxHandler._Engine') as MockEngine:
            instance = MockEngine.return_value
            instance.process.side_effect = boom
            handler._process('/inbox/bad.xlsx')

        assert len(results) == 1
        assert results[0][1] is None
        assert results[0][2] is boom

    def test_on_error_stop_sets_flag(self, tmp_path):
        handler = self._make_handler(on_error='stop')

        with patch('grepxcel.watcher._XlsxHandler._Engine') as MockEngine:
            instance = MockEngine.return_value
            instance.process.side_effect = RuntimeError('boom')
            handler._process('/inbox/bad.xlsx')

        assert handler.should_stop() is True

    def test_on_error_continue_does_not_set_flag(self, tmp_path):
        handler = self._make_handler(on_error='continue')

        with patch('grepxcel.watcher._XlsxHandler._Engine') as MockEngine:
            instance = MockEngine.return_value
            instance.process.side_effect = RuntimeError('boom')
            handler._process('/inbox/bad.xlsx')

        assert handler.should_stop() is False

    def test_writes_json_to_output_dir(self, tmp_path):
        out = tmp_path / 'output'
        out.mkdir()
        handler = self._make_handler(output_dir=out)
        fake_result = {'inv': {'number': 'INV-001'}}

        with patch('grepxcel.watcher._XlsxHandler._Engine') as MockEngine:
            instance = MockEngine.return_value
            instance.process.return_value = fake_result
            handler._process('/inbox/invoice.xlsx')

        out_file = out / 'invoice.json'
        assert out_file.exists()
        data = json.loads(out_file.read_text())
        assert data == fake_result

    def test_writes_json_to_stdout(self, tmp_path, capsys):
        handler = self._make_handler(output_dir=None)
        fake_result = {'inv': {'number': 'INV-001'}}

        with patch('grepxcel.watcher._XlsxHandler._Engine') as MockEngine:
            instance = MockEngine.return_value
            instance.process.return_value = fake_result
            handler._process('/inbox/invoice.xlsx')

        captured = capsys.readouterr()
        data = json.loads(captured.out)
        assert data == fake_result


# ── CLI integration ───────────────────────────────────────────────────────────

class TestWatchCLI:
    """Test the CLI wires up correctly without actually starting the observer."""

    def test_watch_help(self):
        from grepxcel.cli import main
        with pytest.raises(SystemExit) as exc:
            main(['watch', '--help'])
        assert exc.value.code == 0

    def test_watch_missing_pattern_exits_1(self, tmp_path):
        from grepxcel.cli import main
        d = tmp_path / 'inbox'
        d.mkdir()
        with pytest.raises(SystemExit) as exc:
            main(['watch', '-p', str(tmp_path / 'nonexistent.xlsx'),
                  str(d)])
        assert exc.value.code == 1

    def test_watch_missing_dir_exits_1(self, tmp_path):
        from grepxcel.cli import main
        p = tmp_path / 'pattern.xlsx'
        p.write_bytes(b'dummy')
        with pytest.raises(SystemExit) as exc:
            main(['watch', '-p', str(p), str(tmp_path / 'nodir')])
        assert exc.value.code == 1

    def test_watch_no_watchdog_exits_1(self, tmp_path):
        from grepxcel.cli import main
        p = tmp_path / 'pattern.xlsx'
        p.write_bytes(b'dummy')
        d = tmp_path / 'inbox'
        d.mkdir()
        with patch.dict(sys.modules, {'watchdog': None,
                                       'watchdog.observers': None,
                                       'watchdog.events': None}):
            with pytest.raises(SystemExit) as exc:
                main(['watch', '-p', str(p), str(d)])
            assert exc.value.code == 1
