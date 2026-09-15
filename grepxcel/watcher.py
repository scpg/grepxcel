"""
Watch mode — monitor a directory for new Excel files and extract them automatically.

Usage:
    grepxcel watch -p pattern.xlsx inbox/
    grepxcel watch -p pattern.xlsx inbox/ -o output/
    grepxcel watch -p pattern.xlsx inbox/ --recursive --format csv -o output/

The watcher starts a persistent process that monitors *directory* for newly created
or moved-in ``.xlsx`` files.  Each file is processed with the given pattern as soon
as it appears.  Results go to stdout (one JSON blob per file, newline-delimited) or,
if ``-o`` is given, to individual ``.json`` files in *output_dir*.

Temp files (``~$*.xlsx``, created by Excel while a workbook is open) are silently
skipped.  Errors during a single extraction are logged to stderr and the watcher
continues; use ``--on-error stop`` to exit on the first error instead.

Requires:
    pip install 'grepxcel[watch]'   # adds the watchdog dependency
"""

from __future__ import annotations

import json
import os
import sys
import time
import datetime
import threading
from pathlib import Path
from typing import Callable


# ── Internal helpers ──────────────────────────────────────────────────────────

def _is_temp_file(path: str) -> bool:
    """True for Excel lock/temp files that should never be processed."""
    name = os.path.basename(path)
    return name.startswith('~$') or name.startswith('.')


def _json_default(obj):
    """JSON serialiser for datetime objects produced by openpyxl."""
    if isinstance(obj, (datetime.date, datetime.datetime, datetime.time)):
        return obj.isoformat()
    if isinstance(obj, datetime.timedelta):
        return str(obj)
    raise TypeError(f'Type {type(obj).__name__} is not JSON serialisable')


# ── Event handler ─────────────────────────────────────────────────────────────

class _XlsxHandler:
    """Watchdog-compatible event handler that extracts new .xlsx files.

    Does not inherit from watchdog directly so the module can be imported
    without watchdog installed (the import happens lazily in ``watch()``).
    """

    def __init__(
        self,
        pattern_path: str,
        output_dir: str | None,
        fmt: str,
        sheet: str | None,
        on_error: str,
        quiet: bool,
        engine_kwargs: dict,
        callback: Callable[[str, dict | None, Exception | None], None] | None,
    ):
        self.pattern_path = pattern_path
        self.output_dir = output_dir
        self.fmt = fmt
        self.sheet = sheet
        self.on_error = on_error          # 'continue' | 'stop'
        self.quiet = quiet
        self.engine_kwargs = engine_kwargs
        self.callback = callback
        self._stop_event = threading.Event()

        # Lazy import to keep the module importable without watchdog
        from .engine import Engine
        from .logger import Logger, VerbosityLevel
        self._Engine = Engine
        self._Logger = Logger
        self._VerbosityLevel = VerbosityLevel

    def should_stop(self) -> bool:
        return self._stop_event.is_set()

    def stop(self) -> None:
        self._stop_event.set()

    def dispatch(self, event) -> None:
        """Called by watchdog for every file-system event."""
        # We only care about file creation and moves-in
        from watchdog.events import FileCreatedEvent, FileMovedEvent
        if isinstance(event, FileMovedEvent):
            path = event.dest_path
        elif isinstance(event, FileCreatedEvent):
            path = event.src_path
        else:
            return

        if event.is_directory:
            return
        if not path.lower().endswith('.xlsx'):
            return
        if _is_temp_file(path):
            return

        self._process(path)

    def _process(self, path: str) -> None:
        """Extract *path* and write result or call callback."""
        if not self.quiet:
            ts = datetime.datetime.now().strftime('%H:%M:%S')
            print(f'[{ts}] {path}', file=sys.stderr)

        logger = self._Logger(level=self._VerbosityLevel.QUIET)
        engine = self._Engine(**self.engine_kwargs)
        try:
            result = engine.process(
                self.pattern_path,
                path,
                logger=logger,
                sheet=self.sheet,
                output_format=self.fmt if self.fmt not in ('csv', 'xlsx') else 'nested',
            )
        except Exception as exc:
            if not self.quiet:
                print(f'  ERROR: {exc}', file=sys.stderr)
            if self.callback:
                self.callback(path, None, exc)
            if self.on_error == 'stop':
                self.stop()
            return

        if self.callback:
            self.callback(path, result, None)
            return

        if self.output_dir:
            out_name = Path(path).stem + '.json'
            out_path = Path(self.output_dir) / out_name
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2, default=_json_default)
            if not self.quiet:
                print(f'  → {out_path}', file=sys.stderr)
        else:
            # stdout: one JSON object per file, newline-delimited
            print(json.dumps(result, default=_json_default))


# ── Public API ────────────────────────────────────────────────────────────────

def watch(
    pattern_path: str,
    directory: str,
    *,
    output_dir: str | None = None,
    recursive: bool = False,
    fmt: str = 'nested',
    sheet: str | None = None,
    on_error: str = 'continue',
    quiet: bool = False,
    max_size_mb: float = 5.0,
    max_uncompressed_mb: float = 50.0,
    callback: Callable[[str, dict | None, Exception | None], None] | None = None,
) -> None:
    """Monitor *directory* and extract new .xlsx files as they arrive.

    Blocks until interrupted (KeyboardInterrupt / SIGINT).

    Args:
        pattern_path:        Path to the pattern file (.xlsx or .csv).
        directory:           Directory to watch.
        output_dir:          Write extracted JSON here instead of stdout.
        recursive:           Also watch subdirectories.
        fmt:                 Output format passed to the engine ('nested', 'legacy').
        sheet:               Sheet name or index (default: active sheet).
        on_error:            'continue' (default) or 'stop' — behaviour on extraction error.
        quiet:               Suppress progress messages on stderr.
        max_size_mb:         Compressed file size limit passed to the engine.
        max_uncompressed_mb: Uncompressed content limit passed to the engine.
        callback:            Optional function(path, result, error) called per file
                             instead of writing to stdout/disk.  When supplied,
                             ``output_dir`` is ignored.

    Raises:
        ImportError:   if watchdog is not installed.
        FileNotFoundError: if *pattern_path* or *directory* does not exist.
        ValueError:    if *on_error* is not 'continue' or 'stop'.
    """
    # Validate arguments before importing watchdog so callers get the right
    # exception regardless of whether watchdog is installed.
    if on_error not in ('continue', 'stop'):
        raise ValueError(f"on_error must be 'continue' or 'stop', got {on_error!r}")

    pattern_path = os.path.abspath(pattern_path)
    directory = os.path.abspath(directory)

    if not os.path.isfile(pattern_path):
        raise FileNotFoundError(f'Pattern file not found: {pattern_path}')
    if not os.path.isdir(directory):
        raise FileNotFoundError(f'Watch directory not found: {directory}')

    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError as exc:
        raise ImportError(
            "watch mode requires the 'watchdog' package.\n"
            "Install it with:  pip install 'grepxcel[watch]'"
        ) from exc

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    engine_kwargs: dict = {}
    if max_size_mb:
        engine_kwargs['max_file_mb'] = max_size_mb
    if max_uncompressed_mb:
        engine_kwargs['max_uncompressed_mb'] = max_uncompressed_mb

    handler = _XlsxHandler(
        pattern_path=pattern_path,
        output_dir=output_dir,
        fmt=fmt,
        sheet=sheet,
        on_error=on_error,
        quiet=quiet,
        engine_kwargs=engine_kwargs,
        callback=callback,
    )

    # Wrap in a watchdog-compatible handler
    class _WatchdogBridge(FileSystemEventHandler):
        def dispatch(self, event):
            handler.dispatch(event)

    observer = Observer()
    observer.schedule(_WatchdogBridge(), directory, recursive=recursive)
    observer.start()

    if not quiet:
        print(
            f'Watching {directory}{"/**" if recursive else ""}  '
            f'— pattern: {os.path.basename(pattern_path)}  '
            f'— Press Ctrl+C to stop',
            file=sys.stderr,
        )

    try:
        while not handler.should_stop():
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        observer.stop()
        observer.join()
        if not quiet:
            print('\nStopped.', file=sys.stderr)
