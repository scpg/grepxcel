"""Tests for grepxcel.color: TTY/NO_COLOR gating and glyph colouring."""
import io

import pytest

from grepxcel.color import colorize_marks, paint, should_color

ESC = '\033['


class _FakeTTY(io.StringIO):
    def isatty(self):
        return True


class TestShouldColor:
    def test_tty_enables(self, monkeypatch):
        monkeypatch.delenv('NO_COLOR', raising=False)
        monkeypatch.delenv('GREPXCEL_FORCE_COLOR', raising=False)
        assert should_color(_FakeTTY()) is True

    def test_non_tty_disables(self, monkeypatch):
        monkeypatch.delenv('NO_COLOR', raising=False)
        monkeypatch.delenv('GREPXCEL_FORCE_COLOR', raising=False)
        assert should_color(io.StringIO()) is False  # plain buffer, not a tty

    def test_no_color_wins_over_tty(self, monkeypatch):
        monkeypatch.setenv('NO_COLOR', '1')
        assert should_color(_FakeTTY()) is False

    def test_no_color_wins_over_force(self, monkeypatch):
        monkeypatch.setenv('NO_COLOR', '')        # presence, even empty, disables
        monkeypatch.setenv('GREPXCEL_FORCE_COLOR', '1')
        assert should_color(_FakeTTY()) is False

    def test_force_color_on_non_tty(self, monkeypatch):
        monkeypatch.delenv('NO_COLOR', raising=False)
        monkeypatch.setenv('GREPXCEL_FORCE_COLOR', '1')
        assert should_color(io.StringIO()) is True


class TestColorizeMarks:
    def test_disabled_is_identity(self):
        line = '  ✓  ready'
        assert colorize_marks(line, enabled=False) == line

    def test_each_glyph_wrapped(self):
        for glyph in ('✓', '✗', '⚠'):
            out = colorize_marks(f'x {glyph} y', enabled=True)
            assert ESC in out and glyph in out and out.endswith('y')

    def test_plain_line_unchanged(self):
        assert colorize_marks('no marks here', enabled=True) == 'no marks here'

    def test_paint_disabled_is_identity(self):
        assert paint('hi', 'red', enabled=False) == 'hi'

    def test_paint_unknown_color_is_identity(self):
        assert paint('hi', 'chartreuse', enabled=True) == 'hi'
