"""ANSI colouring for interactive terminals.

Colour is applied only when writing to a TTY and the ``NO_COLOR`` environment
variable is unset (see https://no-color.org); ``GREPXCEL_FORCE_COLOR`` overrides
the TTY check (useful for piping into a pager that understands ANSI, and for
tests). The same semantic palette is reused across the verbose trace, the run
summary, ``doctor`` and ``validate-pattern`` so status marks read consistently
everywhere. File logs and piped/redirected output stay plain text.

The public surface is deliberately small:

  * ``MARK_OK / MARK_WARN / MARK_FAIL / MARK_INFO`` — canonical status emoji.
    Import these instead of bare Unicode glyphs; one change here updates the
    whole CLI consistently.
  * ``should_color(stream)`` — decide whether to colour for a given stream.
  * ``paint(text, colour, enabled)`` — wrap a span in one colour.
  * ``colorize_marks(text, enabled)`` — no-op kept for call-site compatibility;
    emoji circles are self-coloured and need no ANSI wrapping.
"""
import os

# ── Status emoji (self-coloured; no ANSI needed) ─────────────────────────────
MARK_OK   = '🟢'   # success / VALID / extraction match
MARK_WARN = '🟡'   # warning / lbl-anchor / conditional
MARK_FAIL = '🔴'   # error / INVALID / fatal
MARK_INFO = '🔵'   # informational / neutral

_RESET = '\033[0m'
_CODES = {
    'green':      '\033[92m',   # bright green  (vivid success signal)
    'red':        '\033[91m',   # bright red    (vivid error/failure)
    'yellow':     '\033[93m',   # bright yellow (vivid warning/anchor)
    'cyan':       '\033[96m',   # bright cyan   (vivid label/field)
    'dim':        '\033[2m',
    'bold':       '\033[1m',
    # Combined codes for verdict words — use standard (30-37) not bright (90-97)
    # so they stay a true saturated hue regardless of the terminal's palette theme.
    'bold_green': '\033[1;32m', # bold + standard green → VALID
    'bold_red':   '\033[1;31m', # bold + standard red   → INVALID
}

# Emoji marks are self-coloured; _GLYPH_COLOR is kept empty so colorize_marks()
# is a harmless pass-through after the switch to emoji.
_GLYPH_COLOR: dict = {}


def should_color(stream) -> bool:
    """Return True when ANSI colour is appropriate for ``stream``."""
    if os.environ.get('NO_COLOR') is not None:
        return False
    if os.environ.get('GREPXCEL_FORCE_COLOR'):
        return True
    isatty = getattr(stream, 'isatty', None)
    return bool(isatty()) if callable(isatty) else False


def paint(text: str, color: str, enabled: bool = True) -> str:
    """Wrap ``text`` in a single colour; a no-op when disabled or unknown."""
    if not enabled or color not in _CODES:
        return text
    return f'{_CODES[color]}{text}{_RESET}'


def colorize_marks(text: str, enabled: bool = True) -> str:
    """Pass-through kept for call-site compatibility.

    Previously coloured ``✓ ✗ ⚠`` glyphs with ANSI codes; those glyphs have
    been replaced with self-coloured emoji (🟢 🔴 🟡 🔵) that need no ANSI
    wrapping.  Call sites that still pass the ``enabled`` flag are unaffected.
    """
    return text
