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

Palette notes
-------------
We use 256-colour extended codes (``\\033[38;5;Nm``) instead of the 16-colour
high-intensity variants (90–97) for all semantic colours.  High-intensity codes
rely on the terminal's configurable 16-colour palette, which many themes remap
to unexpected or near-invisible hues.  256-colour codes address a fixed cube
that is immune to palette remapping and supported by every modern terminal
(gnome-terminal, xterm, iTerm2, Windows Terminal, …).
"""
import os

# ── Status emoji (self-coloured; no ANSI needed) ─────────────────────────────
MARK_OK   = '🟢'   # success / VALID / extraction match
MARK_WARN = '🟡'   # warning / lbl-anchor / conditional
MARK_FAIL = '🔴'   # error / INVALID / fatal
MARK_INFO = '🔵'   # informational / neutral

_RESET = '\033[0m'
_CODES = {
    # 256-colour palette — immune to terminal theme remapping of the 16-colour
    # high-intensity range (90–97).  Indices from the standard xterm-256 cube.
    'green':      '\033[38;5;82m',    # vivid lime-green  — extracted data / var role
    'red':        '\033[38;5;196m',   # vivid red         — errors / failures
    'yellow':     '\033[38;5;220m',   # vivid amber       — anchors / warnings / lbl role
    'cyan':       '\033[38;5;51m',    # vivid cyan        — field names / structural labels
    # Standard SGR codes (no palette dependency)
    'dim':        '\033[2m',          # secondary / metadata
    'bold':       '\033[1m',          # headings / filenames
    # Combined: verdict words need maximum contrast — bold + fixed hue
    'bold_green': '\033[1;38;5;82m',  # bold vivid green  — VALID
    'bold_red':   '\033[1;38;5;196m', # bold vivid red    — INVALID
}

# Emoji marks are self-coloured (no ANSI needed), but must be replaced with
# plain ASCII when the output stream is not a TTY.
_GLYPH_COLOR: dict = {}

_MARK_PLAIN: dict[str, str] = {
    MARK_OK:   '+',
    MARK_WARN: '!',
    MARK_FAIL: 'x',
    MARK_INFO: 'i',
}


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
    """Replace emoji status marks with ASCII equivalents when colour is disabled.

    When *enabled* is False (non-TTY / piped output) each emoji mark is
    substituted with its plain ASCII counterpart (🟢→+ 🔴→x 🟡→! 🔵→i)
    so pipelines and log files stay clean.  When *enabled* is True the emoji
    are left as-is (they are self-coloured and need no ANSI wrapping).
    """
    if not enabled:
        for emoji, plain in _MARK_PLAIN.items():
            text = text.replace(emoji, plain)
    return text
