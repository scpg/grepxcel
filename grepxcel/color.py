"""ANSI colouring for interactive terminals.

Colour is applied only when writing to a TTY and the ``NO_COLOR`` environment
variable is unset (see https://no-color.org); ``GREPXCEL_FORCE_COLOR`` overrides
the TTY check (useful for piping into a pager that understands ANSI, and for
tests). The same semantic palette is reused across the verbose trace, the run
summary, ``doctor`` and ``validate-pattern`` so status marks read consistently
everywhere. File logs and piped/redirected output stay plain text.

The public surface is deliberately small:

  * ``should_color(stream)`` — decide whether to colour for a given stream.
  * ``paint(text, colour, enabled)`` — wrap a span in one colour.
  * ``colorize_marks(text, enabled)`` — colour the status glyphs (✓ ✗ ⚠) found
    anywhere in a line. Call sites keep emitting plain glyphs; colouring is
    applied once at the output boundary, so the file log never sees ANSI codes.
"""
import os

_RESET = '\033[0m'
_CODES = {
    'green':  '\033[92m',   # bright green  (vivid success signal)
    'red':    '\033[91m',   # bright red    (vivid error/failure)
    'yellow': '\033[93m',   # bright yellow (vivid warning/anchor)
    'cyan':   '\033[96m',   # bright cyan   (vivid label/field)

    'dim':    '\033[2m',
    'bold':   '\033[1m',
}

# Status glyphs → semantic colour. Any line carrying a mark is coloured the same
# way regardless of which subsystem emitted it.
_GLYPH_COLOR = {'✓': 'green', '✗': 'red', '⚠': 'yellow'}


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
    """Colour each status glyph (✓ ✗ ⚠) in ``text`` with its semantic colour."""
    if not enabled:
        return text
    for glyph, color in _GLYPH_COLOR.items():
        if glyph in text:
            text = text.replace(glyph, f'{_CODES[color]}{glyph}{_RESET}')
    return text
