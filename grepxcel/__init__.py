from .engine import Engine
from .logger import Logger, VerbosityLevel

__version__ = '0.1.1'


def extract(pattern, data, *, sheet=None, all_sheets=False,
            output_format='nested', logger=None):
    """Extract data from an Excel file using a pattern — the one-call API.

    This is the recommended entry point for programmatic use; it wraps
    :class:`Engine` and returns the extracted data directly.

        import grepxcel
        result = grepxcel.extract("pattern.xlsx", "data.xlsx")

    Args:
        pattern: path (``str`` or ``Path``) to the pattern file (``.xlsx`` or ``.csv``).
        data:    path (``str`` or ``Path``) to the data file (``.xlsx``).
        sheet:   sheet name or 0-based index to read; default is the active sheet.
                 Mutually exclusive with ``all_sheets``.
        all_sheets: if True, process every sheet and return ``{sheet_name: result}``.
        output_format: ``'nested'`` (default) or ``'legacy'``.
        logger:  a :class:`Logger` for progress/warnings. By default extraction is
                 silent (``VerbosityLevel.QUIET``) — pass your own logger to see output.

    Returns:
        The extracted data as a ``dict`` (keyed by sheet name when ``all_sheets``).

    For finer control (custom security limits, reusing one engine across calls),
    use :class:`Engine` directly.
    """
    if all_sheets and sheet is not None:
        raise ValueError("pass either `sheet` or `all_sheets=True`, not both")

    if logger is None:
        logger = Logger(level=VerbosityLevel.QUIET)

    engine = Engine()
    if all_sheets:
        return engine.process_all(str(pattern), str(data),
                                  logger=logger, output_format=output_format)
    return engine.process(str(pattern), str(data),
                          logger=logger, sheet=sheet, output_format=output_format)


__all__ = ['Engine', 'Logger', 'VerbosityLevel', 'extract', '__version__']
