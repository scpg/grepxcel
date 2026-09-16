from .engine import Engine
from .logger import Logger, VerbosityLevel
from .utils import flatten_nested as _flatten_pairs
from .utils import flatten_table_instances as _flatten_table_instances

__version__ = '0.3.1'


def extract(pattern, data, *, sheet=None, all_sheets=False,
            output_format='nested', flat_tables=False, logger=None):
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
        flat_tables: if True, collapse table instance wrappers so each table key
                 maps directly to a list of row dicts instead of a list of
                 ``{"_source": ..., "data": [...]}`` envelopes.  Makes it easier
                 to feed table data into a database or pandas without unwrapping.
                 Ignored when ``output_format='legacy'``.
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
        raw = engine.process_all(str(pattern), str(data),
                                 logger=logger, output_format=output_format)
        if flat_tables and output_format != 'legacy':
            return {name: _flatten_table_instances(sheet_data)
                    for name, sheet_data in raw.items()}
        return raw

    raw = engine.process(str(pattern), str(data),
                         logger=logger, sheet=sheet, output_format=output_format)
    if flat_tables and output_format != 'legacy':
        return _flatten_table_instances(raw)
    return raw


def extract_df(pattern, data, *, backend=None, sheet=None, all_sheets=False,
               output_format='nested', logger=None):
    """Extract data from an Excel file and return DataFrames.

    Wraps :func:`extract` and converts the result into a ``dict`` of
    DataFrames — one per table key, plus ``'_scalars'`` for non-table fields.

        import grepxcel
        frames = grepxcel.extract_df("pattern.xlsx", "data.xlsx")
        frames['items']  # pandas or polars DataFrame

    Args:
        pattern: path to the pattern file (``.xlsx`` or ``.csv``).
        data:    path to the data file (``.xlsx``).
        backend: ``'pandas'``, ``'polars'``, or ``None`` (auto-detect).
        sheet:   sheet name or 0-based index; default is the active sheet.
        all_sheets: if True, process every sheet; returns
                    ``{sheet_name: {key: DataFrame}}``.
        output_format: passed through to :func:`extract`.
        logger:  a :class:`Logger` instance; silent by default.

    Returns:
        ``dict[str, DataFrame]`` — one DataFrame per table key, plus
        ``'_scalars'`` for scalar fields (omitted when there are none).
        When ``all_sheets=True``, returns
        ``dict[str, dict[str, DataFrame]]``.

    Raises:
        ImportError: if neither pandas nor polars is installed.
        ValueError:  if *backend* is not one of the accepted values.
    """
    pd, pl = None, None
    if backend == 'pandas':
        import pandas as pd
    elif backend == 'polars':
        import polars as pl
    elif backend is None:
        try:
            import pandas as pd
        except ImportError:
            try:
                import polars as pl
            except ImportError:
                raise ImportError(
                    "extract_df() requires pandas or polars. "
                    "Install one with: pip install 'grepxcel[pandas]' "
                    "or pip install 'grepxcel[polars]'"
                )
    else:
        raise ValueError(
            f"backend must be 'pandas', 'polars', or None, got {backend!r}"
        )

    raw = extract(pattern, data, sheet=sheet, all_sheets=all_sheets,
                  output_format=output_format, logger=logger)

    if all_sheets:
        return {name: _to_frames(sheet_data, pd, pl)
                for name, sheet_data in raw.items()}
    return _to_frames(raw, pd, pl)


def _make_df(rows, pd, pl):
    return pd.DataFrame(rows) if pd is not None else pl.DataFrame(rows)


def _to_frames(result, pd, pl):
    """Convert an extract() result dict into ``{key: DataFrame}``.

    Per-table ``data`` rows form the main frame for each table key. When a table
    also carries ``header`` or ``footer`` fields (e.g. subtotals), those are
    exposed as separate ``{key}__header`` / ``{key}__footer`` frames — one row
    per instance — so callers can reach subtotals without parsing the raw JSON.
    Scalar fields collapse into a single-row ``_scalars`` frame.
    """
    frames = {}
    scalars = {}

    for key, value in result.items():
        if key == '_meta':
            continue
        if isinstance(value, list):
            rows = []
            headers = []
            footers = []
            for instance in value:
                for data_row in instance.get('data', []):
                    rows.append(data_row)
                if instance.get('header'):
                    headers.append(dict(_flatten_pairs(instance['header'])))
                if instance.get('footer'):
                    footers.append(dict(_flatten_pairs(instance['footer'])))
            if rows:
                frames[key] = _make_df(rows, pd, pl)
            if headers:
                frames[f'{key}__header'] = _make_df(headers, pd, pl)
            if footers:
                frames[f'{key}__footer'] = _make_df(footers, pd, pl)
        elif isinstance(value, dict):
            scalars.update(dict(_flatten_pairs(value, key)))
        else:
            scalars[key] = value

    if scalars:
        frames['_scalars'] = _make_df([scalars], pd, pl)

    return frames


__all__ = ['Engine', 'Logger', 'VerbosityLevel', 'extract', 'extract_df',
           '__version__']
