"""Unit tests for grepxcel.extract_df() DataFrame facade."""

import os
import sys

import pytest

import grepxcel
from tests.conftest import find_data_file, find_pattern_xlsx


def _can_import(name: str) -> bool:
    try:
        __import__(name)
        return True
    except ImportError:
        return False

_INVOICE_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '01_simple_invoice')
_INVOICE_PAT = find_pattern_xlsx(_INVOICE_FIX)
_INVOICE_DATA = find_data_file(_INVOICE_FIX)

_CATALOG_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '02_product_catalog')
_CATALOG_PAT = find_pattern_xlsx(_CATALOG_FIX)
_CATALOG_DATA = find_data_file(_CATALOG_FIX)

_EXPENSE_FIX = os.path.join(os.path.dirname(__file__), '..', 'fixtures', '05_expense_report')
_EXPENSE_PAT = find_pattern_xlsx(_EXPENSE_FIX)
_EXPENSE_DATA = find_data_file(_EXPENSE_FIX)


# ── export / importability ────────────────────────────────────────────────────

def test_extract_df_is_exported():
    assert hasattr(grepxcel, 'extract_df')
    assert 'extract_df' in grepxcel.__all__


# ── no backend installed ──────────────────────────────────────────────────────

def test_extract_df_raises_without_backend(monkeypatch):
    import builtins
    real_import = builtins.__import__

    def block_df_libs(name, *args, **kwargs):
        if name in ('pandas', 'polars'):
            raise ImportError(f"No module named '{name}'")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, '__import__', block_df_libs)
    with pytest.raises(ImportError, match='pandas.*polars'):
        grepxcel.extract_df(_INVOICE_PAT, _INVOICE_DATA)


# ── pandas backend ────────────────────────────────────────────────────────────

@pytest.mark.skipif(
    not _can_import('pandas'),
    reason='pandas not installed',
)
class TestPandas:
    def test_scalar_only_returns_single_row_df(self):
        import pandas as pd
        frames = grepxcel.extract_df(_INVOICE_PAT, _INVOICE_DATA, backend='pandas')
        assert isinstance(frames, dict)
        assert '_scalars' in frames
        df = frames['_scalars']
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 1

    def test_table_returns_dataframe_per_table_key(self):
        import pandas as pd
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='pandas')
        assert isinstance(frames, dict)
        table_keys = [k for k in frames if k != '_scalars']
        assert len(table_keys) >= 1
        for key in table_keys:
            df = frames[key]
            assert isinstance(df, pd.DataFrame)
            assert len(df) > 0

    def test_mixed_scalar_and_table(self):
        frames = grepxcel.extract_df(_EXPENSE_PAT, _EXPENSE_DATA, backend='pandas')
        assert '_scalars' in frames
        table_keys = [k for k in frames if k != '_scalars']
        assert len(table_keys) >= 1

    def test_table_data_rows_are_concatenated(self):
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='pandas')
        table_keys = [k for k in frames if k != '_scalars']
        df = frames[table_keys[0]]
        assert len(df) >= 3

    def test_auto_backend_picks_pandas(self):
        import pandas as pd
        frames = grepxcel.extract_df(_INVOICE_PAT, _INVOICE_DATA)
        assert isinstance(frames, dict)
        df = frames['_scalars']
        assert isinstance(df, pd.DataFrame)

    def test_all_sheets(self):
        result = grepxcel.extract_df(_INVOICE_PAT, _INVOICE_DATA,
                                     backend='pandas', all_sheets=True)
        assert isinstance(result, dict)
        for sheet_name, frames in result.items():
            assert isinstance(sheet_name, str)
            assert isinstance(frames, dict)

    def test_no_scalars_key_when_only_tables(self):
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='pandas')
        if '_scalars' in frames:
            assert len(frames['_scalars']) == 0 or frames['_scalars'].empty

    def test_source_column_excluded(self):
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='pandas')
        table_keys = [k for k in frames if k != '_scalars']
        df = frames[table_keys[0]]
        assert '_source' not in df.columns

    def test_footer_frame_present_when_footer_exists(self):
        # expense report has per-instance footers (subtotals)
        frames = grepxcel.extract_df(_EXPENSE_PAT, _EXPENSE_DATA, backend='pandas')
        footer_keys = [k for k in frames if k.endswith('__footer')]
        assert len(footer_keys) >= 1
        fdf = frames[footer_keys[0]]
        assert len(fdf) >= 1

    def test_no_footer_frame_when_absent(self):
        # product catalog has no footers
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='pandas')
        assert not any(k.endswith('__footer') for k in frames)


# ── polars backend ────────────────────────────────────────────────────────────

@pytest.mark.skipif(
    not _can_import('polars'),
    reason='polars not installed',
)
class TestPolars:
    def test_scalar_only_returns_single_row_df(self):
        import polars as pl
        frames = grepxcel.extract_df(_INVOICE_PAT, _INVOICE_DATA, backend='polars')
        assert isinstance(frames, dict)
        assert '_scalars' in frames
        df = frames['_scalars']
        assert isinstance(df, pl.DataFrame)
        assert len(df) == 1

    def test_table_returns_dataframe_per_table_key(self):
        import polars as pl
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='polars')
        table_keys = [k for k in frames if k != '_scalars']
        assert len(table_keys) >= 1
        for key in table_keys:
            df = frames[key]
            assert isinstance(df, pl.DataFrame)
            assert len(df) > 0

    def test_mixed_scalar_and_table(self):
        frames = grepxcel.extract_df(_EXPENSE_PAT, _EXPENSE_DATA, backend='polars')
        assert '_scalars' in frames
        table_keys = [k for k in frames if k != '_scalars']
        assert len(table_keys) >= 1

    def test_table_data_rows_are_concatenated(self):
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='polars')
        table_keys = [k for k in frames if k != '_scalars']
        df = frames[table_keys[0]]
        assert len(df) >= 3

    def test_source_column_excluded(self):
        frames = grepxcel.extract_df(_CATALOG_PAT, _CATALOG_DATA, backend='polars')
        table_keys = [k for k in frames if k != '_scalars']
        df = frames[table_keys[0]]
        assert '_source' not in df.columns


# ── invalid backend ──────────────────────────────────────────────────────────

def test_extract_df_invalid_backend():
    with pytest.raises(ValueError, match='backend'):
        grepxcel.extract_df(_INVOICE_PAT, _INVOICE_DATA, backend='dask')
