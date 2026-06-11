"""Tests for the `docs` reference generator (grepxcel.docs_generator)."""
import openpyxl

from grepxcel.docs_generator import DocsGenerator


def test_output_is_byte_deterministic(tmp_path):
    """Two generations must be byte-identical so the docs-reference CI check
    (a byte diff) is meaningful — openpyxl otherwise stamps now()."""
    a = tmp_path / 'a.xlsx'
    b = tmp_path / 'b.xlsx'
    DocsGenerator().write(str(a))
    DocsGenerator().write(str(b))
    assert a.read_bytes() == b.read_bytes()


def test_output_is_a_valid_workbook(tmp_path):
    out = tmp_path / 'ref.xlsx'
    DocsGenerator().write(str(out))
    wb = openpyxl.load_workbook(out)
    assert wb.active.title == 'pattern-reference'


def test_source_date_epoch_is_honored(tmp_path, monkeypatch):
    monkeypatch.setenv('SOURCE_DATE_EPOCH', '1577836800')  # 2020-01-01T00:00:00Z
    out = tmp_path / 'ref.xlsx'
    DocsGenerator().write(str(out))
    wb = openpyxl.load_workbook(out)
    assert wb.properties.created.year == 2020


def test_lists_the_new_field_types(tmp_path):
    # The reference must mention the types added later (text/number/boolean).
    out = tmp_path / 'ref.xlsx'
    DocsGenerator().write(str(out))
    wb = openpyxl.load_workbook(out)
    text = ' '.join(
        str(c.value) for row in wb.active.iter_rows() for c in row if c.value
    )
    for t in ('text', 'number', 'boolean'):
        assert t in text
