"""
Integration tests: run the engine against each test fixture and assert
on the extracted structure and values.
"""

import os
import datetime
import pytest
from engine import Engine, Logger, VerbosityLevel

FIXTURES = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


def run(fixture_name: str):
    folder = os.path.join(FIXTURES, fixture_name)
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(
        pattern_file=os.path.join(folder, 'pattern.xlsx'),
        data_file=os.path.join(folder, 'data.xlsx'),
        logger=lg,
    )
    return result, lg


# ─────────────────────────────────────────────────────────────────────────────
# 01: Simple Invoice
# ─────────────────────────────────────────────────────────────────────────────

class TestSimpleInvoice:
    def setup_method(self):
        self.result, self.lg = run('01_simple_invoice')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_cell_count(self):
        assert len(self.result['cells']) == 8  # 8 named fields (no IGNOREs in output)

    def test_no_tables(self):
        assert self.result['tables'] == []

    def test_invoice_number(self):
        assert self.result['cells']['inv.number'] == 'AB123456'

    def test_invoice_date_is_date(self):
        d = self.result['cells']['inv.date']
        assert isinstance(d, (datetime.date, datetime.datetime))

    def test_client_name(self):
        assert self.result['cells']['client.name'] == 'Alice Wonderland'

    def test_client_email(self):
        assert self.result['cells']['client.email'] == 'alice@wonderland.example'

    def test_net_amount(self):
        assert self.result['cells']['amount.net'] == 120.0

    def test_total_amount(self):
        assert self.result['cells']['amount.gross'] == 144.0


# ─────────────────────────────────────────────────────────────────────────────
# 02: Product Catalog
# ─────────────────────────────────────────────────────────────────────────────

class TestProductCatalog:
    def setup_method(self):
        self.result, self.lg = run('02_product_catalog')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_no_standalone_cells(self):
        assert self.result['cells'] == {}

    def test_three_instances_found(self):
        assert len(self.result['tables']) == 3

    def test_all_same_table_group(self):
        indices = {t['table_index'] for t in self.result['tables']}
        assert indices == {0}

    def test_instance_indices(self):
        inst = [t['instance_index'] for t in self.result['tables']]
        assert inst == [0, 1, 2]

    def test_electronics_row_count(self):
        t0 = self.result['tables'][0]
        assert len(t0['data']) == 3

    def test_stationery_row_count(self):
        t1 = self.result['tables'][1]
        assert len(t1['data']) == 2

    def test_furniture_row_count(self):
        t2 = self.result['tables'][2]
        assert len(t2['data']) == 3

    def test_electronics_header(self):
        header = self.result['tables'][0]['headers'][0]
        assert header['header.product'] == 'Product'
        assert header['header.sku'] == 'SKU'

    def test_first_item_sku(self):
        first_row = self.result['tables'][0]['data'][0]
        assert first_row['item.sku'] == 'ELC001'

    def test_all_skus_valid(self):
        import re
        for table in self.result['tables']:
            for row in table['data']:
                assert re.fullmatch(r'[A-Z]{3}[0-9]{3}', row['item.sku'])

    def test_quantities_are_integers(self):
        for table in self.result['tables']:
            for row in table['data']:
                assert isinstance(row['item.qty'], int)


# ─────────────────────────────────────────────────────────────────────────────
# 03: Purchase Order
# ─────────────────────────────────────────────────────────────────────────────

class TestPurchaseOrder:
    def setup_method(self):
        self.result, self.lg = run('03_purchase_order')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_exactly_one_warning(self):
        assert len(self.lg.issues()) == 1

    def test_warning_is_for_row_item(self):
        w = self.lg.issues()[0]
        assert w.field == 'row.item'

    def test_warning_message(self):
        w = self.lg.issues()[0]
        assert 'does not match' in w.message

    def test_cell_count(self):
        # lbl.po, po.number, lbl.date, po.date, lbl.vendor, vendor.name, lbl.ref, vendor.ref
        assert len(self.result['cells']) == 8

    def test_po_number(self):
        assert self.result['cells']['po.number'] == 'PO-2026'

    def test_vendor_name(self):
        assert self.result['cells']['vendor.name'] == 'Acme Supplies'

    def test_po_date_is_date(self):
        d = self.result['cells']['po.date']
        assert isinstance(d, (datetime.date, datetime.datetime))

    def test_one_table_instance(self):
        assert len(self.result['tables']) == 1

    def test_table_has_three_data_rows(self):
        assert len(self.result['tables'][0]['data']) == 3

    def test_table_has_header(self):
        h = self.result['tables'][0]['headers'][0]
        assert h['col.item'] == 'Item'
        assert h['col.total'] == 'Total'

    def test_table_has_footer(self):
        f = self.result['tables'][0]['footers'][0]
        assert f['footer.label'] == 'Grand Total'
        assert f['footer.value'] == 3030.0

    def test_first_line_item(self):
        row = self.result['tables'][0]['data'][0]
        assert row['row.item'] == 'Laptop'
        assert row['row.qty'] == 2
        assert row['row.total'] == 2400.0

    def test_short_name_item_still_extracted(self):
        # "X" fails the regex but DATA rows are lenient — value is still extracted
        row = self.result['tables'][0]['data'][1]
        assert row['row.item'] == 'X'
