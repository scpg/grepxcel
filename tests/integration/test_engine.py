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

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'inv', 'client', 'amount'}

    def test_no_tables_key(self):
        assert 'tables' not in self.result

    def test_invoice_number(self):
        assert self.result['inv']['number'] == 'AB123456'

    def test_invoice_date_is_date(self):
        d = self.result['inv']['date']
        assert isinstance(d, (datetime.date, datetime.datetime))

    def test_invoice_due_date_is_date(self):
        d = self.result['inv']['due_date']
        assert isinstance(d, (datetime.date, datetime.datetime))

    def test_client_name(self):
        assert self.result['client']['name'] == 'Alice Wonderland'

    def test_client_email(self):
        assert self.result['client']['email'] == 'alice@wonderland.example'

    def test_net_amount(self):
        assert self.result['amount']['net'] == 120.0

    def test_vat_amount(self):
        assert self.result['amount']['vat'] == 24.0

    def test_total_amount(self):
        assert self.result['amount']['gross'] == 144.0


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

    def test_only_item_key(self):
        assert set(self.result.keys()) == {'item'}

    def test_three_instances_found(self):
        assert len(self.result['item']) == 3

    def test_electronics_row_count(self):
        assert len(self.result['item'][0]['data']) == 3

    def test_stationery_row_count(self):
        assert len(self.result['item'][1]['data']) == 2

    def test_furniture_row_count(self):
        assert len(self.result['item'][2]['data']) == 3

    def test_no_header_in_output(self):
        # col_product etc. are lbl: fields — never in output
        for instance in self.result['item']:
            assert 'header' not in instance

    def test_data_row_keys(self):
        first_row = self.result['item'][0]['data'][0]
        assert set(first_row.keys()) == {'name', 'sku', 'qty', 'price'}

    def test_first_item_sku(self):
        assert self.result['item'][0]['data'][0]['sku'] == 'ELC001'

    def test_all_skus_valid(self):
        import re
        for instance in self.result['item']:
            for row in instance['data']:
                assert re.fullmatch(r'[A-Z]{3}[0-9]{3}', row['sku'])

    def test_quantities_are_integers(self):
        for instance in self.result['item']:
            for row in instance['data']:
                assert isinstance(row['qty'], int)


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

    def test_cell_groups_present(self):
        assert set(self.result.keys()) >= {'po', 'vendor', 'row'}

    def test_po_group_keys(self):
        assert set(self.result['po'].keys()) == {'number', 'date'}

    def test_vendor_group_keys(self):
        assert set(self.result['vendor'].keys()) == {'name', 'ref'}

    def test_no_label_fields_in_output(self):
        # po_label, date_label, vendor_label, ref_label are lbl: — must not appear
        for key in ('po_label', 'date_label', 'vendor_label', 'ref_label'):
            assert key not in self.result

    def test_po_number(self):
        assert self.result['po']['number'] == 'PO-2026'

    def test_vendor_name(self):
        assert self.result['vendor']['name'] == 'Acme Supplies'

    def test_vendor_ref(self):
        assert self.result['vendor']['ref'] == 'ACME001'

    def test_po_date_is_date(self):
        d = self.result['po']['date']
        assert isinstance(d, (datetime.date, datetime.datetime))

    def test_one_table_instance(self):
        assert len(self.result['row']) == 1

    def test_table_has_three_data_rows(self):
        assert len(self.result['row'][0]['data']) == 3

    def test_no_header_in_output(self):
        # col_item etc. are lbl: fields — never in output
        assert 'header' not in self.result['row'][0]

    def test_table_has_footer(self):
        f = self.result['row'][0]['footer']
        assert f['label'] == 'Grand Total'
        assert f['value'] == 3030.0

    def test_first_line_item(self):
        row = self.result['row'][0]['data'][0]
        assert row['item'] == 'Laptop'
        assert row['qty'] == 2
        assert row['total'] == 2400.0

    def test_short_name_item_still_extracted(self):
        # "X" fails the regex but DATA rows are lenient — value is still extracted
        row = self.result['row'][0]['data'][1]
        assert row['item'] == 'X'
