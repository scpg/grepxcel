"""
Integration tests: run the engine against each test fixture and assert
on the extracted structure and values.
"""

import os
import tempfile
import datetime
import pytest
import openpyxl as _openpyxl
from grepxcel import Engine, Logger, VerbosityLevel
from tests.conftest import find_pattern_xlsx

FIXTURES = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


def _write_pattern(rows: list) -> str:
    """Write pattern rows to a temp xlsx; caller must os.unlink() it."""
    wb = _openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    fd, path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    wb.save(path)
    return path


def run_custom_pattern(pattern_rows: list, fixture_name: str, sheet=None):
    """Run engine with in-memory pattern against an existing fixture data.xlsx."""
    pat = _write_pattern(pattern_rows)
    data_file = os.path.join(FIXTURES, fixture_name, 'data.xlsx')
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    result = Engine().process(pattern_file=pat, data_file=data_file, logger=lg, **kwargs)
    os.unlink(pat)
    return result, lg


def run(fixture_name: str, sheet=None):
    folder = os.path.join(FIXTURES, fixture_name)
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    result = Engine().process(
        pattern_file=find_pattern_xlsx(folder),
        data_file=os.path.join(folder, 'data.xlsx'),
        logger=lg,
        **kwargs,
    )
    return result, lg


def run_with_pattern_file(fixture_name: str, pattern_file: str, sheet=None):
    """Run engine with an explicit pattern file path against a fixture's data.xlsx."""
    folder = os.path.join(FIXTURES, fixture_name)
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    result = Engine().process(
        pattern_file=os.path.join(folder, pattern_file),
        data_file=os.path.join(folder, 'data.xlsx'),
        logger=lg,
        **kwargs,
    )
    return result, lg


def run_all_sheets(fixture_name: str):
    folder = os.path.join(FIXTURES, fixture_name)
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process_all(
        pattern_file=find_pattern_xlsx(folder),
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
# 02b: Source Provenance (_source on table instances)
# ─────────────────────────────────────────────────────────────────────────────

class TestSourceProvenance:
    """
    Verify that every table instance carries _source with sheet name and
    A1-notation ref. Uses fixture 02 (three plain mini-tables, no header var:).
    """
    def setup_method(self):
        self.result, _ = run('02_product_catalog')
        self.instances = self.result['item']

    def test_all_instances_have_source(self):
        for inst in self.instances:
            assert '_source' in inst

    def test_source_keys(self):
        for inst in self.instances:
            assert set(inst['_source'].keys()) == {'sheet', 'ref'}

    def test_source_sheet_name(self):
        for inst in self.instances:
            assert inst['_source']['sheet'] == 'Sheet1'

    def test_source_ref_a1_notation(self):
        import re
        pattern = re.compile(r'^[A-Z]+\d+:[A-Z]+\d+$')
        for inst in self.instances:
            assert pattern.match(inst['_source']['ref'])

    def test_source_refs_are_distinct(self):
        refs = [inst['_source']['ref'] for inst in self.instances]
        assert len(set(refs)) == len(self.instances)

    def test_first_instance_ref(self):
        # HEADER row 1 → data rows 2–4 → blank separator row 5
        assert self.instances[0]['_source']['ref'] == 'A1:D5'

    def test_second_instance_ref(self):
        assert self.instances[1]['_source']['ref'] == 'A6:D9'

    def test_third_instance_ref(self):
        # Last table: trailing empty row 14 marks end-of-data
        assert self.instances[2]['_source']['ref'] == 'A10:D14'

    def test_refs_are_non_overlapping(self):
        # Extract start rows from each ref and verify monotonic ordering
        def start_row(ref):
            return int(ref.split(':')[0].lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
        rows = [start_row(inst['_source']['ref']) for inst in self.instances]
        assert rows == sorted(rows)

    def test_source_absent_from_scalar_cells(self):
        # _source is only added to table instances, not to scalar cell output
        for key, val in self.result.items():
            if key != 'item':
                assert not isinstance(val, dict) or '_source' not in val


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


# ─────────────────────────────────────────────────────────────────────────────
# 04: Bank Statement
# ─────────────────────────────────────────────────────────────────────────────

class TestBankStatement:
    def setup_method(self):
        self.result, self.lg = run('04_bank_statement')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'account', 'txn'}

    def test_account_holder(self):
        assert self.result['account']['holder'] == 'Jane Smith'

    def test_account_number(self):
        assert self.result['account']['number'] == 'GB29NWBK60161331926819'

    def test_account_period(self):
        assert self.result['account']['period'] == 'January 2026'

    def test_one_table_instance(self):
        assert len(self.result['txn']) == 1

    def test_five_data_rows(self):
        assert len(self.result['txn'][0]['data']) == 5

    def test_first_row_description(self):
        assert self.result['txn'][0]['data'][0]['description'] == 'Opening Balance'

    def test_footer_label(self):
        assert self.result['txn'][0]['footer']['label'] == 'Totals'

    def test_footer_debits(self):
        assert self.result['txn'][0]['footer']['debits'] == 455.5

    def test_footer_credits(self):
        assert self.result['txn'][0]['footer']['credits'] == 5700.0

    def test_no_header_in_output(self):
        assert 'header' not in self.result['txn'][0]

    def test_data_row_keys(self):
        keys = set(self.result['txn'][0]['data'][0].keys())
        assert keys == {'date', 'description', 'debit', 'credit', 'balance'}


# ─────────────────────────────────────────────────────────────────────────────
# 05: Expense Report
# ─────────────────────────────────────────────────────────────────────────────

class TestExpenseReport:
    def setup_method(self):
        self.result, self.lg = run('05_expense_report')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'emp', 'exp'}

    def test_employee_name(self):
        assert self.result['emp']['name'] == 'John Smith'

    def test_employee_department(self):
        assert self.result['emp']['department'] == 'Engineering'

    def test_employee_period(self):
        assert self.result['emp']['period'] == 'Q1 2026'

    def test_employee_manager(self):
        assert self.result['emp']['manager'] == 'Jane Doe'

    def test_three_table_instances(self):
        assert len(self.result['exp']) == 3

    def test_travel_row_count(self):
        assert len(self.result['exp'][0]['data']) == 3

    def test_meals_row_count(self):
        assert len(self.result['exp'][1]['data']) == 3

    def test_accommodation_row_count(self):
        assert len(self.result['exp'][2]['data']) == 2

    def test_travel_subtotal(self):
        assert self.result['exp'][0]['footer']['amount'] == 450.0

    def test_meals_subtotal(self):
        assert self.result['exp'][1]['footer']['amount'] == 168.5

    def test_accommodation_subtotal(self):
        assert self.result['exp'][2]['footer']['amount'] == 390.0

    def test_all_footers_have_label(self):
        for inst in self.result['exp']:
            assert inst['footer']['label'] == 'Subtotal'

    def test_receipt_numbers_valid(self):
        import re
        for inst in self.result['exp']:
            for row in inst['data']:
                assert re.fullmatch(r'REC[0-9]+', row['receipt'])


# ─────────────────────────────────────────────────────────────────────────────
# 06: Merged Cells (horizontal title + vertical category column)
# ─────────────────────────────────────────────────────────────────────────────

class TestMergedCells:
    def setup_method(self):
        self.result, self.lg = run('06_merged_cells', sheet='2026')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'report', 'item'}

    def test_horizontal_merge_title(self):
        assert self.result['report']['title'] == 'SALES CATALOGUE 2026'

    def test_one_table_instance(self):
        assert len(self.result['item']) == 1

    def test_six_data_rows(self):
        assert len(self.result['item'][0]['data']) == 6

    def test_vertical_merge_hardware_category(self):
        rows = self.result['item'][0]['data']
        # A3:A5 merged — rows 0,1,2 must all carry 'Hardware'
        for row in rows[:3]:
            assert row['category'] == 'Hardware', \
                f"Expected 'Hardware', got {row['category']!r}"

    def test_vertical_merge_software_category(self):
        rows = self.result['item'][0]['data']
        # A6:A8 merged — rows 3,4,5 must all carry 'Software'
        for row in rows[3:]:
            assert row['category'] == 'Software', \
                f"Expected 'Software', got {row['category']!r}"

    def test_no_none_categories(self):
        for row in self.result['item'][0]['data']:
            assert row['category'] is not None

    def test_footer_grand_total(self):
        footer = self.result['item'][0]['footer']
        assert footer['label'] == 'Grand Total'
        assert abs(footer['total'] - 2369.98) < 0.01

    def test_no_header_in_output(self):
        assert 'header' not in self.result['item'][0]


class TestMergedCells2025:
    """Second sheet ('2025'): 3 side-by-side tables, each with Hardware/Software/Others
    category columns (vertically merged). Prices double with each table.
    table:* in the pattern finds all three instances."""

    def setup_method(self):
        self.result, self.lg = run('06_merged_cells', sheet='2025')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_horizontal_merge_title(self):
        assert self.result['report']['title'] == 'SALES CATALOGUE 2025'

    def test_three_instances_found(self):
        assert len(self.result['item']) == 3

    def test_each_instance_ten_data_rows(self):
        for inst in self.result['item']:
            assert len(inst['data']) == 10

    def test_categories_all_instances(self):
        expected = (
            ['Hardware'] * 3 + ['Software'] * 3 + ['Others'] * 4
        )
        for inst in self.result['item']:
            assert [r['category'] for r in inst['data']] == expected

    def test_footer_totals_double_each_table(self):
        totals = [inst['footer']['total'] for inst in self.result['item']]
        assert abs(totals[0] - 3369.98) < 0.01
        assert abs(totals[1] - 6739.96) < 0.01
        assert abs(totals[2] - 13479.92) < 0.01

    def test_all_footers_grand_total_label(self):
        for inst in self.result['item']:
            assert inst['footer']['label'] == 'Grand Total'


# ─────────────────────────────────────────────────────────────────────────────
# 07: Timesheet
# ─────────────────────────────────────────────────────────────────────────────

class TestTimesheet:
    def setup_method(self):
        self.result, self.lg = run('07_timesheet')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'emp', 'day'}

    def test_employee_name(self):
        assert self.result['emp']['name'] == 'Bob Martin'

    def test_employee_project(self):
        assert self.result['emp']['project'] == 'grepxcel v2'

    def test_week_start_is_date(self):
        d = self.result['emp']['week_start']
        assert isinstance(d, (datetime.date, datetime.datetime))

    def test_one_table_instance(self):
        assert len(self.result['day']) == 1

    def test_seven_data_rows(self):
        assert len(self.result['day'][0]['data']) == 7

    def test_day_names(self):
        names = [row['name'] for row in self.result['day'][0]['data']]
        assert names == ['Monday', 'Tuesday', 'Wednesday', 'Thursday',
                         'Friday', 'Saturday', 'Sunday']

    def test_zero_hours_weekend(self):
        rows = self.result['day'][0]['data']
        assert rows[5]['hours'] == 0.0  # Saturday
        assert rows[6]['hours'] == 0.0  # Sunday

    def test_footer_total_hours(self):
        f = self.result['day'][0]['footer']
        assert f['label'] == 'Total Hours'
        assert f['hours'] == 37.0


# ─────────────────────────────────────────────────────────────────────────────
# 08: Price List
# ─────────────────────────────────────────────────────────────────────────────

class TestPriceList:
    def setup_method(self):
        self.result, self.lg = run('08_price_list')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'supplier', 'validity', 'prod'}

    def test_supplier_name(self):
        assert self.result['supplier']['name'] == 'TechDistrib GmbH'

    def test_supplier_ref(self):
        assert self.result['supplier']['ref'] == 'TD-2026-Q2'

    def test_validity_dates_are_dates(self):
        assert isinstance(self.result['validity']['from'],
                          (datetime.date, datetime.datetime))
        assert isinstance(self.result['validity']['to'],
                          (datetime.date, datetime.datetime))

    def test_three_table_instances(self):
        assert len(self.result['prod']) == 3

    def test_electronics_count(self):
        assert len(self.result['prod'][0]['data']) == 3

    def test_cables_count(self):
        assert len(self.result['prod'][1]['data']) == 4

    def test_storage_count(self):
        assert len(self.result['prod'][2]['data']) == 2

    def test_no_footer(self):
        for inst in self.result['prod']:
            assert 'footer' not in inst

    def test_product_codes_valid(self):
        import re
        for inst in self.result['prod']:
            for row in inst['data']:
                assert re.fullmatch(r'[A-Z]{2}[0-9]{4}', row['code'])

    def test_data_row_keys(self):
        keys = set(self.result['prod'][0]['data'][0].keys())
        assert keys == {'code', 'name', 'unit', 'price'}


# ─────────────────────────────────────────────────────────────────────────────
# 09: Sales by Region
# ─────────────────────────────────────────────────────────────────────────────

class TestSalesByRegion:
    def setup_method(self):
        self.result, self.lg = run('09_sales_by_region')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'report', 'sale'}

    def test_report_period(self):
        assert self.result['report']['period'] == 'Q1 2026'

    def test_report_currency(self):
        assert self.result['report']['currency'] == 'EUR'

    def test_three_table_instances(self):
        assert len(self.result['sale']) == 3

    def test_europe_row_count(self):
        assert len(self.result['sale'][0]['data']) == 3

    def test_americas_row_count(self):
        assert len(self.result['sale'][1]['data']) == 2

    def test_asia_row_count(self):
        assert len(self.result['sale'][2]['data']) == 3

    def test_footer_labels(self):
        labels = [inst['footer']['label'] for inst in self.result['sale']]
        assert labels == ['Europe Total', 'Americas Total', 'Asia Total']

    def test_europe_revenue(self):
        assert self.result['sale'][0]['footer']['revenue'] == 756000.0

    def test_americas_revenue(self):
        assert self.result['sale'][1]['footer']['revenue'] == 878000.0


# ─────────────────────────────────────────────────────────────────────────────
# 10: Delivery Note
# ─────────────────────────────────────────────────────────────────────────────

class TestDeliveryNote:
    def setup_method(self):
        self.result, self.lg = run('10_delivery_note')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'supplier', 'order', 'delivery', 'line'}

    def test_supplier_name(self):
        assert self.result['supplier']['name'] == 'Acme Supplies Ltd'

    def test_order_number(self):
        assert self.result['order']['number'] == 'ORD-20260501'

    def test_delivery_date_is_date(self):
        assert isinstance(self.result['delivery']['date'],
                          (datetime.date, datetime.datetime))

    def test_delivery_address(self):
        assert 'Warehouse' in self.result['delivery']['address']

    def test_one_table_instance(self):
        assert len(self.result['line']) == 1

    def test_five_data_rows(self):
        assert len(self.result['line'][0]['data']) == 5

    def test_no_footer(self):
        assert 'footer' not in self.result['line'][0]

    def test_first_item(self):
        row = self.result['line'][0]['data'][0]
        assert row['item'] == 'USB-C Cables'
        assert row['ordered'] == 100
        assert row['delivered'] == 100

    def test_partial_delivery_row(self):
        # Wireless Mice: 50 ordered, 48 delivered
        row = self.result['line'][0]['data'][1]
        assert row['ordered'] == 50
        assert row['delivered'] == 48


# ─────────────────────────────────────────────────────────────────────────────
# 11: Loan Schedule
# ─────────────────────────────────────────────────────────────────────────────

class TestLoanSchedule:
    def setup_method(self):
        self.result, self.lg = run('11_loan_schedule')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'loan', 'amort'}

    def test_loan_amount(self):
        assert self.result['loan']['amount'] == 10000.0

    def test_loan_rate(self):
        assert self.result['loan']['rate'] == '0.5%'

    def test_loan_term(self):
        assert self.result['loan']['term'] == 6

    def test_loan_start_is_date(self):
        assert isinstance(self.result['loan']['start'],
                          (datetime.date, datetime.datetime))

    def test_one_table_instance(self):
        assert len(self.result['amort']) == 1

    def test_six_data_rows(self):
        assert len(self.result['amort'][0]['data']) == 6

    def test_payment_numbers_sequential(self):
        nums = [row['payment_no'] for row in self.result['amort'][0]['data']]
        assert nums == [1, 2, 3, 4, 5, 6]

    def test_footer_totals_payment(self):
        f = self.result['amort'][0]['footer']
        assert f['label'] == 'Totals'
        assert abs(f['payment'] - 10175.2) < 0.01

    def test_footer_totals_principal(self):
        assert self.result['amort'][0]['footer']['principal'] == 10000.0

    def test_final_balance_zero(self):
        last_row = self.result['amort'][0]['data'][-1]
        assert last_row['balance'] == 0.0


# ─────────────────────────────────────────────────────────────────────────────
# 12: Multi-Sheet (extracts from 'Details' sheet, ignores 'Summary' and 'Notes')
# ─────────────────────────────────────────────────────────────────────────────

class TestMultiSheet:
    def setup_method(self):
        self.result, self.lg = run('12_multi_sheet', sheet='Details')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'dept', 'emp'}

    def test_department_name(self):
        assert self.result['dept']['name'] == 'Engineering'

    def test_department_code(self):
        assert self.result['dept']['code'] == 'ENG-001'

    def test_one_table_instance(self):
        assert len(self.result['emp']) == 1

    def test_three_staff_rows(self):
        assert len(self.result['emp'][0]['data']) == 3

    def test_staff_names(self):
        names = [row['name'] for row in self.result['emp'][0]['data']]
        assert names == ['Alice Brown', 'Bob Chen', 'Carol Davis']

    def test_no_footer(self):
        assert 'footer' not in self.result['emp'][0]

    def test_decoy_sheet_not_extracted(self):
        # The 'Summary' sheet has 'Department:' → 'DO NOT EXTRACT'
        # If we accidentally read the wrong sheet we'd get that value
        assert self.result['dept']['name'] != 'DO NOT EXTRACT'

    def test_salary_values(self):
        salaries = [row['salary'] for row in self.result['emp'][0]['data']]
        assert salaries == [85000.0, 72000.0, 55000.0]


# ─────────────────────────────────────────────────────────────────────────────
# 13: HR Attendance
# ─────────────────────────────────────────────────────────────────────────────

class TestHRAttendance:
    def setup_method(self):
        self.result, self.lg = run('13_hr_attendance')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'emp', 'att'}

    def test_employee_name(self):
        assert self.result['emp']['name'] == 'Sarah Connor'

    def test_employee_id(self):
        assert self.result['emp']['id'] == 'EMP-0042'

    def test_employee_year(self):
        assert self.result['emp']['year'] == 2026

    def test_one_table_instance(self):
        assert len(self.result['att']) == 1

    def test_three_month_rows(self):
        assert len(self.result['att'][0]['data']) == 3

    def test_month_names(self):
        names = [row['month'] for row in self.result['att'][0]['data']]
        assert names == ['January', 'February', 'March']

    def test_february_zero_absences(self):
        assert self.result['att'][0]['data'][1]['absent'] == 0

    def test_footer_q1_total(self):
        f = self.result['att'][0]['footer']
        assert f['label'] == 'Q1 Total'
        assert f['working'] == 65
        assert f['absent'] == 3


# ─────────────────────────────────────────────────────────────────────────────
# 14: Named Tables (var: field in HEADER row)
# ─────────────────────────────────────────────────────────────────────────────

class TestNamedTables:
    """
    Verify that a var: field placed in the HEADER row is captured in
    instance['header'], while lbl: fields in the same row are suppressed.

    Fixture layout (3 mini-tables):
      HEADER col A = category name  (var: header.category)
      HEADER col B = 'SKU'          (lbl: col_sku  — the specificity anchor)
      HEADER col C = 'Qty'          (lbl: col_qty)
      HEADER col D = 'Price'        (lbl: col_price)
      DATA   col A = product name   (var: item.name)
      DATA   col B = SKU value      (var: item.sku)
      DATA   col C = quantity       (var: item.qty)
      DATA   col D = price          (var: item.price)
    """
    def setup_method(self):
        self.result, self.lg = run('14_named_tables')
        self.instances = self.result['item']

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_three_instances_found(self):
        assert len(self.instances) == 3

    def test_all_instances_have_header(self):
        for inst in self.instances:
            assert 'header' in inst

    def test_header_has_only_category_key(self):
        # Only the var: field appears; lbl: fields (col_sku etc.) are suppressed
        for inst in self.instances:
            assert set(inst['header'].keys()) == {'category'}

    def test_category_names_in_order(self):
        categories = [inst['header']['category'] for inst in self.instances]
        assert categories == ['Electronics', 'Stationery', 'Furniture']

    def test_lbl_field_values_absent_from_header(self):
        # 'SKU', 'Qty', 'Price' are the lbl: cell values — must not leak into output
        forbidden = {'SKU', 'Qty', 'Price'}
        for inst in self.instances:
            assert not forbidden & set(inst['header'].values())

    def test_data_row_keys(self):
        first_row = self.instances[0]['data'][0]
        assert set(first_row.keys()) == {'name', 'sku', 'qty', 'price'}

    def test_electronics_data_count(self):
        assert len(self.instances[0]['data']) == 3

    def test_stationery_data_count(self):
        assert len(self.instances[1]['data']) == 2

    def test_furniture_data_count(self):
        assert len(self.instances[2]['data']) == 3

    def test_electronics_first_item(self):
        row = self.instances[0]['data'][0]
        assert row['name'] == 'Laptop'
        assert row['sku'] == 'ELC001'
        assert row['qty'] == 5
        assert row['price'] == 999.0

    def test_all_skus_valid(self):
        import re
        for inst in self.instances:
            for row in inst['data']:
                assert re.fullmatch(r'[A-Z]{3}[0-9]{3}', row['sku'])

    def test_quantities_are_integers(self):
        for inst in self.instances:
            for row in inst['data']:
                assert isinstance(row['qty'], int)

    def test_source_present_on_all_instances(self):
        for inst in self.instances:
            assert '_source' in inst

    def test_source_sheet(self):
        for inst in self.instances:
            assert inst['_source']['sheet'] == 'Sheet1'

    def test_source_refs_distinct(self):
        refs = [inst['_source']['ref'] for inst in self.instances]
        assert len(set(refs)) == 3

    def test_first_instance_ref(self):
        assert self.instances[0]['_source']['ref'] == 'A1:D5'

    def test_second_instance_ref(self):
        assert self.instances[1]['_source']['ref'] == 'A6:D9'

    def test_third_instance_ref(self):
        assert self.instances[2]['_source']['ref'] == 'A10:D14'

    def test_source_has_no_name_key(self):
        # The category name lives in instance['header'], not in _source
        for inst in self.instances:
            assert 'name' not in inst['_source']


# ─────────────────────────────────────────────────────────────────────────────
# 15: Annual Budget  (cell:A1 absolute references + two INCOME/EXPENSES tables)
# ─────────────────────────────────────────────────────────────────────────────

MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
          'jul', 'aug', 'sep', 'oct', 'nov', 'dec']


class TestAnnualBudget:
    """
    Fixture layout (sheet 'Budget by month'):
      B2  : 'ANNUAL BUDGET'     (cell:B2 IGNORE)
      B4  : 'SUMMARY'           (cell:B4 IGNORE)
      B5/C5: 'Total monthly income'  / 48440
      B6/C6: 'Total monthly expenses' / 30256.72
      B8/C8: 'BALANCE'          / 18183.28  (IGNORE label)
      B10/C10: 'PERCENTAGE OF INCOME SPENT' / 0.6246  (IGNORE label, percentage type)
      B12:P18 — INCOME table: HEADER:1 section title + HEADER:1 col labels + 4 data rows + footer
      B20:P39 — EXPENSES table: same structure, 17 data rows + footer
      Table group key is 'transaction' (from transaction.* DATA fields).
      Each instance has instance['header']['title'] = 'INCOME' or 'EXPENSES'.
    """
    def setup_method(self):
        self.result, self.lg = run('15_annual_budget')

    # ── no errors or warnings ─────────────────────────────────────────────────

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    # ── top-level structure ───────────────────────────────────────────────────

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'summary', 'transaction'}

    def test_two_table_instances(self):
        assert len(self.result['transaction']) == 2

    # ── summary scalars (extracted via cell:B5, cell:C5, etc.) ───────────────

    def test_summary_income(self):
        assert self.result['summary']['income'] == pytest.approx(48440, rel=1e-4)

    def test_summary_expenses(self):
        assert self.result['summary']['expenses'] == pytest.approx(30256.72, rel=1e-4)

    def test_summary_balance(self):
        assert self.result['summary']['balance'] == pytest.approx(18183.28, rel=1e-4)

    def test_summary_pct_spent(self):
        # 30256.72 / 48440 ≈ 0.6246
        assert abs(self.result['summary']['pct_spent'] - 0.6246) < 0.001

    def test_summary_has_exactly_four_keys(self):
        assert set(self.result['summary'].keys()) == {
            'income', 'expenses', 'balance', 'pct_spent'
        }

    # ── INCOME table (instance 0) ─────────────────────────────────────────────

    def test_income_source_sheet(self):
        assert self.result['transaction'][0]['_source']['sheet'] == 'Budget by month'

    def test_income_source_ref(self):
        # HEADER:1 section title row (B12) is included in the span
        assert self.result['transaction'][0]['_source']['ref'] == 'B12:P18'

    def test_income_header_title(self):
        assert self.result['transaction'][0]['header']['title'] == 'INCOME'

    def test_income_four_data_rows(self):
        assert len(self.result['transaction'][0]['data']) == 4

    def test_income_items(self):
        items = [r['item'] for r in self.result['transaction'][0]['data']]
        assert items == ['Income 1', 'Income 2', 'Income 3', 'Other']

    def test_income_data_row_keys(self):
        expected = {'item'} | set(MONTHS) | {'total', 'avg'}
        assert set(self.result['transaction'][0]['data'][0].keys()) == expected

    def test_income1_jan_value(self):
        assert self.result['transaction'][0]['data'][0]['jan'] == 2500

    def test_income1_total(self):
        assert self.result['transaction'][0]['data'][0]['total'] == 30275

    def test_income_footer_keys(self):
        expected = {'label'} | set(MONTHS) | {'annual', 'avg'}
        assert set(self.result['transaction'][0]['footer'].keys()) == expected

    def test_income_footer_label(self):
        assert self.result['transaction'][0]['footer']['label'] == 'Total'

    def test_income_footer_annual(self):
        assert self.result['transaction'][0]['footer']['annual'] == pytest.approx(48440, rel=1e-4)

    # ── EXPENSES table (instance 1) ───────────────────────────────────────────

    def test_expenses_source_ref(self):
        # HEADER:1 section title row (B20) is included in the span
        assert self.result['transaction'][1]['_source']['ref'] == 'B20:P39'

    def test_expenses_header_title(self):
        assert self.result['transaction'][1]['header']['title'] == 'EXPENSES'

    def test_expenses_seventeen_data_rows(self):
        assert len(self.result['transaction'][1]['data']) == 17

    def test_expenses_first_items(self):
        items = [r['item'] for r in self.result['transaction'][1]['data']][:3]
        assert items == ['Children', 'Debt', 'Dining']

    def test_expenses_footer_annual(self):
        annual = self.result['transaction'][1]['footer']['annual']
        assert abs(annual - 30256.72) < 0.01

    def test_expenses_footer_label(self):
        assert self.result['transaction'][1]['footer']['label'] == 'Total'

    def test_expenses_all_items_have_twelve_months(self):
        for row in self.result['transaction'][1]['data']:
            for m in MONTHS:
                assert m in row

    # ── cross-instance consistency ────────────────────────────────────────────

    def test_sources_are_distinct(self):
        refs = [inst['_source']['ref'] for inst in self.result['transaction']]
        assert len(set(refs)) == 2

    def test_income_before_expenses_by_ref(self):
        def start_row(ref):
            return int(ref.split(':')[0].lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
        rows = [start_row(inst['_source']['ref']) for inst in self.result['transaction']]
        assert rows[0] < rows[1]

    def test_no_lbl_fields_in_output(self):
        forbidden = {'lbl_income', 'lbl_expenses_s', 'col_item', 'col_total_hdr'}
        assert not forbidden & set(self.result.keys())
        assert not forbidden & set(self.result.get('summary', {}).keys())


# ═════════════════════════════════════════════════════════════════════════════
# --all-sheets: process_all() — fixture 06 (two sheets: '2026' and '2025')
# ═════════════════════════════════════════════════════════════════════════════

class TestAllSheets:
    """
    process_all() on fixture 06 (06_merged_cells) which has two sheets:
      '2026' — one SALES CATALOGUE table, total 2369.98
      '2025' — three diagonal tables with doubling totals (3369.98, 6739.96, 13479.92)

    Expected output: {'2026': {...}, '2025': {...}}
    """
    def setup_method(self):
        self.result, self.lg = run_all_sheets('06_merged_cells')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_returns_dict_keyed_by_sheet_name(self):
        assert set(self.result.keys()) == {'2026', '2025'}

    def test_2026_has_one_table_instance(self):
        r2026 = self.result['2026']
        # fixture 06 pattern.xlsx uses 'item' as table group
        table_key = [k for k in r2026 if k not in ('_source',) and isinstance(r2026[k], list)]
        assert len(table_key) == 1
        assert len(r2026[table_key[0]]) == 1

    def test_2025_has_three_table_instances(self):
        r2025 = self.result['2025']
        table_key = [k for k in r2025 if isinstance(r2025[k], list)]
        assert len(table_key) == 1
        assert len(r2025[table_key[0]]) == 3

    def test_2026_source_sheet(self):
        r2026 = self.result['2026']
        table_key = [k for k in r2026 if isinstance(r2026[k], list)][0]
        assert r2026[table_key][0]['_source']['sheet'] == '2026'

    def test_2025_source_sheet(self):
        r2025 = self.result['2025']
        table_key = [k for k in r2025 if isinstance(r2025[k], list)][0]
        for inst in r2025[table_key]:
            assert inst['_source']['sheet'] == '2025'

    def test_2026_grand_total(self):
        r2026 = self.result['2026']
        table_key = [k for k in r2026 if isinstance(r2026[k], list)][0]
        footer = r2026[table_key][0].get('footer', {})
        total = footer.get('total') or footer.get('grand_total')
        assert total == pytest.approx(2369.98, rel=1e-4)

    def test_2025_totals_double(self):
        r2025 = self.result['2025']
        table_key = [k for k in r2025 if isinstance(r2025[k], list)][0]
        instances = r2025[table_key]
        totals = [inst['footer'].get('total') or inst['footer'].get('grand_total')
                  for inst in instances]
        assert totals[0] == pytest.approx(3369.98, rel=1e-4)
        assert totals[1] == pytest.approx(2 * totals[0], rel=1e-4)
        assert totals[2] == pytest.approx(4 * totals[0], rel=1e-4)

    def test_result_matches_per_sheet_run(self):
        # process_all() result for '2026' must match process(sheet='2026')
        r_single, _ = run('06_merged_cells', sheet='2026')
        assert self.result['2026'] == r_single

    def test_result_matches_per_sheet_run_2025(self):
        r_single, _ = run('06_merged_cells', sheet='2025')
        assert self.result['2025'] == r_single


# ─── process_all() on a single-sheet workbook ─────────────────────────────────

class TestAllSheetsSingleSheet:
    """process_all() on fixture 01 (one sheet) → dict with one key."""
    def setup_method(self):
        self.result, self.lg = run_all_sheets('01_simple_invoice')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_returns_single_key(self):
        assert len(self.result) == 1

    def test_key_matches_active_sheet(self):
        r_single, _ = run('01_simple_invoice')
        sheet_name = list(self.result.keys())[0]
        assert self.result[sheet_name] == r_single


# ═════════════════════════════════════════════════════════════════════════════
# cell:next syntax — integration regression across fixtures
# ═════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
# cell:next on fixture 01 (Simple Invoice) — full mirror of TestSimpleInvoice
# ─────────────────────────────────────────────────────────────────────────────

class TestCellNextOnInvoice:
    """cell:next is a transparent alias for cell:1; results must be identical."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['var:', 'inv.number',   'string',   r'[A-Z]{2}[0-9]{6}'],
        ['var:', 'inv.date',     'date',     r'.*'],
        ['var:', 'inv.due_date', 'date',     r'.*'],
        ['var:', 'client.name',  'string',   r'.+'],
        ['var:', 'client.email', 'string',   r'.+@.+'],
        ['var:', 'amount.net',   'currency', r'.*'],
        ['var:', 'amount.vat',   'currency', r'.*'],
        ['var:', 'amount.gross', 'currency', r'.*'],
        ['START:'],
        ['cell:next', 'IGNORE'], ['cell:next', 'inv.number'],
        ['cell:next', 'IGNORE'], ['cell:next', 'inv.date'],
        ['cell:next', 'IGNORE'], ['cell:next', 'inv.due_date'],
        ['cell:next', 'IGNORE'], ['cell:next', 'client.name'],
        ['cell:next', 'IGNORE'], ['cell:next', 'client.email'],
        ['cell:next', 'IGNORE'], ['cell:next', 'amount.net'],
        ['cell:next', 'IGNORE'], ['cell:next', 'amount.vat'],
        ['cell:next', 'IGNORE'], ['cell:next', 'amount.gross'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '01_simple_invoice')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'inv', 'client', 'amount'}

    def test_invoice_number(self):
        assert self.result['inv']['number'] == 'AB123456'

    def test_invoice_date_is_date(self):
        assert isinstance(self.result['inv']['date'], (datetime.date, datetime.datetime))

    def test_due_date_is_date(self):
        assert isinstance(self.result['inv']['due_date'], (datetime.date, datetime.datetime))

    def test_client_name(self):
        assert self.result['client']['name'] == 'Alice Wonderland'

    def test_client_email(self):
        assert self.result['client']['email'] == 'alice@wonderland.example'

    def test_net_amount(self):
        assert self.result['amount']['net'] == 120.0

    def test_vat_amount(self):
        assert self.result['amount']['vat'] == 24.0

    def test_gross_amount(self):
        assert self.result['amount']['gross'] == 144.0


# ─────────────────────────────────────────────────────────────────────────────
# cell:next on fixture 04 (Bank Statement)
# ─────────────────────────────────────────────────────────────────────────────

class TestCellNextOnBankStatement:
    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'acc_lbl', 'string', 'Account Holder:'],
        ['lbl:', 'num_lbl', 'string', 'Account Number:'],
        ['lbl:', 'per_lbl', 'string', 'Statement Period:'],
        ['lbl:', 'col_date',    'string', 'Date'],
        ['lbl:', 'col_desc',    'string', 'Description'],
        ['lbl:', 'col_debit',   'string', 'Debit'],
        ['lbl:', 'col_credit',  'string', 'Credit'],
        ['lbl:', 'col_balance', 'string', 'Balance'],
        ['var:', 'account.holder', 'string',   r'.+'],
        ['var:', 'account.number', 'string',   r'[A-Z0-9]+'],
        ['var:', 'account.period', 'string',   r'.+'],
        ['var:', 'txn.date',        'date',     r'.*'],
        ['var:', 'txn.description', 'string',   r'.+'],
        ['var:', 'txn.debit',       'currency', r'.*'],
        ['var:', 'txn.credit',      'currency', r'.*'],
        ['var:', 'txn.balance',     'currency', r'.*'],
        ['var:', 'totals.label',   'string',   'Totals'],
        ['var:', 'totals.debits',  'currency', r'.*'],
        ['var:', 'totals.credits', 'currency', r'.*'],
        ['START:'],
        ['cell:next', 'acc_lbl'], ['cell:next', 'account.holder'],
        ['cell:next', 'num_lbl'], ['cell:next', 'account.number'],
        ['cell:next', 'per_lbl'], ['cell:next', 'account.period'],
        ['table:*'],
        [None, 'HEADER:1', 'col_date', 'col_desc', 'col_debit', 'col_credit', 'col_balance'],
        [None, 'DATA:*',   'txn.date', 'txn.description', 'txn.debit', 'txn.credit', 'txn.balance'],
        [None, 'FOOTER:1', 'totals.label', 'IGNORE', 'totals.debits', 'totals.credits', 'IGNORE'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '04_bank_statement')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_account_holder(self):
        assert self.result['account']['holder'] == 'Jane Smith'

    def test_account_number(self):
        assert self.result['account']['number'] == 'GB29NWBK60161331926819'

    def test_account_period(self):
        assert self.result['account']['period'] == 'January 2026'

    def test_five_transactions(self):
        assert len(self.result['txn'][0]['data']) == 5

    def test_footer_debits(self):
        assert self.result['txn'][0]['footer']['debits'] == 455.5

    def test_footer_credits(self):
        assert self.result['txn'][0]['footer']['credits'] == 5700.0


# ─────────────────────────────────────────────────────────────────────────────
# cell:next on fixture 05 (Expense Report — 4 header cell pairs)
# ─────────────────────────────────────────────────────────────────────────────

class TestCellNextOnExpenseReport:
    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'emp_lbl',  'string', 'Employee:'],
        ['lbl:', 'dept_lbl', 'string', 'Department:'],
        ['lbl:', 'per_lbl',  'string', 'Period:'],
        ['lbl:', 'mgr_lbl',  'string', 'Manager:'],
        ['lbl:', 'col_date', 'string', 'Date'],
        ['lbl:', 'col_desc', 'string', 'Description'],
        ['lbl:', 'col_amt',  'string', 'Amount'],
        ['lbl:', 'col_rcpt', 'string', 'Receipt No'],
        ['var:', 'emp.name',       'string',   r'.+'],
        ['var:', 'emp.department', 'string',   r'.+'],
        ['var:', 'emp.period',     'string',   r'.+'],
        ['var:', 'emp.manager',    'string',   r'.+'],
        ['var:', 'exp.date',        'date',     r'.*'],
        ['var:', 'exp.description', 'string',   r'.+'],
        ['var:', 'exp.amount',      'currency', r'.*'],
        ['var:', 'exp.receipt',     'string',   r'REC[0-9]+'],
        ['var:', 'sub.label',  'string',   'Subtotal'],
        ['var:', 'sub.amount', 'currency', r'.*'],
        ['START:'],
        ['cell:next', 'emp_lbl'],  ['cell:next', 'emp.name'],
        ['cell:next', 'dept_lbl'], ['cell:next', 'emp.department'],
        ['cell:next', 'per_lbl'],  ['cell:next', 'emp.period'],
        ['cell:next', 'mgr_lbl'],  ['cell:next', 'emp.manager'],
        ['table:*'],
        [None, 'HEADER:1', 'col_date', 'col_desc', 'col_amt', 'col_rcpt'],
        [None, 'DATA:*',   'exp.date', 'exp.description', 'exp.amount', 'exp.receipt'],
        [None, 'FOOTER:1', 'IGNORE',   'sub.label', 'sub.amount', 'IGNORE'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '05_expense_report')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_employee_name(self):
        assert self.result['emp']['name'] == 'John Smith'

    def test_employee_department(self):
        assert self.result['emp']['department'] == 'Engineering'

    def test_employee_period(self):
        assert self.result['emp']['period'] == 'Q1 2026'

    def test_three_table_instances(self):
        assert len(self.result['exp']) == 3

    def test_travel_subtotal(self):
        assert self.result['exp'][0]['footer']['amount'] == 450.0

    def test_accommodation_row_count(self):
        assert len(self.result['exp'][2]['data']) == 2


# ─────────────────────────────────────────────────────────────────────────────
# cell:next on fixture 11 (Loan Schedule — 4 header cell pairs, cross-row)
# ─────────────────────────────────────────────────────────────────────────────

class TestCellNextOnLoanSchedule:
    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'amt_lbl',  'string', 'Loan Amount:'],
        ['lbl:', 'rate_lbl', 'string', 'Monthly Rate:'],
        ['lbl:', 'term_lbl', 'string', 'Term (months):'],
        ['lbl:', 'start_lbl','string', 'Start Date:'],
        ['lbl:', 'col_no',        'string', 'Payment #'],
        ['lbl:', 'col_date',      'string', 'Date'],
        ['lbl:', 'col_payment',   'string', 'Payment'],
        ['lbl:', 'col_principal', 'string', 'Principal'],
        ['lbl:', 'col_interest',  'string', 'Interest'],
        ['lbl:', 'col_balance',   'string', 'Balance'],
        ['var:', 'loan.amount',  'currency', r'.*'],
        ['var:', 'loan.rate',    'string',   r'[0-9.]+%'],
        ['var:', 'loan.term',    'integer',  r'[0-9]+'],
        ['var:', 'loan.start',   'date',     r'.*'],
        ['var:', 'amort.payment_no', 'integer',  r'[0-9]+'],
        ['var:', 'amort.date',       'date',     r'.*'],
        ['var:', 'amort.payment',    'currency', r'.*'],
        ['var:', 'amort.principal',  'currency', r'.*'],
        ['var:', 'amort.interest',   'currency', r'.*'],
        ['var:', 'amort.balance',    'currency', r'.*'],
        ['var:', 'summary.label',     'string',   'Totals'],
        ['var:', 'summary.payment',   'currency', r'.*'],
        ['var:', 'summary.principal', 'currency', r'.*'],
        ['var:', 'summary.interest',  'currency', r'.*'],
        ['START:'],
        ['cell:next', 'amt_lbl'],   ['cell:next', 'loan.amount'],
        ['cell:next', 'rate_lbl'],  ['cell:next', 'loan.rate'],
        ['cell:next', 'term_lbl'],  ['cell:next', 'loan.term'],
        ['cell:next', 'start_lbl'], ['cell:next', 'loan.start'],
        ['table:1'],
        [None, 'HEADER:1', 'col_no', 'col_date', 'col_payment', 'col_principal', 'col_interest', 'col_balance'],
        [None, 'DATA:*',   'amort.payment_no', 'amort.date', 'amort.payment', 'amort.principal', 'amort.interest', 'amort.balance'],
        [None, 'FOOTER:1', 'summary.label', 'IGNORE', 'summary.payment', 'summary.principal', 'summary.interest', 'IGNORE'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '11_loan_schedule')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_loan_amount(self):
        assert self.result['loan']['amount'] == 10000.0

    def test_loan_rate(self):
        assert self.result['loan']['rate'] == '0.5%'

    def test_loan_term(self):
        assert self.result['loan']['term'] == 6

    def test_six_amort_rows(self):
        assert len(self.result['amort'][0]['data']) == 6

    def test_footer_payment(self):
        assert abs(self.result['amort'][0]['footer']['payment'] - 10175.2) < 0.01


# ═════════════════════════════════════════════════════════════════════════════
# cell:A1 absolute references — spot-checks on known coordinates
# ═════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
# Absolute refs on fixture 01: cells scattered across rows 1, 3, 5
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout (from generate_fixtures.py):
#   A1='Invoice No:' B1='AB123456'  C1='Date:'     D1=datetime  E1='Due Date:' F1=datetime
#   A3='Client:'     B3='Alice..'   D3='Email:'     E3='alice@..'
#   A5='Net:'        B5=120.0       C5='VAT:'       D5=24.0      E5='Total:'   F5=144.0

class TestAbsoluteRefsInvoice:
    """cell:A1 absolute addressing jumps directly to non-sequential cells."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['var:', 'inv.number',   'string',   r'[A-Z]{2}[0-9]{6}'],
        ['var:', 'inv.date',     'date',     r'.*'],
        ['var:', 'inv.due_date', 'date',     r'.*'],
        ['var:', 'client.name',  'string',   r'.+'],
        ['var:', 'client.email', 'string',   r'.+@.+'],
        ['var:', 'amount.net',   'currency', r'.*'],
        ['var:', 'amount.vat',   'currency', r'.*'],
        ['var:', 'amount.gross', 'currency', r'.*'],
        ['START:'],
        ['cell:A1', 'IGNORE'],       # 'Invoice No:'
        ['cell:B1', 'inv.number'],   # 'AB123456'
        ['cell:C1', 'IGNORE'],       # 'Date:'
        ['cell:D1', 'inv.date'],     # datetime(2026,3,15)
        ['cell:E1', 'IGNORE'],       # 'Due Date:'
        ['cell:F1', 'inv.due_date'], # datetime(2026,4,14)
        ['cell:A3', 'IGNORE'],       # 'Client:'
        ['cell:B3', 'client.name'],  # 'Alice Wonderland'
        ['cell:D3', 'IGNORE'],       # 'Email:'
        ['cell:E3', 'client.email'], # 'alice@wonderland.example'
        ['cell:A5', 'IGNORE'],       # 'Net:'
        ['cell:B5', 'amount.net'],   # 120.0
        ['cell:C5', 'IGNORE'],       # 'VAT:'
        ['cell:D5', 'amount.vat'],   # 24.0
        ['cell:E5', 'IGNORE'],       # 'Total:'
        ['cell:F5', 'amount.gross'], # 144.0
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '01_simple_invoice')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_keys(self):
        assert set(self.result.keys()) == {'inv', 'client', 'amount'}

    def test_invoice_number(self):
        assert self.result['inv']['number'] == 'AB123456'

    def test_inv_date_is_date(self):
        assert isinstance(self.result['inv']['date'], (datetime.date, datetime.datetime))

    def test_client_name(self):
        assert self.result['client']['name'] == 'Alice Wonderland'

    def test_client_email(self):
        assert self.result['client']['email'] == 'alice@wonderland.example'

    def test_net_amount(self):
        assert self.result['amount']['net'] == 120.0

    def test_vat_amount(self):
        assert self.result['amount']['vat'] == 24.0

    def test_gross_amount(self):
        assert self.result['amount']['gross'] == 144.0


# ─────────────────────────────────────────────────────────────────────────────
# Absolute refs on fixture 07 (Timesheet) — header cells + table
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout:
#   A1='Employee:'      B1='Bob Martin'
#   A2='Week Starting:' B2=datetime(2026,5,11)
#   A3='Project:'       B3='grepxcel v2'
#   A5='Day' B5='Hours' C5='Task'   ← TABLE HEADER
#   Rows 6–12: 7 day rows
#   A13='Total Hours'   B13=37.0   ← FOOTER

class TestAbsoluteRefsTimesheet:
    """Absolute refs extract header scalars; subsequent table:1 still works."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['lbl:', 'emp_lbl',  'string', 'Employee:'],
        ['lbl:', 'week_lbl', 'string', 'Week Starting:'],
        ['lbl:', 'proj_lbl', 'string', 'Project:'],
        ['lbl:', 'col_day',   'string', 'Day'],
        ['lbl:', 'col_hours', 'string', 'Hours'],
        ['lbl:', 'col_task',  'string', 'Task'],
        ['var:', 'emp.name',       'string',   r'.+'],
        ['var:', 'emp.week_start', 'date',     r'.*'],
        ['var:', 'emp.project',    'string',   r'.+'],
        ['var:', 'day.name',   'string',   r'.+'],
        ['var:', 'day.hours',  'currency', r'.*'],
        ['var:', 'day.task',   'string',   r'.*'],
        ['var:', 'total.label', 'string',   'Total Hours'],
        ['var:', 'total.hours', 'currency', r'.*'],
        ['START:'],
        ['cell:A1', 'emp_lbl'],        # validates 'Employee:'
        ['cell:B1', 'emp.name'],       # 'Bob Martin'
        ['cell:A2', 'week_lbl'],       # validates 'Week Starting:'
        ['cell:B2', 'emp.week_start'], # datetime(2026,5,11)
        ['cell:A3', 'proj_lbl'],       # validates 'Project:'
        ['cell:B3', 'emp.project'],    # 'grepxcel v2'
        ['table:1'],
        [None, 'HEADER:1', 'col_day', 'col_hours', 'col_task'],
        [None, 'DATA:*',   'day.name', 'day.hours', 'day.task'],
        [None, 'FOOTER:1', 'total.label', 'total.hours', 'IGNORE'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '07_timesheet')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_employee_name(self):
        assert self.result['emp']['name'] == 'Bob Martin'

    def test_week_start_is_date(self):
        assert isinstance(self.result['emp']['week_start'], (datetime.date, datetime.datetime))

    def test_project(self):
        assert self.result['emp']['project'] == 'grepxcel v2'

    def test_lbl_fields_absent(self):
        # lbl: fields must not appear in output
        for key in ('emp_lbl', 'week_lbl', 'proj_lbl'):
            assert key not in self.result

    def test_seven_day_rows(self):
        assert len(self.result['day'][0]['data']) == 7

    def test_total_hours(self):
        assert self.result['day'][0]['footer']['hours'] == 37.0

    def test_table_present(self):
        assert 'day' in self.result


# ─────────────────────────────────────────────────────────────────────────────
# Absolute refs on fixture 10 (Delivery Note) — 2-row header, 4 fields
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout:
#   A1='Supplier:'      B1='Acme Supplies Ltd'  C1='Order No:'  D1='ORD-20260501'
#   A2='Delivery Date:' B2=datetime(2026,5,8)   C2='Deliver To:' D2='Warehouse B, Unit 12'
#   A4='Item' B4='Ordered' C4='Delivered' D4='Unit' E4='Notes'  ← TABLE HEADER
#   Rows 5–9: 5 delivery line items (no footer)

class TestAbsoluteRefsDeliveryNote:
    """Absolute refs address a compact 2-row header with 4 label/value pairs."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['lbl:', 'sup_lbl',  'string', 'Supplier:'],
        ['lbl:', 'ord_lbl',  'string', 'Order No:'],
        ['lbl:', 'del_lbl',  'string', 'Delivery Date:'],
        ['lbl:', 'addr_lbl', 'string', 'Deliver To:'],
        ['lbl:', 'col_item',      'string', 'Item'],
        ['lbl:', 'col_ordered',   'string', 'Ordered'],
        ['lbl:', 'col_delivered', 'string', 'Delivered'],
        ['lbl:', 'col_unit',      'string', 'Unit'],
        ['lbl:', 'col_notes',     'string', 'Notes'],
        ['var:', 'supplier.name',    'string', r'.+'],
        ['var:', 'order.number',     'string', r'ORD-[0-9]+'],
        ['var:', 'delivery.date',    'date',   r'.*'],
        ['var:', 'delivery.address', 'string', r'.+'],
        ['var:', 'line.item',      'string',  r'.+'],
        ['var:', 'line.ordered',   'integer', r'[0-9]+'],
        ['var:', 'line.delivered', 'integer', r'[0-9]+'],
        ['var:', 'line.unit',      'string',  r'.+'],
        ['var:', 'line.notes',     'string',  r'.*'],
        ['START:'],
        ['cell:A1', 'sup_lbl'],          # 'Supplier:'
        ['cell:B1', 'supplier.name'],    # 'Acme Supplies Ltd'
        ['cell:C1', 'ord_lbl'],          # 'Order No:'
        ['cell:D1', 'order.number'],     # 'ORD-20260501'
        ['cell:A2', 'del_lbl'],          # 'Delivery Date:'
        ['cell:B2', 'delivery.date'],    # datetime(2026,5,8)
        ['cell:C2', 'addr_lbl'],         # 'Deliver To:'
        ['cell:D2', 'delivery.address'], # 'Warehouse B, Unit 12'
        ['table:1'],
        [None, 'HEADER:1', 'col_item', 'col_ordered', 'col_delivered', 'col_unit', 'col_notes'],
        [None, 'DATA:*',   'line.item', 'line.ordered', 'line.delivered', 'line.unit', 'line.notes'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '10_delivery_note')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_supplier_name(self):
        assert self.result['supplier']['name'] == 'Acme Supplies Ltd'

    def test_order_number(self):
        assert self.result['order']['number'] == 'ORD-20260501'

    def test_delivery_date_is_date(self):
        assert isinstance(self.result['delivery']['date'], (datetime.date, datetime.datetime))

    def test_delivery_address(self):
        assert 'Warehouse' in self.result['delivery']['address']

    def test_five_line_items(self):
        assert len(self.result['line'][0]['data']) == 5

    def test_first_item_fully_delivered(self):
        row = self.result['line'][0]['data'][0]
        assert row['item'] == 'USB-C Cables'
        assert row['ordered'] == 100
        assert row['delivered'] == 100

    def test_partial_delivery_item(self):
        # Wireless Mice: 50 ordered, 48 delivered
        row = self.result['line'][0]['data'][1]
        assert row['ordered'] == 50
        assert row['delivered'] == 48


# ─────────────────────────────────────────────────────────────────────────────
# Absolute refs on fixture 11 (Loan Schedule) — 2-row, 4 label/value pairs
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout:
#   A1='Loan Amount:'   B1=10000.0  C1='Monthly Rate:'  D1='0.5%'
#   A2='Term (months):' B2=6        C2='Start Date:'    D2=datetime(2026,2,1)
#   A4='Payment #' … ← TABLE HEADER

class TestAbsoluteRefsLoanSchedule:
    """Absolute refs on a 2-row, 4-pair header before an amortisation table."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'amt_lbl',  'string', 'Loan Amount:'],
        ['lbl:', 'rate_lbl', 'string', 'Monthly Rate:'],
        ['lbl:', 'term_lbl', 'string', 'Term (months):'],
        ['lbl:', 'start_lbl','string', 'Start Date:'],
        ['lbl:', 'col_no',        'string', 'Payment #'],
        ['lbl:', 'col_date',      'string', 'Date'],
        ['lbl:', 'col_payment',   'string', 'Payment'],
        ['lbl:', 'col_principal', 'string', 'Principal'],
        ['lbl:', 'col_interest',  'string', 'Interest'],
        ['lbl:', 'col_balance',   'string', 'Balance'],
        ['var:', 'loan.amount',  'currency', r'.*'],
        ['var:', 'loan.rate',    'string',   r'[0-9.]+%'],
        ['var:', 'loan.term',    'integer',  r'[0-9]+'],
        ['var:', 'loan.start',   'date',     r'.*'],
        ['var:', 'amort.payment_no', 'integer',  r'[0-9]+'],
        ['var:', 'amort.date',       'date',     r'.*'],
        ['var:', 'amort.payment',    'currency', r'.*'],
        ['var:', 'amort.principal',  'currency', r'.*'],
        ['var:', 'amort.interest',   'currency', r'.*'],
        ['var:', 'amort.balance',    'currency', r'.*'],
        ['var:', 'summary.label',     'string',   'Totals'],
        ['var:', 'summary.payment',   'currency', r'.*'],
        ['var:', 'summary.principal', 'currency', r'.*'],
        ['var:', 'summary.interest',  'currency', r'.*'],
        ['START:'],
        ['cell:A1', 'amt_lbl'],    # 'Loan Amount:'
        ['cell:B1', 'loan.amount'],# 10000.0
        ['cell:C1', 'rate_lbl'],   # 'Monthly Rate:'
        ['cell:D1', 'loan.rate'],  # '0.5%'
        ['cell:A2', 'term_lbl'],   # 'Term (months):'
        ['cell:B2', 'loan.term'],  # 6
        ['cell:C2', 'start_lbl'],  # 'Start Date:'
        ['cell:D2', 'loan.start'], # datetime(2026,2,1)
        ['table:1'],
        [None, 'HEADER:1', 'col_no', 'col_date', 'col_payment', 'col_principal', 'col_interest', 'col_balance'],
        [None, 'DATA:*',   'amort.payment_no', 'amort.date', 'amort.payment', 'amort.principal', 'amort.interest', 'amort.balance'],
        [None, 'FOOTER:1', 'summary.label', 'IGNORE', 'summary.payment', 'summary.principal', 'summary.interest', 'IGNORE'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '11_loan_schedule')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_loan_amount(self):
        assert self.result['loan']['amount'] == 10000.0

    def test_loan_rate(self):
        assert self.result['loan']['rate'] == '0.5%'

    def test_loan_term(self):
        assert self.result['loan']['term'] == 6

    def test_loan_start_is_date(self):
        assert isinstance(self.result['loan']['start'], (datetime.date, datetime.datetime))

    def test_six_amort_rows(self):
        assert len(self.result['amort'][0]['data']) == 6

    def test_payment_numbers_sequential(self):
        nums = [r['payment_no'] for r in self.result['amort'][0]['data']]
        assert nums == [1, 2, 3, 4, 5, 6]

    def test_footer_payment(self):
        assert abs(self.result['amort'][0]['footer']['payment'] - 10175.2) < 0.01

    def test_footer_principal(self):
        assert self.result['amort'][0]['footer']['principal'] == 10000.0


# ─────────────────────────────────────────────────────────────────────────────
# Absolute refs on fixture 13 (HR Attendance) — header spans 2 rows, 3 pairs
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout:
#   A1='Employee:'    B1='Sarah Connor'  C1='Employee ID:'  D1='EMP-0042'
#   A2='Year:'        B2=2026
#   A4='Month' … ← TABLE HEADER

class TestAbsoluteRefsHRAttendance:
    """Absolute refs span two rows: (A1,B1,C1,D1) then (A2,B2)."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['lbl:', 'emp_lbl',  'string', 'Employee:'],
        ['lbl:', 'id_lbl',   'string', 'Employee ID:'],
        ['lbl:', 'year_lbl', 'string', 'Year:'],
        ['lbl:', 'col_month',   'string', 'Month'],
        ['lbl:', 'col_working', 'string', 'Working Days'],
        ['lbl:', 'col_absent',  'string', 'Days Absent'],
        ['lbl:', 'col_reason',  'string', 'Reason'],
        ['var:', 'emp.name', 'string',  r'.+'],
        ['var:', 'emp.id',   'string',  r'EMP-[0-9]{4}'],
        ['var:', 'emp.year', 'integer', r'[0-9]{4}'],
        ['var:', 'att.month',   'string',  r'.+'],
        ['var:', 'att.working', 'integer', r'[0-9]+'],
        ['var:', 'att.absent',  'integer', r'[0-9]+'],
        ['var:', 'att.reason',  'string',  r'.*'],
        ['var:', 'totals.label',   'string',  'Q1 Total'],
        ['var:', 'totals.working', 'integer', r'[0-9]+'],
        ['var:', 'totals.absent',  'integer', r'[0-9]+'],
        ['START:'],
        ['cell:A1', 'emp_lbl'],  # 'Employee:'
        ['cell:B1', 'emp.name'], # 'Sarah Connor'
        ['cell:C1', 'id_lbl'],   # 'Employee ID:'
        ['cell:D1', 'emp.id'],   # 'EMP-0042'
        ['cell:A2', 'year_lbl'], # 'Year:'
        ['cell:B2', 'emp.year'], # 2026
        ['table:1'],
        [None, 'HEADER:1', 'col_month', 'col_working', 'col_absent', 'col_reason'],
        [None, 'DATA:*',   'att.month', 'att.working', 'att.absent', 'att.reason'],
        [None, 'FOOTER:1', 'totals.label', 'totals.working', 'totals.absent', 'IGNORE'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '13_hr_attendance')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_employee_name(self):
        assert self.result['emp']['name'] == 'Sarah Connor'

    def test_employee_id(self):
        assert self.result['emp']['id'] == 'EMP-0042'

    def test_employee_year(self):
        assert self.result['emp']['year'] == 2026

    def test_lbl_fields_absent(self):
        for key in ('emp_lbl', 'id_lbl', 'year_lbl'):
            assert key not in self.result

    def test_three_month_rows(self):
        assert len(self.result['att'][0]['data']) == 3

    def test_footer_q1_working(self):
        assert self.result['att'][0]['footer']['working'] == 65

    def test_footer_q1_absent(self):
        assert self.result['att'][0]['footer']['absent'] == 3


# ═════════════════════════════════════════════════════════════════════════════
# Mixed cell:next + cell:A1 in the same pattern
# ═════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
# Mixed syntax on fixture 03 (Purchase Order)
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout (all row 1):
#   A1='PO Number:' B1='PO-2026'  C1='Date:'   D1=datetime(2026,5,1)
#   E1='Vendor:'    F1='Acme Supplies'  G1='Ref:'  H1='ACME001'
#
# Strategy: cell:A1 absolute for PO label (A1), then cell:next for value (B1),
# then cell:C1 absolute for Date label (skips B1 which is consumed), then
# cell:next for the rest sequentially.

class TestMixedSyntaxPurchaseOrder:
    """Mix cell:next and cell:A1 in a single START sequence."""

    _PAT = [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'po_label',     'string',   'PO Number:'],
        ['lbl:', 'date_label',   'string',   'Date:'],
        ['lbl:', 'vendor_label', 'string',   'Vendor:'],
        ['lbl:', 'ref_label',    'string',   'Ref:'],
        ['lbl:', 'col_item',     'string',   'Item'],
        ['lbl:', 'col_desc',     'string',   'Description'],
        ['lbl:', 'col_qty',      'string',   'Qty'],
        ['lbl:', 'col_price',    'string',   'Unit Price'],
        ['lbl:', 'col_total',    'string',   'Total'],
        ['var:', 'po.number',    'string',   r'PO-[0-9]{4}'],
        ['var:', 'po.date',      'date',     r'.*'],
        ['var:', 'vendor.name',  'string',   r'.+'],
        ['var:', 'vendor.ref',   'string',   r'[A-Z0-9]+'],
        ['var:', 'row.item',     'string',   r'.{3,50}'],
        ['var:', 'row.desc',     'string',   r'.*'],
        ['var:', 'row.qty',      'integer',  r'[1-9][0-9]*'],
        ['var:', 'row.price',    'currency', r'.*'],
        ['var:', 'row.total',    'currency', r'.*'],
        ['var:', 'footer.label', 'string',   'Grand Total'],
        ['var:', 'footer.value', 'currency', r'.*'],
        ['START:'],
        # Absolute ref for first label, then cell:next for its value
        ['cell:A1',   'po_label'],    # 'PO Number:' (absolute)
        ['cell:next', 'po.number'],   # B1 = 'PO-2026' (next non-consumed)
        # Absolute ref jumps to 'Date:' at C1, then next picks up D1
        ['cell:C1',   'date_label'],  # 'Date:' (absolute, B1 already consumed)
        ['cell:next', 'po.date'],     # D1 = datetime (next non-consumed)
        # Remaining pairs via cell:next
        ['cell:next', 'vendor_label'],# E1 = 'Vendor:'
        ['cell:next', 'vendor.name'], # F1 = 'Acme Supplies'
        ['cell:next', 'ref_label'],   # G1 = 'Ref:'
        ['cell:next', 'vendor.ref'],  # H1 = 'ACME001'
        ['table:*'],
        [None, 'HEADER:1', 'col_item',     'col_desc', 'col_qty', 'col_price', 'col_total'],
        [None, 'DATA:*',   'row.item',     'row.desc', 'row.qty', 'row.price', 'row.total'],
        [None, 'FOOTER:1', 'footer.label', 'IGNORE',   'IGNORE',  'IGNORE',    'footer.value'],
        ['END:'],
    ]

    def setup_method(self):
        self.result, self.lg = run_custom_pattern(self._PAT, '03_purchase_order')

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_exactly_one_warning(self):
        # 'X' in row.item fails the regex r'.{3,50}' — same as original fixture
        assert len(self.lg.issues()) == 1

    def test_po_number(self):
        assert self.result['po']['number'] == 'PO-2026'

    def test_po_date_is_date(self):
        assert isinstance(self.result['po']['date'], (datetime.date, datetime.datetime))

    def test_vendor_name(self):
        assert self.result['vendor']['name'] == 'Acme Supplies'

    def test_vendor_ref(self):
        assert self.result['vendor']['ref'] == 'ACME001'

    def test_one_table_instance(self):
        assert len(self.result['row']) == 1

    def test_three_data_rows(self):
        assert len(self.result['row'][0]['data']) == 3

    def test_footer_grand_total(self):
        assert self.result['row'][0]['footer']['value'] == 3030.0


class TestSheetSelectionByNameOrIndex:
    """Regression: a numerically-named sheet must be selectable by NAME.

    The CLI used to convert '2025' to the integer index 2025 (out of range),
    so `--sheet 2025` could never reach a sheet literally named '2025'. Sheet
    resolution now matches a name first and only falls back to a 0-based index
    when no sheet has that name. 06_merged_cells has sheets '2026' and '2025'.
    """

    def test_cli_resolve_sheet_keeps_numeric_name_as_string(self):
        import argparse
        from grepxcel.cli import _resolve_sheet
        assert _resolve_sheet(argparse.Namespace(sheet='2025')) == '2025'  # not int
        assert _resolve_sheet(argparse.Namespace(sheet='Details')) == 'Details'
        assert _resolve_sheet(argparse.Namespace(sheet=None)) is None

    def test_numeric_sheet_name_selects_by_name(self):
        r2025, lg25 = run('06_merged_cells', sheet='2025')
        r2026, lg26 = run('06_merged_cells', sheet='2026')
        assert not lg25.has_errors()
        assert not lg26.has_errors()
        # Name resolution picked different sheets → different extracted data.
        assert r2025 != r2026

    def test_numeric_string_falls_back_to_index_when_no_such_name(self):
        # No sheet is named '0', so '0' is resolved as the first sheet (index 0),
        # matching the integer-index behaviour exactly.
        by_str, lg = run('06_merged_cells', sheet='0')
        by_int, _  = run('06_merged_cells', sheet=0)
        assert not lg.has_errors()
        assert by_str == by_int

    def test_out_of_range_numeric_string_errors(self):
        _, lg = run('06_merged_cells', sheet='99')
        assert lg.has_errors()


# ─────────────────────────────────────────────────────────────────────────────
# ignore.case config — end-to-end case-insensitive matching
# ─────────────────────────────────────────────────────────────────────────────

def _write_data(rows: list) -> str:
    """Write data rows to a temp xlsx; caller must os.unlink() it."""
    wb = _openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    fd, path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    wb.save(path)
    return path


def _run(pattern_rows: list, data_rows: list):
    pat = _write_pattern(pattern_rows)
    data = _write_data(data_rows)
    lg = Logger(level=VerbosityLevel.QUIET)
    result = Engine().process(pattern_file=pat, data_file=data, logger=lg)
    os.unlink(pat)
    os.unlink(data)
    return result, lg


class TestIgnoreCaseEndToEnd:
    _DATA = [['PAID']]  # value differs in case from the pattern regex 'paid'

    def test_case_sensitive_by_default_warns(self):
        pat = [
            ['var:', 'status', 'string', r'paid'],
            ['START:'], ['cell:next', 'status'], ['END:'],
        ]
        result, lg = _run(pat, self._DATA)
        # value is still extracted, but it fails its pattern → a warning
        assert result['status'] == 'PAID'
        assert lg.has_warnings()

    def test_ignore_case_yes_no_warning(self):
        pat = [
            ['config:', 'ignore.case', 'yes'],
            ['var:', 'status', 'string', r'paid'],
            ['START:'], ['cell:next', 'status'], ['END:'],
        ]
        result, lg = _run(pat, self._DATA)
        assert result['status'] == 'PAID'
        assert not lg.has_warnings()


# ═════════════════════════════════════════════════════════════════════════════
# Verbose per-field extraction trace (-v): 'field ← cell = value ✓/✗'
# ═════════════════════════════════════════════════════════════════════════════

class TestVerboseExtractionTrace:
    """At VERBOSE, the engine emits a per-field trace for scalar cells and table
    DATA fields so users can see what was extracted from where and whether it
    passed validation — the key aid for debugging non-matching patterns."""

    def _verbose_run(self, fixture_name: str):
        folder = os.path.join(FIXTURES, fixture_name)
        lg = Logger(level=VerbosityLevel.VERBOSE)
        Engine().process(
            pattern_file=find_pattern_xlsx(folder),
            data_file=os.path.join(folder, 'data.xlsx'),
            logger=lg,
        )

    def test_scalar_and_table_field_traces(self, capsys):
        # fixture 03 has scalar cells + a table whose 'X' item fails its regex.
        self._verbose_run('03_purchase_order')
        err = capsys.readouterr().err
        assert 'po.number' in err and '←' in err          # scalar cell trace
        assert 'row.item' in err and '✓' in err           # table DATA field, passing
        assert '✗' in err and 'does not match' in err     # the failing 'X' item

    def test_clean_fixture_has_no_fail_marks(self, capsys):
        # fixture 01 is all-valid → every field passes, no ✗ in the trace.
        self._verbose_run('01_simple_invoice')
        err = capsys.readouterr().err
        assert '✓' in err
        assert '✗' not in err

    def test_quiet_emits_no_trace(self, capsys):
        folder = os.path.join(FIXTURES, '01_simple_invoice')
        lg = Logger(level=VerbosityLevel.QUIET)
        Engine().process(
            pattern_file=find_pattern_xlsx(folder),
            data_file=os.path.join(folder, 'data.xlsx'),
            logger=lg,
        )
        assert '←' not in capsys.readouterr().err


# ═════════════════════════════════════════════════════════════════════════════
# seek: cursor repositioning
# ═════════════════════════════════════════════════════════════════════════════

class TestSeekEngine:
    """End-to-end tests for seek: instruction (cursor repositioning without reading)."""

    def test_seek_to_earlier_cell_enables_backward_read(self):
        """cell:C1 reads C1 (cursor → D1); seek:A1 repositions; cell:next reads A1."""
        pat = [
            ['var:', 'last',  'string', '.*'],
            ['var:', 'first', 'string', '.*'],
            ['START:'],
            ['cell:C1', 'last'],    # read C1; cursor advances past C1
            ['seek:A1'],            # reposition cursor to A1
            ['cell:next', 'first'], # A1 not consumed → reads A1
            ['END:'],
        ]
        data = [['a_value', None, 'c_value']]  # A1="a_value", B1=None, C1="c_value"
        result, lg = _run(pat, data)
        assert result.get('last') == 'c_value'
        assert result.get('first') == 'a_value'
        assert not lg.has_errors()

    def test_seek_skips_already_consumed_cells(self):
        """seek back to a consumed cell; cell:next skips it and reads the next one."""
        pat = [
            ['var:', 'x', 'string', '.*'],
            ['var:', 'y', 'string', '.*'],
            ['START:'],
            ['cell:A1', 'x'],   # reads and consumes A1
            ['seek:A1'],        # reposition cursor back to A1
            ['cell:next', 'y'], # A1 consumed → reads A2
            ['END:'],
        ]
        data = [['A1_value'], ['A2_value']]
        result, lg = _run(pat, data)
        assert result.get('x') == 'A1_value'
        assert result.get('y') == 'A2_value'
        assert not lg.has_errors()

    def test_seek_forward_skips_cells_before_target(self):
        """seek:C1 with cursor at A1 skips A1 and B1; cell:next reads C1."""
        pat = [
            ['var:', 'x', 'string', '.*'],
            ['START:'],
            ['seek:C1'],        # jump forward: cursor set to C1 position
            ['cell:next', 'x'], # reads C1 (first non-empty at or after C1)
            ['END:'],
        ]
        data = [['skip_a', 'skip_b', 'read_c']]
        result, lg = _run(pat, data)
        assert result.get('x') == 'read_c'
        assert not lg.has_errors()


# ─────────────────────────────────────────────────────────────────────────────
# 20: Sales Report  (Quelldaten sheet; pattern-manual.xlsx targets pivot cache
#     that openpyxl cannot read without Excel refresh — use draft pattern here)
# ─────────────────────────────────────────────────────────────────────────────

class TestSalesReport:
    def setup_method(self):
        self.result, self.lg = run_with_pattern_file(
            '20_sales_report', 'pattern-from-draft.csv', sheet='Quelldaten'
        )

    def test_no_errors(self):
        assert not self.lg.has_errors()

    def test_no_warnings(self):
        assert self.lg.issues() == []

    def test_top_level_key(self):
        assert list(self.result.keys()) == ['line']

    def test_one_table_instance(self):
        assert len(self.result['line']) == 1

    def test_data_row_count(self):
        # Quelldaten has 278 rows: 1 header + 277 data rows
        assert len(self.result['line'][0]['data']) == 277

    def test_first_row_fields(self):
        first = self.result['line'][0]['data'][0]
        assert first['product'] == 'Alice Mutton'
        assert first['customer'] == 'ANTON'
        assert first['q2'] == 702

    def test_last_row_fields(self):
        last = self.result['line'][0]['data'][-1]
        assert last['product'] == 'Veggie-spread'
        assert last['customer'] == 'WHITC'
        assert last['q3'] == pytest.approx(842.88)
