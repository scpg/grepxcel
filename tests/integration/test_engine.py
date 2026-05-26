"""
Integration tests: run the engine against each test fixture and assert
on the extracted structure and values.
"""

import os
import datetime
import pytest
from engine import Engine, Logger, VerbosityLevel

FIXTURES = os.path.join(os.path.dirname(__file__), '..', 'fixtures')


def run(fixture_name: str, sheet=None):
    folder = os.path.join(FIXTURES, fixture_name)
    lg = Logger(level=VerbosityLevel.QUIET)
    kwargs = {} if sheet is None else {'sheet': sheet}
    result = Engine().process(
        pattern_file=os.path.join(folder, 'pattern.xlsx'),
        data_file=os.path.join(folder, 'data.xlsx'),
        logger=lg,
        **kwargs,
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
