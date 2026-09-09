#!/usr/bin/env python3
"""
Generate test fixture xlsx files in tests/fixtures/.

Each fixture is a pair:
  {name}_pattern-from-draft.xlsx  — the extraction pattern (programmatically authored here)
  {name}_data.xlsx                — the source spreadsheet

  01_simple_invoice      — cells only, no tables; all values valid
  02_product_catalog     — one table:* group, three mini-table instances
  03_purchase_order      — cells + table with FOOTER; includes one deliberate warning
  04_bank_statement      — account header cells + transaction table with totals footer
  05_expense_report      — employee header + 3 expense-category tables with subtotals
  06_merged_cells        — horizontal merge (title) + vertical merge (category column)
  07_timesheet           — employee header + 7-row daily hours table with total footer
  08_price_list          — supplier header + 3 product-category tables
  09_sales_by_region     — report header + 3 regional tables each with a footer
  10_delivery_note       — supplier/delivery info + ordered-vs-delivered items table
  11_loan_schedule       — loan parameters + 6-row amortisation table with totals
  12_multi_sheet         — data.xlsx has 3 sheets; pattern extracts from 'Details' sheet
  13_hr_attendance       — employee header + quarterly attendance table with totals
  14_named_tables        — var: field in HEADER row captures the mini-table category name
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

from openpyxl import Workbook

ROOT     = os.path.join(os.path.dirname(__file__), '..')
FIXTURES = os.path.join(ROOT, 'tests', 'fixtures')

from grepxcel.pattern_colors import colorize_pattern_file


def save(name: str, pattern_wb: Workbook, data_wb: Workbook):
    folder = os.path.join(FIXTURES, name)
    os.makedirs(folder, exist_ok=True)
    pattern_path = os.path.join(folder, f'{name}_pattern-from-draft.xlsx')
    pattern_wb.save(pattern_path)
    colorize_pattern_file(pattern_path)
    data_path = os.path.join(folder, f'{name}_data.xlsx')
    if not os.path.exists(data_path):
        data_wb.save(data_path)
        print(f'  {name}/  (data.xlsx created)')
    else:
        print(f'  {name}/  (data.xlsx preserved)')


# ─────────────────────────────────────────────────────────────────────────────
# 01: Simple Invoice — sequential cells only, all values valid
# ─────────────────────────────────────────────────────────────────────────────

def fixture_01():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'IGNORE'], ['cell:1', 'inv.number'],
        ['cell:1', 'IGNORE'], ['cell:1', 'inv.date'],
        ['cell:1', 'IGNORE'], ['cell:1', 'inv.due_date'],
        ['cell:1', 'IGNORE'], ['cell:1', 'client.name'],
        ['cell:1', 'IGNORE'], ['cell:1', 'client.email'],
        ['cell:1', 'IGNORE'], ['cell:1', 'amount.net'],
        ['cell:1', 'IGNORE'], ['cell:1', 'amount.vat'],
        ['cell:1', 'IGNORE'], ['cell:1', 'amount.gross'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Invoice No:';  ds['B1'] = 'AB123456'
    ds['C1'] = 'Date:';        ds['D1'] = datetime.datetime(2026, 3, 15)
    ds['E1'] = 'Due Date:';    ds['F1'] = datetime.datetime(2026, 4, 14)
    ds['A3'] = 'Client:';      ds['B3'] = 'Alice Wonderland'
    ds['D3'] = 'Email:';       ds['E3'] = 'alice@wonderland.example'
    ds['A5'] = 'Net:';         ds['B5'] = 120.00
    ds['C5'] = 'VAT:';         ds['D5'] = 24.00
    ds['E5'] = 'Total:';       ds['F5'] = 144.00
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 02: Product Catalog — one table:* group, three mini-table instances
# ─────────────────────────────────────────────────────────────────────────────

def fixture_02():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        ['lbl:', 'col_product', 'string',  'Product'],
        ['lbl:', 'col_sku',     'string',  'SKU'],
        ['lbl:', 'col_qty',     'string',  'Qty'],
        ['lbl:', 'col_price',   'string',  'Price'],
        ['var:', 'item.name',   'string',  r'.+'],
        ['var:', 'item.sku',    'string',  r'[A-Z]{3}[0-9]{3}'],
        ['var:', 'item.qty',    'integer', r'[1-9][0-9]*'],
        ['var:', 'item.price',  'currency', r'.*'],
        ['START:'],
        ['table:*'],
        [None, 'HEADER:1', 'col_product', 'col_sku', 'col_qty', 'col_price'],
        [None, 'DATA:*',   'item.name',   'item.sku', 'item.qty', 'item.price'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    rows = [
        ('Product', 'SKU',    'Qty', 'Price'),
        ('Laptop',  'ELC001', 5,     999.0),
        ('Mouse',   'ELC002', 50,    25.0),
        ('Keyboard','ELC003', 30,    45.0),
        (None, None, None, None),
        ('Product', 'SKU',    'Qty', 'Price'),
        ('Pen',     'STN001', 500,   2.0),
        ('Notebook','STN002', 200,   8.0),
        (None, None, None, None),
        ('Product', 'SKU',    'Qty', 'Price'),
        ('Chair',   'FRN001', 10,    250.0),
        ('Desk',    'FRN002', 5,     450.0),
        ('Lamp',    'FRN003', 20,    35.0),
    ]
    for row in rows:
        ds.append(list(row))
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 03: Purchase Order — cell headers + table with FOOTER + one deliberate warning
# ─────────────────────────────────────────────────────────────────────────────

def fixture_03():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'po_label'],    ['cell:1', 'po.number'],
        ['cell:1', 'date_label'],  ['cell:1', 'po.date'],
        ['cell:1', 'vendor_label'],['cell:1', 'vendor.name'],
        ['cell:1', 'ref_label'],   ['cell:1', 'vendor.ref'],
        ['table:*'],
        [None, 'HEADER:1', 'col_item',     'col_desc', 'col_qty', 'col_price', 'col_total'],
        [None, 'DATA:*',   'row.item',     'row.desc', 'row.qty', 'row.price', 'row.total'],
        [None, 'FOOTER:1', 'footer.label', 'IGNORE',   'IGNORE',  'IGNORE',    'footer.value'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'PO Number:'; ds['B1'] = 'PO-2026'
    ds['C1'] = 'Date:';      ds['D1'] = datetime.datetime(2026, 5, 1)
    ds['E1'] = 'Vendor:';    ds['F1'] = 'Acme Supplies'
    ds['G1'] = 'Ref:';       ds['H1'] = 'ACME001'
    ds['A3'] = 'Item';   ds['B3'] = 'Description'; ds['C3'] = 'Qty'
    ds['D3'] = 'Unit Price'; ds['E3'] = 'Total'
    ds['A4'] = 'Laptop'; ds['B4'] = 'Dell XPS 15';   ds['C4'] = 2; ds['D4'] = 1200.0; ds['E4'] = 2400.0
    ds['A5'] = 'X';      ds['B5'] = 'Monitor Stand'; ds['C5'] = 4; ds['D5'] = 45.0;   ds['E5'] = 180.0
    ds['A6'] = 'Dock';   ds['B6'] = 'USB-C Hub';     ds['C6'] = 6; ds['D6'] = 75.0;   ds['E6'] = 450.0
    ds['A7'] = 'Grand Total'; ds['E7'] = 3030.0
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 04: Bank Statement — account header + transaction table + totals footer
# ─────────────────────────────────────────────────────────────────────────────
#
# Data layout (LR scan):
#   Row 1: Account Holder: | Jane Smith
#   Row 2: Account Number: | GB29NWBK60161331926819
#   Row 3: Statement Period: | January 2026
#   Row 5: Date | Description | Debit | Credit | Balance    ← HEADER
#   Rows 6–10: 5 transactions (zero used for absent debit/credit)
#   Row 11: Totals | | 455.50 | 5700.00 |                  ← FOOTER

def fixture_04():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'acc_lbl'], ['cell:1', 'account.holder'],
        ['cell:1', 'num_lbl'], ['cell:1', 'account.number'],
        ['cell:1', 'per_lbl'], ['cell:1', 'account.period'],
        ['table:*'],
        [None, 'HEADER:1', 'col_date', 'col_desc', 'col_debit', 'col_credit', 'col_balance'],
        [None, 'DATA:*',   'txn.date', 'txn.description', 'txn.debit', 'txn.credit', 'txn.balance'],
        [None, 'FOOTER:1', 'totals.label', 'IGNORE', 'totals.debits', 'totals.credits', 'IGNORE'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Account Holder:';   ds['B1'] = 'Jane Smith'
    ds['A2'] = 'Account Number:';   ds['B2'] = 'GB29NWBK60161331926819'
    ds['A3'] = 'Statement Period:'; ds['B3'] = 'January 2026'
    # row 4 blank
    ds['A5'] = 'Date'; ds['B5'] = 'Description'
    ds['C5'] = 'Debit'; ds['D5'] = 'Credit'; ds['E5'] = 'Balance'
    data = [
        (datetime.datetime(2026,1,1),  'Opening Balance',  0.0,    2500.0, 2500.0),
        (datetime.datetime(2026,1,5),  'Grocery Store',    85.5,   0.0,    2414.5),
        (datetime.datetime(2026,1,10), 'Monthly Salary',   0.0,    3200.0, 5614.5),
        (datetime.datetime(2026,1,15), 'Electricity Bill', 120.0,  0.0,    5494.5),
        (datetime.datetime(2026,1,20), 'Online Shopping',  250.0,  0.0,    5244.5),
    ]
    for i, (dt, desc, deb, crd, bal) in enumerate(data, start=6):
        ds.cell(i, 1, dt); ds.cell(i, 2, desc)
        ds.cell(i, 3, deb); ds.cell(i, 4, crd); ds.cell(i, 5, bal)
    ds['A11'] = 'Totals'; ds['C11'] = 455.5; ds['D11'] = 5700.0
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 05: Expense Report — employee header + 3 expense-category tables with subtotals
# ─────────────────────────────────────────────────────────────────────────────
#
# Three table:* instances (Travel / Meals / Accommodation), each with
# HEADER | DATA rows | FOOTER ("Subtotal").
#
# Data layout:
#   Row 1: Employee: | John Smith | Department: | Engineering
#   Row 2: Period: | Q1 2026 | Manager: | Jane Doe
#   Row 4: Date | Description | Amount | Receipt No  ← table 1 HEADER
#   Rows 5–7: travel rows
#   Row 8: | Subtotal | 450.00 |                     ← FOOTER
#   Row 9: blank
#   Rows 10–: meals table   (same structure)
#   Rows …: accommodation table

def fixture_05():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'emp_lbl'],  ['cell:1', 'emp.name'],
        ['cell:1', 'dept_lbl'], ['cell:1', 'emp.department'],
        ['cell:1', 'per_lbl'],  ['cell:1', 'emp.period'],
        ['cell:1', 'mgr_lbl'],  ['cell:1', 'emp.manager'],
        ['table:*'],
        [None, 'HEADER:1', 'col_date', 'col_desc', 'col_amt', 'col_rcpt'],
        [None, 'DATA:*',   'exp.date', 'exp.description', 'exp.amount', 'exp.receipt'],
        [None, 'FOOTER:1', 'IGNORE',   'sub.label', 'sub.amount', 'IGNORE'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Employee:';   ds['B1'] = 'John Smith'
    ds['C1'] = 'Department:'; ds['D1'] = 'Engineering'
    ds['A2'] = 'Period:';     ds['B2'] = 'Q1 2026'
    ds['C2'] = 'Manager:';    ds['D2'] = 'Jane Doe'

    def write_expense_table(start_row, rows, subtotal):
        ds.cell(start_row, 1, 'Date');       ds.cell(start_row, 2, 'Description')
        ds.cell(start_row, 3, 'Amount');     ds.cell(start_row, 4, 'Receipt No')
        for i, (dt, desc, amt, rcpt) in enumerate(rows, start=start_row + 1):
            ds.cell(i, 1, dt); ds.cell(i, 2, desc)
            ds.cell(i, 3, amt); ds.cell(i, 4, rcpt)
        footer_row = start_row + 1 + len(rows)
        ds.cell(footer_row, 2, 'Subtotal'); ds.cell(footer_row, 3, subtotal)
        return footer_row + 2  # next table starts 2 rows later (blank separator)

    travel = [
        (datetime.datetime(2026,1,15), 'Flight LHR-CDG', 320.0, 'REC001'),
        (datetime.datetime(2026,1,15), 'Train CDG-Paris', 85.0, 'REC002'),
        (datetime.datetime(2026,1,16), 'Taxi to hotel',   45.0, 'REC003'),
    ]
    meals = [
        (datetime.datetime(2026,1,15), 'Team lunch',     55.0, 'REC004'),
        (datetime.datetime(2026,1,16), 'Client dinner',  95.0, 'REC005'),
        (datetime.datetime(2026,1,17), 'Airport snacks', 18.5, 'REC006'),  # 168.5 → round to 168.5
    ]
    accomm = [
        (datetime.datetime(2026,1,15), 'Hotel Paris night 1', 195.0, 'REC007'),
        (datetime.datetime(2026,1,16), 'Hotel Paris night 2', 195.0, 'REC008'),
    ]
    r = write_expense_table(4,  travel, 450.0)
    r = write_expense_table(r,  meals,  168.5)
    write_expense_table(r, accomm, 390.0)
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 06: Merged Cells — two sheets, horizontal title merge + vertical category merges
# ─────────────────────────────────────────────────────────────────────────────
#
# Sheet '2026' (active): 1 table, cols A-C, Hardware+Software only.
#   Merges: A1:C1 (title), A3:A5 (Hardware), A6:A8 (Software)
#
# Sheet '2025': 3 side-by-side tables arranged diagonally (one column right each),
#   each adding an 'Others' category and doubling all prices.
#   Table 1: col A, rows  2–13  — ×1 prices, merges A3:A5 / A6:A8 / A9:A12
#   Table 2: col C, rows 18–29  — ×2 prices, merges C19:C21 / C22:C24 / C25:C28
#   Table 3: col D, rows 32–43  — ×4 prices, merges D33:D35 / D36:D38 / D39:D42
#
# Pattern uses table:* so the engine finds all three instances on '2025'.

def fixture_06():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'col_category', 'string',   'Category'],
        ['lbl:', 'col_product',  'string',   'Product'],
        ['lbl:', 'col_price',    'string',   'Price'],
        ['var:', 'report.title',    'string',   r'SALES.*'],
        ['var:', 'item.category',   'string',   r'.+'],
        ['var:', 'item.name',       'string',   r'.+'],
        ['var:', 'item.price',      'currency', r'.*'],
        ['var:', 'footer.label',    'string',   'Grand Total'],
        ['var:', 'footer.total',    'currency', r'.*'],
        ['START:'],
        ['cell:1', 'report.title'],
        ['cell:1', 'IGNORE'],
        ['cell:1', 'IGNORE'],
        ['table:*'],
        [None, 'HEADER:1', 'col_category', 'col_product', 'col_price'],
        [None, 'DATA:*',   'item.category', 'item.name',  'item.price'],
        [None, 'FOOTER:1', 'IGNORE',        'footer.label', 'footer.total'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook()

    # ── Sheet '2026' (active) — Hardware + Software only ─────────────────────
    s26 = d.active; s26.title = '2026'
    s26['A1'] = 'SALES CATALOGUE 2026'; s26.merge_cells('A1:C1')
    s26['A2'] = 'Category'; s26['B2'] = 'Product'; s26['C2'] = 'Price'
    s26['A3'] = 'Hardware'; s26['B3'] = 'Laptop 15"';    s26['C3'] = 1200.0
    s26['B4'] = 'Wireless Mouse';                         s26['C4'] = 29.99
    s26['B5'] = 'Keyboard';                               s26['C5'] = 89.99
    s26.merge_cells('A3:A5')
    s26['A6'] = 'Software'; s26['B6'] = 'Office Suite';  s26['C6'] = 300.0
    s26['B7'] = 'Design Suite';                           s26['C7'] = 600.0
    s26['B8'] = 'Dev Tools';                              s26['C8'] = 150.0
    s26.merge_cells('A6:A8')
    s26['B9'] = 'Grand Total'; s26['C9'] = 2369.98

    # ── Sheet '2025' — 3 diagonal tables, prices ×1 / ×2 / ×4 ───────────────
    s25 = d.create_sheet('2025')
    s25['A1'] = 'SALES CATALOGUE 2025'; s25.merge_cells('A1:C1')

    _cats = [
        ('Hardware', [('Laptop 15"', 1200.0), ('Wireless Mouse', 29.99), ('Keyboard', 89.99)]),
        ('Software', [('Office Suite', 300.0), ('Design Suite', 600.0), ('Dev Tools', 150.0)]),
        ('Others',   [('other-1', 100.0), ('other-2', 200.0), ('other-3', 300.0), ('other-4', 400.0)]),
    ]

    def _write_table(ws, start_row, col, cats, factor, footer_total):
        ws.cell(start_row, col,     'Category')
        ws.cell(start_row, col + 1, 'Product')
        ws.cell(start_row, col + 2, 'Price')
        cur = start_row + 1
        for cat_name, items in cats:
            cat_start = cur
            for i, (name, price) in enumerate(items):
                if i == 0:
                    ws.cell(cur, col, cat_name)
                ws.cell(cur, col + 1, name)
                ws.cell(cur, col + 2, round(price * factor, 2))
                cur += 1
            if len(items) > 1:
                ws.merge_cells(start_row=cat_start, start_column=col,
                               end_row=cur - 1,    end_column=col)
        ws.cell(cur, col + 1, 'Grand Total')
        ws.cell(cur, col + 2, footer_total)

    _write_table(s25, start_row=2,  col=1, cats=_cats, factor=1, footer_total=3369.98)
    _write_table(s25, start_row=18, col=3, cats=_cats, factor=2, footer_total=6739.96)
    _write_table(s25, start_row=32, col=4, cats=_cats, factor=4, footer_total=13479.92)

    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 07: Timesheet — employee header + 7-row daily hours table + total footer
# ─────────────────────────────────────────────────────────────────────────────

def fixture_07():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'emp_lbl'],  ['cell:1', 'emp.name'],
        ['cell:1', 'week_lbl'], ['cell:1', 'emp.week_start'],
        ['cell:1', 'proj_lbl'], ['cell:1', 'emp.project'],
        ['table:1'],
        [None, 'HEADER:1', 'col_day', 'col_hours', 'col_task'],
        [None, 'DATA:*',   'day.name', 'day.hours', 'day.task'],
        [None, 'FOOTER:1', 'total.label', 'total.hours', 'IGNORE'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Employee:';     ds['B1'] = 'Bob Martin'
    ds['A2'] = 'Week Starting:';ds['B2'] = datetime.datetime(2026, 5, 11)
    ds['A3'] = 'Project:';      ds['B3'] = 'grepxcel v2'
    ds['A5'] = 'Day'; ds['B5'] = 'Hours'; ds['C5'] = 'Task'
    days = [
        ('Monday',    8.0,  'Architecture review'),
        ('Tuesday',   7.5,  'Engine development'),
        ('Wednesday', 8.0,  'Engine development'),
        ('Thursday',  6.0,  'Code review'),
        ('Friday',    7.5,  'Testing'),
        ('Saturday',  0.0,  '-'),
        ('Sunday',    0.0,  '-'),
    ]
    for i, (day, hrs, task) in enumerate(days, start=6):
        ds.cell(i, 1, day); ds.cell(i, 2, hrs); ds.cell(i, 3, task)
    ds['A13'] = 'Total Hours'; ds['B13'] = 37.0
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 08: Price List — supplier header + 3 product-category tables
# ─────────────────────────────────────────────────────────────────────────────

def fixture_08():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'sup_lbl',   'string', 'Supplier:'],
        ['lbl:', 'ref_lbl',   'string', 'Ref:'],
        ['lbl:', 'from_lbl',  'string', 'Valid From:'],
        ['lbl:', 'to_lbl',    'string', 'Valid To:'],
        ['lbl:', 'col_code',  'string', 'Code'],
        ['lbl:', 'col_name',  'string', 'Product'],
        ['lbl:', 'col_unit',  'string', 'Unit'],
        ['lbl:', 'col_price', 'string', 'Unit Price'],
        ['var:', 'supplier.name', 'string',   r'.+'],
        ['var:', 'supplier.ref',  'string',   r'[A-Z0-9\-]+'],
        ['var:', 'validity.from', 'date',     r'.*'],
        ['var:', 'validity.to',   'date',     r'.*'],
        ['var:', 'prod.code',  'string',   r'[A-Z]{2}[0-9]{4}'],
        ['var:', 'prod.name',  'string',   r'.+'],
        ['var:', 'prod.unit',  'string',   r'.+'],
        ['var:', 'prod.price', 'currency', r'.*'],
        ['START:'],
        ['cell:1', 'sup_lbl'],  ['cell:1', 'supplier.name'],
        ['cell:1', 'ref_lbl'],  ['cell:1', 'supplier.ref'],
        ['cell:1', 'from_lbl'], ['cell:1', 'validity.from'],
        ['cell:1', 'to_lbl'],   ['cell:1', 'validity.to'],
        ['table:*'],
        [None, 'HEADER:1', 'col_code', 'col_name', 'col_unit', 'col_price'],
        [None, 'DATA:*',   'prod.code', 'prod.name', 'prod.unit', 'prod.price'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Supplier:';   ds['B1'] = 'TechDistrib GmbH'
    ds['C1'] = 'Ref:';        ds['D1'] = 'TD-2026-Q2'
    ds['A2'] = 'Valid From:'; ds['B2'] = datetime.datetime(2026, 4, 1)
    ds['C2'] = 'Valid To:';   ds['D2'] = datetime.datetime(2026, 6, 30)

    def write_price_table(start, items):
        ds.cell(start, 1, 'Code');  ds.cell(start, 2, 'Product')
        ds.cell(start, 3, 'Unit');  ds.cell(start, 4, 'Unit Price')
        for i, (code, name, unit, price) in enumerate(items, start + 1):
            ds.cell(i, 1, code); ds.cell(i, 2, name)
            ds.cell(i, 3, unit); ds.cell(i, 4, price)
        return start + 1 + len(items) + 1  # next table start (blank row gap)

    electronics = [
        ('EL0001', 'USB-C Hub 7-port',   'pcs',  49.99),
        ('EL0002', 'Wireless Keyboard',  'pcs',  79.00),
        ('EL0003', 'HD Webcam 1080p',    'pcs',  55.50),
    ]
    cables = [
        ('CA0001', 'HDMI 2.0 2m cable', 'pcs',   9.99),
        ('CA0002', 'USB-C to USB-A 1m', 'pcs',   6.49),
        ('CA0003', 'DisplayPort 1.4 2m','pcs',  12.99),
        ('CA0004', 'Ethernet Cat6 5m',  'pcs',   8.75),
    ]
    storage = [
        ('ST0001', 'SSD 1TB 2.5"',      'pcs',  89.00),
        ('ST0002', 'NVMe SSD 500GB',    'pcs',  65.00),
    ]
    r = write_price_table(4, electronics)
    r = write_price_table(r, cables)
    write_price_table(r, storage)
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 09: Sales by Region — report header + 3 regional tables each with footer
# ─────────────────────────────────────────────────────────────────────────────

def fixture_09():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'per_lbl', 'string', 'Period:'],
        ['lbl:', 'cur_lbl', 'string', 'Currency:'],
        ['lbl:', 'col_product', 'string', 'Product'],
        ['lbl:', 'col_units',   'string', 'Units'],
        ['lbl:', 'col_revenue', 'string', 'Revenue'],
        ['lbl:', 'col_target',  'string', 'Target'],
        ['var:', 'report.period',   'string',   r'.+'],
        ['var:', 'report.currency', 'string',   r'[A-Z]{3}'],
        ['var:', 'sale.product', 'string',   r'.+'],
        ['var:', 'sale.units',   'integer',  r'[0-9]+'],
        ['var:', 'sale.revenue', 'currency', r'.*'],
        ['var:', 'sale.target',  'currency', r'.*'],
        ['var:', 'region.label',   'string',   r'.+ Total'],
        ['var:', 'region.revenue', 'currency', r'.*'],
        ['var:', 'region.target',  'currency', r'.*'],
        ['START:'],
        ['cell:1', 'per_lbl'], ['cell:1', 'report.period'],
        ['cell:1', 'cur_lbl'], ['cell:1', 'report.currency'],
        ['table:*'],
        [None, 'HEADER:1', 'col_product', 'col_units', 'col_revenue', 'col_target'],
        [None, 'DATA:*',   'sale.product', 'sale.units', 'sale.revenue', 'sale.target'],
        [None, 'FOOTER:1', 'region.label', 'IGNORE', 'region.revenue', 'region.target'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Period:';   ds['B1'] = 'Q1 2026'
    ds['C1'] = 'Currency:'; ds['D1'] = 'EUR'

    def write_region(start, label, items, rev_total, tgt_total):
        ds.cell(start, 1, 'Product'); ds.cell(start, 2, 'Units')
        ds.cell(start, 3, 'Revenue'); ds.cell(start, 4, 'Target')
        for i, (prod, units, rev, tgt) in enumerate(items, start + 1):
            ds.cell(i, 1, prod); ds.cell(i, 2, units)
            ds.cell(i, 3, rev);  ds.cell(i, 4, tgt)
        fr = start + 1 + len(items)
        ds.cell(fr, 1, label); ds.cell(fr, 3, rev_total); ds.cell(fr, 4, tgt_total)
        return fr + 2

    europe = [('Laptop Pro', 320, 384000.0, 400000.0),
              ('Tablet Air', 210, 147000.0, 150000.0),
              ('Phone X',    450, 225000.0, 250000.0)]
    americas = [('Laptop Pro', 510, 612000.0, 600000.0),
                ('Tablet Air', 380, 266000.0, 250000.0)]
    asia = [('Laptop Pro', 290, 348000.0, 350000.0),
            ('Phone X',    820, 410000.0, 400000.0),
            ('Smart Watch',600, 180000.0, 200000.0)]

    r = write_region(3, 'Europe Total',   europe,   756000.0, 800000.0)
    r = write_region(r, 'Americas Total', americas, 878000.0, 850000.0)
    write_region(r,     'Asia Total',     asia,     938000.0, 950000.0)
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 10: Delivery Note — supplier/delivery info + ordered-vs-delivered items table
# ─────────────────────────────────────────────────────────────────────────────

def fixture_10():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'sup_lbl'],  ['cell:1', 'supplier.name'],
        ['cell:1', 'ord_lbl'],  ['cell:1', 'order.number'],
        ['cell:1', 'del_lbl'],  ['cell:1', 'delivery.date'],
        ['cell:1', 'addr_lbl'], ['cell:1', 'delivery.address'],
        ['table:1'],
        [None, 'HEADER:1', 'col_item', 'col_ordered', 'col_delivered', 'col_unit', 'col_notes'],
        [None, 'DATA:*',   'line.item', 'line.ordered', 'line.delivered', 'line.unit', 'line.notes'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Supplier:';      ds['B1'] = 'Acme Supplies Ltd'
    ds['C1'] = 'Order No:';      ds['D1'] = 'ORD-20260501'
    ds['A2'] = 'Delivery Date:'; ds['B2'] = datetime.datetime(2026, 5, 8)
    ds['C2'] = 'Deliver To:';    ds['D2'] = 'Warehouse B, Unit 12'
    ds['A4'] = 'Item'; ds['B4'] = 'Ordered'; ds['C4'] = 'Delivered'
    ds['D4'] = 'Unit'; ds['E4'] = 'Notes'
    items = [
        ('USB-C Cables',    100, 100, 'box', 'OK'),
        ('Wireless Mice',    50,  48, 'pcs', '2 damaged in transit'),
        ('HDMI Cables',      75,  75, 'pcs', 'OK'),
        ('Laptop Stands',    20,  20, 'pcs', 'OK'),
        ('Monitor Risers',   30,  25, 'pcs', 'Backorder: 5 units'),
    ]
    for i, (item, ordered, delivered, unit, notes) in enumerate(items, start=5):
        ds.cell(i, 1, item); ds.cell(i, 2, ordered); ds.cell(i, 3, delivered)
        ds.cell(i, 4, unit); ds.cell(i, 5, notes)
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 11: Loan Schedule — loan parameters + 6-row amortisation table with totals
# ─────────────────────────────────────────────────────────────────────────────

def fixture_11():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['var:', 'summary.label',      'string',   'Totals'],
        ['var:', 'summary.payment',    'currency', r'.*'],
        ['var:', 'summary.principal',  'currency', r'.*'],
        ['var:', 'summary.interest',   'currency', r'.*'],
        ['START:'],
        ['cell:1', 'amt_lbl'],   ['cell:1', 'loan.amount'],
        ['cell:1', 'rate_lbl'],  ['cell:1', 'loan.rate'],
        ['cell:1', 'term_lbl'],  ['cell:1', 'loan.term'],
        ['cell:1', 'start_lbl'], ['cell:1', 'loan.start'],
        ['table:1'],
        [None, 'HEADER:1', 'col_no', 'col_date', 'col_payment', 'col_principal', 'col_interest', 'col_balance'],
        [None, 'DATA:*',   'amort.payment_no', 'amort.date', 'amort.payment', 'amort.principal', 'amort.interest', 'amort.balance'],
        [None, 'FOOTER:1', 'summary.label', 'IGNORE', 'summary.payment', 'summary.principal', 'summary.interest', 'IGNORE'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Loan Amount:';    ds['B1'] = 10000.0
    ds['C1'] = 'Monthly Rate:';   ds['D1'] = '0.5%'
    ds['A2'] = 'Term (months):';  ds['B2'] = 6
    ds['C2'] = 'Start Date:';     ds['D2'] = datetime.datetime(2026, 2, 1)
    ds['A4'] = 'Payment #'; ds['B4'] = 'Date';      ds['C4'] = 'Payment'
    ds['D4'] = 'Principal'; ds['E4'] = 'Interest';  ds['F4'] = 'Balance'
    schedule = [
        (1, datetime.datetime(2026,2,1),  1702.0, 1652.0, 50.0,  8348.0),
        (2, datetime.datetime(2026,3,1),  1702.0, 1660.3, 41.7,  6687.7),
        (3, datetime.datetime(2026,4,1),  1702.0, 1668.6, 33.4,  5019.1),
        (4, datetime.datetime(2026,5,1),  1702.0, 1676.9, 25.1,  3342.2),
        (5, datetime.datetime(2026,6,1),  1702.0, 1685.3, 16.7,  1656.9),
        (6, datetime.datetime(2026,7,1),  1665.2, 1656.9,  8.3,     0.0),
    ]
    for i, row in enumerate(schedule, start=5):
        for j, val in enumerate(row, start=1):
            ds.cell(i, j, val)
    ds['A11'] = 'Totals'; ds['C11'] = 10175.2; ds['D11'] = 10000.0; ds['E11'] = 175.2
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 12: Multi-Sheet — data.xlsx has 3 sheets; extraction targets 'Details'
# ─────────────────────────────────────────────────────────────────────────────
#
# Tests Engine.process(sheet='Details').  Pattern and pattern-processing logic
# are sheet-agnostic; only the data file has multiple sheets.

def fixture_12():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        ['lbl:', 'dept_lbl', 'string', 'Department:'],
        ['lbl:', 'code_lbl', 'string', 'Cost Centre:'],
        ['lbl:', 'col_name',   'string', 'Name'],
        ['lbl:', 'col_role',   'string', 'Role'],
        ['lbl:', 'col_salary', 'string', 'Salary'],
        ['var:', 'dept.name', 'string',   r'.+'],
        ['var:', 'dept.code', 'string',   r'[A-Z]{3}-[0-9]{3}'],
        ['var:', 'emp.name',   'string',   r'.+'],
        ['var:', 'emp.role',   'string',   r'.+'],
        ['var:', 'emp.salary', 'currency', r'.*'],
        ['START:'],
        ['cell:1', 'dept_lbl'], ['cell:1', 'dept.name'],
        ['cell:1', 'code_lbl'], ['cell:1', 'dept.code'],
        ['table:1'],
        [None, 'HEADER:1', 'col_name', 'col_role', 'col_salary'],
        [None, 'DATA:*',   'emp.name', 'emp.role', 'emp.salary'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook()
    # Sheet 1: Summary (decoy — should not be processed)
    summary = d.active; summary.title = 'Summary'
    summary['A1'] = 'Q1 2026 Summary'
    summary['A2'] = 'Department:'; summary['B2'] = 'DO NOT EXTRACT'

    # Sheet 2: Details (target sheet)
    details = d.create_sheet('Details')
    details['A1'] = 'Department:'; details['B1'] = 'Engineering'
    details['A2'] = 'Cost Centre:'; details['B2'] = 'ENG-001'
    details['A4'] = 'Name'; details['B4'] = 'Role'; details['C4'] = 'Salary'
    staff = [
        ('Alice Brown',  'Lead Engineer',  85000.0),
        ('Bob Chen',     'Senior Dev',     72000.0),
        ('Carol Davis',  'Junior Dev',     55000.0),
    ]
    for i, (name, role, sal) in enumerate(staff, start=5):
        details.cell(i, 1, name); details.cell(i, 2, role); details.cell(i, 3, sal)

    # Sheet 3: Notes (another decoy)
    notes = d.create_sheet('Notes')
    notes['A1'] = 'Reviewed by Finance'

    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 13: HR Attendance — employee header + quarterly attendance table with totals
# ─────────────────────────────────────────────────────────────────────────────

def fixture_13():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
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
        ['cell:1', 'emp_lbl'],  ['cell:1', 'emp.name'],
        ['cell:1', 'id_lbl'],   ['cell:1', 'emp.id'],
        ['cell:1', 'year_lbl'], ['cell:1', 'emp.year'],
        ['table:1'],
        [None, 'HEADER:1', 'col_month', 'col_working', 'col_absent', 'col_reason'],
        [None, 'DATA:*',   'att.month', 'att.working', 'att.absent', 'att.reason'],
        [None, 'FOOTER:1', 'totals.label', 'totals.working', 'totals.absent', 'IGNORE'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Employee:';    ds['B1'] = 'Sarah Connor'
    ds['C1'] = 'Employee ID:'; ds['D1'] = 'EMP-0042'
    ds['A2'] = 'Year:';        ds['B2'] = 2026
    ds['A4'] = 'Month'; ds['B4'] = 'Working Days'
    ds['C4'] = 'Days Absent'; ds['D4'] = 'Reason'
    months = [
        ('January',  22, 1, 'Sick leave'),
        ('February', 20, 0, '-'),
        ('March',    23, 2, 'Personal leave'),
    ]
    for i, (m, w, a, r) in enumerate(months, start=5):
        ds.cell(i, 1, m); ds.cell(i, 2, w)
        ds.cell(i, 3, a); ds.cell(i, 4, r)
    ds['A8'] = 'Q1 Total'; ds['B8'] = 65; ds['C8'] = 3
    return p, d


# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# 14: Named Tables — var: field in HEADER captures the category name
# ─────────────────────────────────────────────────────────────────────────────
#
# Demonstrates that var: fields work in HEADER rows: the first column of each
# HEADER row is the category name (extracted into instance['header']['category'])
# while the remaining columns are lbl: anchors (SKU / Qty / Price).
#
# The engine anchors on the lbl: columns — "SKU" in col B is the specificity
# anchor that distinguishes HEADER rows from DATA rows.
#
# Data layout (3 mini-tables, same column positions):
#   Row  1: Electronics | SKU    | Qty  | Price   ← HEADER 1
#   Rows 2-4: data rows (Laptop/Mouse/Keyboard)
#   Row  5: (blank separator)
#   Row  6: Stationery  | SKU    | Qty  | Price   ← HEADER 2
#   Rows 7-8: data rows (Pen/Notebook)
#   Row  9: (blank separator)
#   Row 10: Furniture   | SKU    | Qty  | Price   ← HEADER 3
#   Rows 11-13: data rows (Chair/Desk/Lamp)

def fixture_14():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        # var: field — extracted from col A of each HEADER row
        ['var:', 'header.category', 'string',   r'.+'],
        # lbl: anchors — used only for HEADER matching, never in output
        ['lbl:', 'col_sku',         'string',   'SKU'],
        ['lbl:', 'col_qty',         'string',   'Qty'],
        ['lbl:', 'col_price',       'string',   'Price'],
        # data fields
        ['var:', 'item.name',       'string',   r'.+'],
        ['var:', 'item.sku',        'string',   r'[A-Z]{3}[0-9]{3}'],
        ['var:', 'item.qty',        'integer',  r'[1-9][0-9]*'],
        ['var:', 'item.price',      'currency', r'.*'],
        ['START:'],
        ['table:*'],
        [None, 'HEADER:1', 'header.category', 'col_sku',  'col_qty',  'col_price'],
        [None, 'DATA:*',   'item.name',        'item.sku', 'item.qty', 'item.price'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    rows = [
        # Table 1 — Electronics (category name in col A of HEADER)
        ('Electronics', 'SKU',    'Qty',  'Price'),
        ('Laptop',      'ELC001', 5,      999.0),
        ('Mouse',       'ELC002', 50,     25.0),
        ('Keyboard',    'ELC003', 30,     45.0),
        (None, None, None, None),
        # Table 2 — Stationery
        ('Stationery',  'SKU',    'Qty',  'Price'),
        ('Pen',         'STN001', 500,    2.0),
        ('Notebook',    'STN002', 200,    8.0),
        (None, None, None, None),
        # Table 3 — Furniture
        ('Furniture',   'SKU',    'Qty',  'Price'),
        ('Chair',       'FRN001', 10,     250.0),
        ('Desk',        'FRN002', 5,      450.0),
        ('Lamp',        'FRN003', 20,     35.0),
    ]
    for row in rows:
        ds.append(list(row))
    return p, d


def main():
    print('Generating test fixtures...')
    save('01_simple_invoice',  *fixture_01())
    save('02_product_catalog', *fixture_02())
    save('03_purchase_order',  *fixture_03())
    save('04_bank_statement',  *fixture_04())
    save('05_expense_report',  *fixture_05())
    save('06_merged_cells',    *fixture_06())
    save('07_timesheet',       *fixture_07())
    save('08_price_list',      *fixture_08())
    save('09_sales_by_region', *fixture_09())
    save('10_delivery_note',   *fixture_10())
    save('11_loan_schedule',   *fixture_11())
    save('12_multi_sheet',     *fixture_12())
    save('13_hr_attendance',   *fixture_13())
    save('14_named_tables',    *fixture_14())
    print('Done.')


if __name__ == '__main__':
    main()
