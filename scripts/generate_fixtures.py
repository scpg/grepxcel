"""
Generate test fixture xlsx files in tests/fixtures/.

Each fixture is a pair: pattern.xlsx (what to look for) + data.xlsx (the actual spreadsheet).

  01_simple_invoice   — cells only, no tables; all values valid
  02_product_catalog  — one table:* group, three mini-table instances
  03_purchase_order   — cells + table with FOOTER; includes one deliberate warning
"""

import os
import sys
import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from openpyxl import Workbook

ROOT = os.path.join(os.path.dirname(__file__), '..')
FIXTURES = os.path.join(ROOT, 'tests', 'fixtures')


def save(name: str, pattern_wb: Workbook, data_wb: Workbook):
    folder = os.path.join(FIXTURES, name)
    os.makedirs(folder, exist_ok=True)
    pattern_wb.save(os.path.join(folder, 'pattern.xlsx'))
    data_wb.save(os.path.join(folder, 'data.xlsx'))
    print(f'  {name}/')


# ─────────────────────────────────────────────────────────────────────────────
# 01: Simple Invoice — sequential cells only, all values valid
# ─────────────────────────────────────────────────────────────────────────────
#
# Pattern uses IGNORE to skip label cells ("Invoice No:", "Date:", etc.)
# and named fields to extract the actual values that follow each label.
#
# Data layout (LR scan reads left→right, top→bottom):
#   Row 1: Invoice No: | AB123456 | Date: | 2026-03-15 | Due Date: | 2026-04-14
#   Row 3: Client:     | Alice Wonderland  | Email: | alice@wonderland.example
#   Row 5: Net: | 120.0 | VAT: | 24.0 | Total: | 144.0

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
        ['cell:1', 'IGNORE'],        # "Invoice No:"
        ['cell:1', 'inv.number'],
        ['cell:1', 'IGNORE'],        # "Date:"
        ['cell:1', 'inv.date'],
        ['cell:1', 'IGNORE'],        # "Due Date:"
        ['cell:1', 'inv.due_date'],
        ['cell:1', 'IGNORE'],        # "Client:"
        ['cell:1', 'client.name'],
        ['cell:1', 'IGNORE'],        # "Email:"
        ['cell:1', 'client.email'],
        ['cell:1', 'IGNORE'],        # "Net:"
        ['cell:1', 'amount.net'],
        ['cell:1', 'IGNORE'],        # "VAT:"
        ['cell:1', 'amount.vat'],
        ['cell:1', 'IGNORE'],        # "Total:"
        ['cell:1', 'amount.gross'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    ds['A1'] = 'Invoice No:';          ds['B1'] = 'AB123456'
    ds['C1'] = 'Date:';                ds['D1'] = datetime.datetime(2026, 3, 15)
    ds['E1'] = 'Due Date:';            ds['F1'] = datetime.datetime(2026, 4, 14)
    ds['A3'] = 'Client:';              ds['B3'] = 'Alice Wonderland'
    ds['D3'] = 'Email:';               ds['E3'] = 'alice@wonderland.example'
    ds['A5'] = 'Net:';                 ds['B5'] = 120.00
    ds['C5'] = 'VAT:';                 ds['D5'] = 24.00
    ds['E5'] = 'Total:';               ds['F5'] = 144.00
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 02: Product Catalog — one table:* group, three mini-table instances
# ─────────────────────────────────────────────────────────────────────────────
#
# Each mini-table is a product category with a header row ("Product|SKU|Qty|Price")
# followed by any number of data rows.  Tables are separated by a blank row.
#
# Data layout:
#   Row  1: Product | SKU | Qty | Price     ← table 1 header (Electronics)
#   Rows 2–4: data rows
#   Row  5: (empty)
#   Row  6: Product | SKU | Qty | Price     ← table 2 header (Stationery)
#   Rows 7–8: data rows
#   Row  9: (empty)
#   Row 10: Product | SKU | Qty | Price     ← table 3 header (Furniture)
#   Rows 11–13: data rows

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
        # Table 1 — Electronics
        ('Product',   'SKU',    'Qty',  'Price'),
        ('Laptop',    'ELC001', 5,      999.0),
        ('Mouse',     'ELC002', 50,     25.0),
        ('Keyboard',  'ELC003', 30,     45.0),
        (None, None, None, None),   # separator
        # Table 2 — Stationery
        ('Product',   'SKU',    'Qty',  'Price'),
        ('Pen',       'STN001', 500,    2.0),
        ('Notebook',  'STN002', 200,    8.0),
        (None, None, None, None),   # separator
        # Table 3 — Furniture
        ('Product',   'SKU',    'Qty',  'Price'),
        ('Chair',     'FRN001', 10,     250.0),
        ('Desk',      'FRN002', 5,      450.0),
        ('Lamp',      'FRN003', 20,     35.0),
    ]
    for row in rows:
        ds.append(list(row))
    return p, d


# ─────────────────────────────────────────────────────────────────────────────
# 03: Purchase Order — cell headers + table with FOOTER + one deliberate warning
# ─────────────────────────────────────────────────────────────────────────────
#
# Demonstrates: mixing cell:1 and table:*, FOOTER detection, and a
# validation warning (row.item uses regex .{3,50}; one item named "X"
# is too short and triggers a warning with the full field context).
#
# Data layout:
#   Row 1: PO Number: | PO-2026 | Date: | 2026-05-01 | Vendor: | Acme Supplies | Ref: | ACME001
#   Row 2: (empty)
#   Row 3: Item | Description | Qty | Unit Price | Total    ← HEADER
#   Row 4: Laptop      | Dell XPS 15 | 2 | 1200.0 | 2400.0
#   Row 5: X           | Monitor     | 4 |   45.0 |  180.0  ← "X" fails .{3,50} → warning
#   Row 6: Dock        | USB-C Hub   | 6 |   75.0 |  450.0
#   Row 7: Grand Total |             |   |         | 3030.0  ← FOOTER

def fixture_03():
    p = Workbook(); ps = p.active; ps.title = 'Pattern'
    for row in [
        ['config:', 'read.direction', 'LR'],
        ['config:', 'currency.sign',  '€'],
        # label anchors (match only, never in output)
        ['lbl:', 'po_label',     'string',   'PO Number:'],
        ['lbl:', 'date_label',   'string',   'Date:'],
        ['lbl:', 'vendor_label', 'string',   'Vendor:'],
        ['lbl:', 'ref_label',    'string',   'Ref:'],
        # table column header anchors
        ['lbl:', 'col_item',     'string',   'Item'],
        ['lbl:', 'col_desc',     'string',   'Description'],
        ['lbl:', 'col_qty',      'string',   'Qty'],
        ['lbl:', 'col_price',    'string',   'Unit Price'],
        ['lbl:', 'col_total',    'string',   'Total'],
        # cell value fields
        ['var:', 'po.number',    'string',   r'PO-[0-9]{4}'],
        ['var:', 'po.date',      'date',     r'.*'],
        ['var:', 'vendor.name',  'string',   r'.+'],
        ['var:', 'vendor.ref',   'string',   r'[A-Z0-9]+'],
        # table data fields (.{3,50} → "X" in the data will fail → warning)
        ['var:', 'row.item',     'string',   r'.{3,50}'],
        ['var:', 'row.desc',     'string',   r'.*'],
        ['var:', 'row.qty',      'integer',  r'[1-9][0-9]*'],
        ['var:', 'row.price',    'currency', r'.*'],
        ['var:', 'row.total',    'currency', r'.*'],
        # footer value fields
        ['var:', 'footer.label', 'string',   'Grand Total'],
        ['var:', 'footer.value', 'currency', r'.*'],
        ['START:'],
        ['cell:1', 'po_label'],
        ['cell:1', 'po.number'],
        ['cell:1', 'date_label'],
        ['cell:1', 'po.date'],
        ['cell:1', 'vendor_label'],
        ['cell:1', 'vendor.name'],
        ['cell:1', 'ref_label'],
        ['cell:1', 'vendor.ref'],
        ['table:*'],
        [None, 'HEADER:1', 'col_item',      'col_desc',  'col_qty',  'col_price',  'col_total'],
        [None, 'DATA:*',   'row.item',       'row.desc',  'row.qty',  'row.price',  'row.total'],
        [None, 'FOOTER:1', 'footer.label',   'IGNORE',    'IGNORE',   'IGNORE',     'footer.value'],
        ['END:'],
    ]:
        ps.append(row)

    d = Workbook(); ds = d.active; ds.title = 'Sheet1'
    # Row 1: header cells
    ds['A1'] = 'PO Number:';   ds['B1'] = 'PO-2026'
    ds['C1'] = 'Date:';        ds['D1'] = datetime.datetime(2026, 5, 1)
    ds['E1'] = 'Vendor:';      ds['F1'] = 'Acme Supplies'
    ds['G1'] = 'Ref:';         ds['H1'] = 'ACME001'
    # Row 2: empty (separator)
    # Row 3: table header
    ds['A3'] = 'Item';   ds['B3'] = 'Description'
    ds['C3'] = 'Qty';    ds['D3'] = 'Unit Price';  ds['E3'] = 'Total'
    # Rows 4–6: line items ("X" on row 5 intentionally fails .{3,50})
    ds['A4'] = 'Laptop';       ds['B4'] = 'Dell XPS 15';    ds['C4'] = 2;  ds['D4'] = 1200.0; ds['E4'] = 2400.0
    ds['A5'] = 'X';            ds['B5'] = 'Monitor Stand';  ds['C5'] = 4;  ds['D5'] = 45.0;   ds['E5'] = 180.0
    ds['A6'] = 'Dock';         ds['B6'] = 'USB-C Hub';      ds['C6'] = 6;  ds['D6'] = 75.0;   ds['E6'] = 450.0
    # Row 7: footer
    ds['A7'] = 'Grand Total';  ds['E7'] = 3030.0
    return p, d


# ─────────────────────────────────────────────────────────────────────────────

def main():
    print('Generating test fixtures...')
    save('01_simple_invoice',  *fixture_01())
    save('02_product_catalog', *fixture_02())
    save('03_purchase_order',  *fixture_03())
    print('Done.')


if __name__ == '__main__':
    main()
