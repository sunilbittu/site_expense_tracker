import shutil
import os
import sys
import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
from build_dashboard import MONTH_SHEETS

SOURCE_WORKBOOK = os.path.join(os.path.dirname(__file__), '..', 'expense_tracker_2026.xlsm')


def sample_workbook(tmp_dir):
    """Copy the real workbook into tmp_dir, clear all pre-existing demo data
    from every month sheet, and inject one sample Receipts row and one
    sample Payments row into Jan2026. Returns the new path."""
    dest = os.path.join(tmp_dir, 'sample.xlsm')
    shutil.copyfile(SOURCE_WORKBOOK, dest)

    wb = openpyxl.load_workbook(dest, keep_vba=True)

    # Clear any pre-existing demo data on every month sheet so the fixture
    # is self-contained regardless of what the live workbook currently holds.
    # The entry table's row count (and thus where the Summary block starts)
    # varies per sheet, so find each sheet's own boundary rather than
    # assuming a fixed row range - clearing past it would wipe the Summary
    # formulas (Opening Balance, Total Receipts, etc) themselves.
    for month in MONTH_SHEETS:
        month_ws = wb[month]
        summary_row = None
        for r in range(1, month_ws.max_row + 1):
            if month_ws.cell(row=r, column=1).value == 'Summary':
                summary_row = r
                break
        if summary_row is None:
            raise ValueError("Could not find 'Summary' row in sheet %r" % month)
        for row in range(3, summary_row):
            for col in range(1, 16):  # A-O covers both Receipts (A-F) and Payments (I-O)
                month_ws.cell(row=row, column=col).value = None

    ws = wb['Jan2026']
    # Receipts row: SNO(A) Date(B) Receipts(C) Amount(D) Category(E) Payment Mode(F)
    ws['A3'] = 1
    ws['D3'] = 1000
    ws['E3'] = 'Sales'
    ws['F3'] = 'Bank'
    # Payments row: Date(I) Payments(J) Amount(K) Main Category(L) Sub-Category(M) Payment Mode(N)
    ws['K3'] = 400
    ws['L3'] = 'Materials'
    ws['M3'] = 'Cement'
    ws['N3'] = 'Cash'
    wb.save(dest)
    return dest
