# scripts/test_build_dashboard.py
import os
import sys
import tempfile

import formulas

sys.path.insert(0, os.path.dirname(__file__))
import build_dashboard
from test_helpers import sample_workbook

REAL_WORKBOOK = os.path.join(os.path.dirname(__file__), '..', 'expense_tracker_2026.xlsx')


def calc(path):
    xl = formulas.ExcelModel().loads(path).finish()
    return xl.calculate()


def cell_value(sol, path, sheet, cell):
    key = "'[%s]%s'!%s" % (os.path.basename(path), sheet.upper(), cell)
    v = sol[key].value
    try:
        return v[0][0]
    except (TypeError, IndexError):
        return v


def test_read_categories():
    import openpyxl
    wb = openpyxl.load_workbook(REAL_WORKBOOK)
    income, main, main_sub = build_dashboard.read_categories(wb)
    assert income == ['Investor Capital', 'Partner Contribution', 'Loan Received', 'Sales']
    assert len(main) == 10
    assert sum(len(v) for v in main_sub.values()) == 68
    print("test_read_categories PASSED")


def test_monthly_trend_table():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_dashboard.build(path)
        sol = calc(path)
        # Jan2026 row is MONTHLY_TREND_FIRST_ROW (45): Receipts=1000, Payments=400, Closing=600
        row = build_dashboard.MONTHLY_TREND_FIRST_ROW
        assert cell_value(sol, path, 'Dashboard', 'B%d' % row) == 1000
        assert cell_value(sol, path, 'Dashboard', 'C%d' % row) == 400
        assert cell_value(sol, path, 'Dashboard', 'D%d' % row) == 600
        print("test_monthly_trend_table PASSED")


def test_kpi_cards():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_dashboard.build(path)
        sol = calc(path)
        assert cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['receipts']) == 1000
        assert cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['payments']) == 400
        assert cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['net']) == 600
        # Dec2026 closing balance carries Jan's 600 forward through empty months
        assert cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['balance']) == 600
        print("test_kpi_cards PASSED")


def test_category_tables():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_dashboard.build(path)
        sol = calc(path)

        import openpyxl
        wb = openpyxl.load_workbook(path)
        dash = wb['Dashboard']

        # Main Category spend: find the "Materials" row, assert amount == 400
        materials_row = None
        for r in range(build_dashboard.MAIN_CAT_FIRST_ROW, build_dashboard.MAIN_CAT_LAST_ROW + 1):
            if dash.cell(row=r, column=6).value == 'Materials':
                materials_row = r
        assert materials_row is not None
        assert cell_value(sol, path, 'Dashboard', 'G%d' % materials_row) == 400

        # Income by Category: find the "Sales" row, assert amount == 1000
        sales_row = None
        for r in range(build_dashboard.INCOME_CAT_FIRST_ROW, build_dashboard.INCOME_CAT_LAST_ROW + 1):
            if dash.cell(row=r, column=9).value == 'Sales':
                sales_row = r
        assert sales_row is not None
        assert cell_value(sol, path, 'Dashboard', 'J%d' % sales_row) == 1000

        # Bank vs Cash: Bank total == 1000 (the receipt), Cash total == 400 (the payment)
        bank_row, cash_row = build_dashboard.BANKCASH_FIRST_ROW, build_dashboard.BANKCASH_FIRST_ROW + 1
        assert dash.cell(row=bank_row, column=12).value == 'Bank'
        assert cell_value(sol, path, 'Dashboard', 'M%d' % bank_row) == 1000
        assert dash.cell(row=cash_row, column=12).value == 'Cash'
        assert cell_value(sol, path, 'Dashboard', 'M%d' % cash_row) == 400

        print("test_category_tables PASSED")


if __name__ == '__main__':
    test_read_categories()
    test_monthly_trend_table()
    test_kpi_cards()
    test_category_tables()
