import sys

import openpyxl
from openpyxl.utils import get_column_letter

MONTH_SHEETS = [
    'Jan2026', 'Feb2026', 'Mar2026', 'Apr2026', 'May2026', 'Jun2026',
    'Jul2026', 'Aug2026', 'Sep2026', 'Oct2026', 'Nov2026', 'Dec2026',
]

MONTHLY_TREND_HEADER_ROW = 44
MONTHLY_TREND_FIRST_ROW = 45
MONTHLY_TREND_LAST_ROW = MONTHLY_TREND_FIRST_ROW + len(MONTH_SHEETS) - 1  # 56

KPI_LABEL_ROW = 3
KPI_VALUE_ROW = 4
KPI_CELLS = {'balance': 'A4', 'receipts': 'C4', 'payments': 'E4', 'net': 'G4'}


def read_categories(wb):
    """Read Income Categories, Main Categories, and the Main->Sub-category
    map directly from the existing Categories sheet."""
    cat = wb['Categories']

    income_categories = []
    r = 3
    while cat.cell(row=r, column=1).value:
        income_categories.append(cat.cell(row=r, column=1).value)
        r += 1

    main_categories = []
    r = 3
    while cat.cell(row=r, column=6).value:
        main_categories.append(cat.cell(row=r, column=6).value)
        r += 1

    main_sub = {}
    col = 8
    while cat.cell(row=2, column=col).value:
        main = cat.cell(row=2, column=col).value
        subs = []
        r = 3
        while cat.cell(row=r, column=col).value:
            subs.append(cat.cell(row=r, column=col).value)
            r += 1
        main_sub[main] = subs
        col += 1

    return income_categories, main_categories, main_sub


def _add_monthly_trend_table(ws):
    ws.cell(row=MONTHLY_TREND_HEADER_ROW, column=1, value='Month')
    ws.cell(row=MONTHLY_TREND_HEADER_ROW, column=2, value='Receipts')
    ws.cell(row=MONTHLY_TREND_HEADER_ROW, column=3, value='Payments')
    ws.cell(row=MONTHLY_TREND_HEADER_ROW, column=4, value='Closing Balance')

    for i, month in enumerate(MONTH_SHEETS):
        row = MONTHLY_TREND_FIRST_ROW + i
        ws.cell(row=row, column=1, value=month)
        ws.cell(row=row, column=2, value="='%s'!B104" % month)
        ws.cell(row=row, column=3, value="='%s'!B105" % month)
        ws.cell(row=row, column=4, value="='%s'!B106" % month)


def _add_kpi_cards(ws):
    trend_receipts_range = "B%d:B%d" % (MONTHLY_TREND_FIRST_ROW, MONTHLY_TREND_LAST_ROW)
    trend_payments_range = "C%d:C%d" % (MONTHLY_TREND_FIRST_ROW, MONTHLY_TREND_LAST_ROW)

    ws['A3'] = 'Current Cash Balance'
    ws['A4'] = "='%s'!B106" % MONTH_SHEETS[-1]

    ws['C3'] = 'YTD Total Receipts'
    ws['C4'] = "=SUM(%s)" % trend_receipts_range

    ws['E3'] = 'YTD Total Payments'
    ws['E4'] = "=SUM(%s)" % trend_payments_range

    ws['G3'] = 'Net Position (YTD)'
    ws['G4'] = "=%s-%s" % (KPI_CELLS['receipts'], KPI_CELLS['payments'])


def build(path):
    wb = openpyxl.load_workbook(path)
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']
    ws = wb.create_sheet('Dashboard', 0)

    ws.cell(row=1, column=1, value='Site Expense Dashboard')
    _add_monthly_trend_table(ws)
    _add_kpi_cards(ws)

    wb.save(path)


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else 'expense_tracker_2026.xlsx')
