import sys

import openpyxl

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from build_dashboard import MONTH_SHEETS

WORKERS_HEADER_ROW = 2
WORKERS_FIRST_ROW = 3
WORKERS_LAST_ROW = 52  # 50-worker capacity


def _add_workers_sheet(ws):
    ws.cell(row=1, column=1, value='Workers')
    headers = ['Worker ID', 'Name', 'Role', 'Rate Type', 'Rate or Agreed Amount',
               'Default Payment Mode', 'Active', 'Active Name']
    for i, header in enumerate(headers, start=1):
        ws.cell(row=WORKERS_HEADER_ROW, column=i, value=header)

    for row in range(WORKERS_FIRST_ROW, WORKERS_LAST_ROW + 1):
        ws.cell(row=row, column=8, value='=IF(G%d="Y",B%d,"")' % (row, row))


def build(path):
    wb = openpyxl.load_workbook(path, keep_vba=True)

    if 'Workers' in wb.sheetnames:
        del wb['Workers']
    workers_ws = wb.create_sheet('Workers')
    _add_workers_sheet(workers_ws)

    wb.save(path)


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else 'expense_tracker_2026.xlsm')
