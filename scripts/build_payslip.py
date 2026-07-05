import sys

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, Border, Side

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from build_dashboard import MONTH_SHEETS

WORKERS_HEADER_ROW = 2
WORKERS_FIRST_ROW = 3
WORKERS_LAST_ROW = 52  # 50-worker capacity

PAYENTRIES_HEADER_ROW = 2
PAYENTRIES_FIRST_ROW = 3
PAYENTRIES_LAST_ROW = 502  # 500-entry capacity

PAYSLIP_SHEET_NAME = 'Payslip'
PAYSLIP_WORKER_CELL = 'B3'
PAYSLIP_MONTH_CELL = 'B4'
PAYSLIP_WORKER_ID_CELL = 'B6'
PAYSLIP_ROLE_CELL = 'B7'
PAYSLIP_RATE_TYPE_CELL = 'B8'
PAYSLIP_RATE_CELL = 'B9'
PAYSLIP_MESSAGE_CELL = 'B12'
PAYSLIP_DAYS_WORKED_CELL = 'B13'
PAYSLIP_GROSS_PAY_CELL = 'B14'
PAYSLIP_DEDUCTION_CELL = 'B15'
PAYSLIP_NET_PAY_CELL = 'B16'
PAYSLIP_PAYMENT_MODE_CELL = 'B17'
PAYSLIP_MATCH_ROW_CELL = 'Z1'  # hidden helper


def _add_workers_sheet(ws):
    ws.cell(row=1, column=1, value='Workers')
    headers = ['Worker ID', 'Name', 'Role', 'Rate Type', 'Rate or Agreed Amount',
               'Default Payment Mode', 'Active', 'Active Name']
    for i, header in enumerate(headers, start=1):
        ws.cell(row=WORKERS_HEADER_ROW, column=i, value=header)

    for row in range(WORKERS_FIRST_ROW, WORKERS_LAST_ROW + 1):
        ws.cell(row=row, column=8, value='=IF(G%d="Y",B%d,"")' % (row, row))


def _add_payentries_sheet(ws):
    ws.cell(row=1, column=1, value='PayEntries')
    headers = ['Worker ID', 'Worker Name', 'Month', 'Days Worked', 'Gross Pay',
               'Advance/Deduction', 'Net Pay', 'Payment Mode', 'Payments Row Ref',
               'Comments', 'Match Key']
    for i, header in enumerate(headers, start=1):
        ws.cell(row=PAYENTRIES_HEADER_ROW, column=i, value=header)

    for row in range(PAYENTRIES_FIRST_ROW, PAYENTRIES_LAST_ROW + 1):
        worker_id_cell = 'A%d' % row
        rate_type_lookup = 'INDEX(Workers!$D$%d:$D$%d,MATCH(%s,Workers!$A$%d:$A$%d,0))' % (
            WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_id_cell, WORKERS_FIRST_ROW, WORKERS_LAST_ROW)
        rate_lookup = 'INDEX(Workers!$E$%d:$E$%d,MATCH(%s,Workers!$A$%d:$A$%d,0))' % (
            WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_id_cell, WORKERS_FIRST_ROW, WORKERS_LAST_ROW)
        name_lookup = 'INDEX(Workers!$B$%d:$B$%d,MATCH(%s,Workers!$A$%d:$A$%d,0))' % (
            WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_id_cell, WORKERS_FIRST_ROW, WORKERS_LAST_ROW)
        default_mode_lookup = 'INDEX(Workers!$F$%d:$F$%d,MATCH(%s,Workers!$A$%d:$A$%d,0))' % (
            WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_id_cell, WORKERS_FIRST_ROW, WORKERS_LAST_ROW)

        ws.cell(row=row, column=2, value='=IF(%s="","",IFERROR(%s,"Invalid Worker ID"))' % (worker_id_cell, name_lookup))
        ws.cell(row=row, column=5, value=(
            '=IF(%s="","",IFERROR(IF(%s="Per Day",D%d*%s,%s),"Invalid Worker ID"))'
            % (worker_id_cell, rate_type_lookup, row, rate_lookup, rate_lookup)
        ))
        ws.cell(row=row, column=7, value='=IF(OR(%s="",E%d="Invalid Worker ID"),"",E%d-F%d)' % (worker_id_cell, row, row, row))
        ws.cell(row=row, column=8, value='=IF(%s="","",IFERROR(%s,"Invalid Worker ID"))' % (worker_id_cell, default_mode_lookup))
        ws.cell(row=row, column=11, value='=A%d&"|"&C%d' % (row, row))

    ws.column_dimensions['K'].hidden = True


def _add_payslip_sheet(ws):
    ws.cell(row=1, column=1, value='Payslip')

    ws['A3'] = 'Worker:'
    ws['A4'] = 'Month:'
    ws['A6'] = 'Worker ID:'
    ws['A7'] = 'Role:'
    ws['A8'] = 'Rate Type:'
    ws['A9'] = 'Rate / Agreed Amount:'
    ws['A13'] = 'Days Worked:'
    ws['A14'] = 'Gross Pay:'
    ws['A15'] = 'Advance/Deduction:'
    ws['A16'] = 'Net Pay:'
    ws['A17'] = 'Payment Mode:'
    ws['A18'] = 'Generated On:'

    worker_range = 'Workers!$B$%d:$B$%d' % (WORKERS_FIRST_ROW, WORKERS_LAST_ROW)
    worker_id_range = 'Workers!$A$%d:$A$%d' % (WORKERS_FIRST_ROW, WORKERS_LAST_ROW)

    ws[PAYSLIP_WORKER_ID_CELL] = '=IF($B$3="","",INDEX(%s,MATCH($B$3,%s,0)))' % (worker_id_range, worker_range)
    ws[PAYSLIP_ROLE_CELL] = '=IF($B$3="","",INDEX(Workers!$C$%d:$C$%d,MATCH($B$3,%s,0)))' % (
        WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_range)
    ws[PAYSLIP_RATE_TYPE_CELL] = '=IF($B$3="","",INDEX(Workers!$D$%d:$D$%d,MATCH($B$3,%s,0)))' % (
        WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_range)
    ws[PAYSLIP_RATE_CELL] = '=IF($B$3="","",INDEX(Workers!$E$%d:$E$%d,MATCH($B$3,%s,0)))' % (
        WORKERS_FIRST_ROW, WORKERS_LAST_ROW, worker_range)

    key_range = 'PayEntries!$K$%d:$K$%d' % (PAYENTRIES_FIRST_ROW, PAYENTRIES_LAST_ROW)
    ws[PAYSLIP_MATCH_ROW_CELL] = '=IFERROR(MATCH(%s&"|"&%s,%s,0),0)' % (
        PAYSLIP_WORKER_ID_CELL, PAYSLIP_MONTH_CELL, key_range)
    ws.column_dimensions['Z'].hidden = True

    ws[PAYSLIP_MESSAGE_CELL] = '=IF(%s=0,"No pay entry found for this period","")' % PAYSLIP_MATCH_ROW_CELL

    def _payentries_lookup(col_letter):
        rng = 'PayEntries!$%s$%d:$%s$%d' % (col_letter, PAYENTRIES_FIRST_ROW, col_letter, PAYENTRIES_LAST_ROW)
        return '=IF(%s=0,"",INDEX(%s,%s))' % (PAYSLIP_MATCH_ROW_CELL, rng, PAYSLIP_MATCH_ROW_CELL)

    ws[PAYSLIP_DAYS_WORKED_CELL] = _payentries_lookup('D')
    ws[PAYSLIP_GROSS_PAY_CELL] = _payentries_lookup('E')
    ws[PAYSLIP_DEDUCTION_CELL] = _payentries_lookup('F')
    ws[PAYSLIP_NET_PAY_CELL] = _payentries_lookup('G')
    ws[PAYSLIP_PAYMENT_MODE_CELL] = _payentries_lookup('H')
    ws['B18'] = '=TODAY()'


def _add_payslip_dropdowns(ws):
    active_name_range = "'Workers'!$H$%d:$H$%d" % (WORKERS_FIRST_ROW, WORKERS_LAST_ROW)
    dv_worker = DataValidation(type='list', formula1=active_name_range, allow_blank=True)
    ws.add_data_validation(dv_worker)
    dv_worker.add(PAYSLIP_WORKER_CELL)

    month_list = '"' + ','.join(MONTH_SHEETS) + '"'
    dv_month = DataValidation(type='list', formula1=month_list, allow_blank=True)
    ws.add_data_validation(dv_month)
    dv_month.add(PAYSLIP_MONTH_CELL)


def _style_payslip(ws):
    title_font = Font(bold=True, size=14)
    label_font = Font(bold=True)
    value_font = Font(size=12)
    thin = Side(style='thin', color='B0B0B0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws['A1'].font = title_font

    for row in [3, 4, 6, 7, 8, 9, 13, 14, 15, 16, 17, 18]:
        label_cell = ws.cell(row=row, column=1)
        label_cell.font = label_font
        value_cell = ws.cell(row=row, column=2)
        value_cell.font = value_font
        value_cell.border = border

    ws[PAYSLIP_NET_PAY_CELL].font = Font(bold=True, size=13)

    for col_letter, width in [('A', 22), ('B', 30)]:
        ws.column_dimensions[col_letter].width = width

    ws.print_area = 'A1:B18'
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build(path):
    wb = openpyxl.load_workbook(path, keep_vba=True)

    if 'Workers' in wb.sheetnames:
        del wb['Workers']
    workers_ws = wb.create_sheet('Workers')
    _add_workers_sheet(workers_ws)

    if 'PayEntries' in wb.sheetnames:
        del wb['PayEntries']
    payentries_ws = wb.create_sheet('PayEntries')
    _add_payentries_sheet(payentries_ws)

    if 'Payslip' in wb.sheetnames:
        del wb['Payslip']
    payslip_ws = wb.create_sheet(PAYSLIP_SHEET_NAME)
    _add_payslip_sheet(payslip_ws)
    _add_payslip_dropdowns(payslip_ws)
    _style_payslip(payslip_ws)

    wb.save(path)


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else 'expense_tracker_2026.xlsm')
