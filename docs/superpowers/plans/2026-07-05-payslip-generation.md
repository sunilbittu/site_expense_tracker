# Payslip Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Workers, PayEntries, and Payslip sheets to `expense_tracker_2026.xlsm`, plus a VBA macro that writes a payslip's net pay into the target month's Payments table and exports the Payslip sheet to PDF, per `docs/superpowers/specs/2026-07-05-payslip-generation-design.md`.

**Architecture:** A single script, `scripts/build_payslip.py`, exposes a `build(path)` function that (re)creates the Workers, PayEntries, and Payslip sheets in the workbook at `path`. Workers is a plain roster; PayEntries is a formula-driven pay-period log keyed by Worker ID + Month; Payslip is a pure lookup/display template with no stored data of its own. A separate `.bas` file holds the VBA macro source — same pattern as the existing `scripts/ExportDashboardToPDF.bas` — since VBA cannot be authored programmatically from this environment (no Excel/COM automation on macOS) and must be pasted into the VBA editor by hand.

**Tech Stack:** Python 3, `openpyxl` (build workbook structure/formulas), `formulas` (evaluate formulas in tests to assert real computed values). VBA (hand-pasted, not test-automatable).

## Global Constraints

- Workbook under edit: `expense_tracker_2026.xlsm` (repo root). **Always** load with `openpyxl.load_workbook(path, keep_vba=True)` and never with the default `keep_vba=False` — this workbook carries a real VBA macro (`ExportDashboardToPDF`) that must survive every edit.
- `MONTH_SHEETS = ['Jan2026','Feb2026','Mar2026','Apr2026','May2026','Jun2026','Jul2026','Aug2026','Sep2026','Oct2026','Nov2026','Dec2026']` — already defined in `scripts/build_dashboard.py`; import it, do not redefine it.
- Existing Payments table layout on each month sheet (already built, do not change): columns I–O = `Date, Payments, Amount, Main Category, Sub-Category, Payment Mode, Comments`, data rows 3–100. Amount = column K, Main Category = column L, Sub-Category = column M, Payment Mode = column N, Comments = column O.
- Existing category data (already built, do not change): `Labour & Contractors` is a Main Category with sub-categories including `Labour Wages` and `Contractor Payments` (confirmed present in the Categories sheet's sub-category grid).
- Test fixture: use `scripts/test_helpers.sample_workbook(tmp_dir)` for any test that needs a workbook copy — it already copies the real `.xlsm`, preserves the macro, and clears pre-existing demo data from every month sheet before injecting its own fixture row. Do not write a second/competing fixture helper.
- No new dependencies beyond `openpyxl` and `formulas` (already in `requirements.txt`).
- The `formulas` library evaluates `.xlsm` files correctly (confirmed) — use the same `calc()`/`cell_value()` test helpers already defined in `scripts/test_build_dashboard.py`'s pattern (re-implement identically in the new test file, since test files in this project are plain scripts, not an importable shared module).

---

### Task 1: Script scaffold and Workers sheet

**Files:**
- Create: `scripts/build_payslip.py`
- Create: `scripts/test_build_payslip.py`

**Interfaces:**
- Produces: `build_payslip.WORKERS_HEADER_ROW = 2`, `WORKERS_FIRST_ROW = 3`, `WORKERS_LAST_ROW = 52` (50-worker capacity), columns: `A` Worker ID, `B` Name, `C` Role, `D` Rate Type, `E` Rate or Agreed Amount, `F` Default Payment Mode, `G` Active, `H` Active Name (hidden helper: `=IF(G{row}="Y",B{row},"")`, used by Task 3's Payslip Worker dropdown so only Active workers are selectable).
- Produces: `build_payslip.build(path: str) -> None` — top-level entry point, idempotently (re)builds the `Workers`, `PayEntries`, and `Payslip` sheets in the workbook at `path` and saves it. This task only implements the Workers sheet; later tasks add to the same function.
- Consumes: `build_dashboard.MONTH_SHEETS` (import from `scripts/build_dashboard.py`).

- [ ] **Step 1: Write the failing test**

```python
# scripts/test_build_payslip.py
import os
import sys
import tempfile

import formulas

sys.path.insert(0, os.path.dirname(__file__))
import build_payslip
from test_helpers import sample_workbook

REAL_WORKBOOK = os.path.join(os.path.dirname(__file__), '..', 'expense_tracker_2026.xlsm')


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


def test_workers_sheet_structure_and_active_name_formula():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_payslip.build(path)

        import openpyxl
        wb = openpyxl.load_workbook(path)
        workers = wb['Workers']
        assert workers.cell(row=build_payslip.WORKERS_HEADER_ROW, column=1).value == 'Worker ID'
        assert workers.cell(row=build_payslip.WORKERS_HEADER_ROW, column=8).value == 'Active Name'

        # Row 3: an Active labourer -> Active Name should mirror Name
        workers.cell(row=3, column=1, value='W001')
        workers.cell(row=3, column=2, value='Ramesh')
        workers.cell(row=3, column=3, value='Labourer')
        workers.cell(row=3, column=4, value='Per Day')
        workers.cell(row=3, column=5, value=500)
        workers.cell(row=3, column=6, value='Cash')
        workers.cell(row=3, column=7, value='Y')
        # Row 4: an Inactive worker -> Active Name should be blank
        workers.cell(row=4, column=1, value='W002')
        workers.cell(row=4, column=2, value='Suresh')
        workers.cell(row=4, column=3, value='Labourer')
        workers.cell(row=4, column=4, value='Per Day')
        workers.cell(row=4, column=5, value=500)
        workers.cell(row=4, column=6, value='Cash')
        workers.cell(row=4, column=7, value='N')
        wb.save(path)

        sol = calc(path)
        assert cell_value(sol, path, 'Workers', 'H3') == 'Ramesh'
        assert cell_value(sol, path, 'Workers', 'H4') == ''
        print("test_workers_sheet_structure_and_active_name_formula PASSED")


if __name__ == '__main__':
    test_workers_sheet_structure_and_active_name_formula()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_payslip.py`
Expected: `ModuleNotFoundError: No module named 'build_payslip'`

- [ ] **Step 3: Write `scripts/build_payslip.py` (scaffold + Workers sheet)**

```python
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_payslip.py`
Expected: `test_workers_sheet_structure_and_active_name_formula PASSED`

- [ ] **Step 5: Commit**

```bash
git add scripts/build_payslip.py scripts/test_build_payslip.py
git commit -m "Add payslip script scaffold and Workers roster sheet"
```

---

### Task 2: PayEntries sheet

**Files:**
- Modify: `scripts/build_payslip.py`
- Modify: `scripts/test_build_payslip.py`

**Interfaces:**
- Consumes: `WORKERS_FIRST_ROW`, `WORKERS_LAST_ROW` from Task 1 (Workers sheet columns A Worker ID, B Name, C Role, D Rate Type, E Rate/Agreed Amount, F Default Payment Mode).
- Produces: `PAYENTRIES_HEADER_ROW = 2`, `PAYENTRIES_FIRST_ROW = 3`, `PAYENTRIES_LAST_ROW = 502` (500-entry capacity), columns: `A` Worker ID, `B` Worker Name, `C` Month, `D` Days Worked, `E` Gross Pay, `F` Advance/Deduction, `G` Net Pay, `H` Payment Mode, `I` Payments Row Ref, `J` Comments, `K` Match Key (hidden helper: `=A{row}&"|"&C{row}`, used by Task 3's Payslip sheet to look up a PayEntries row by Worker ID + Month in one `MATCH`).
- Produces: `_add_payentries_sheet(ws)` wired into `build()`.

- [ ] **Step 1: Write the failing test**

Add to `scripts/test_build_payslip.py` (above the `if __name__ == '__main__':` block):

```python
def test_payentries_formulas():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_payslip.build(path)

        import openpyxl
        wb = openpyxl.load_workbook(path)
        workers = wb['Workers']
        workers.cell(row=3, column=1, value='W001')
        workers.cell(row=3, column=2, value='Ramesh')
        workers.cell(row=3, column=3, value='Labourer')
        workers.cell(row=3, column=4, value='Per Day')
        workers.cell(row=3, column=5, value=500)
        workers.cell(row=3, column=6, value='Cash')
        workers.cell(row=3, column=7, value='Y')

        workers.cell(row=4, column=1, value='W010')
        workers.cell(row=4, column=2, value='ACME Contractors')
        workers.cell(row=4, column=3, value='Contractor')
        workers.cell(row=4, column=4, value='Fixed Amount')
        workers.cell(row=4, column=5, value=50000)
        workers.cell(row=4, column=6, value='Bank')
        workers.cell(row=4, column=7, value='Y')

        entries = wb['PayEntries']
        # Labourer: 20 days x 500/day = 10000 gross, 1000 advance -> 9000 net
        entries.cell(row=3, column=1, value='W001')
        entries.cell(row=3, column=2, value='Ramesh')
        entries.cell(row=3, column=3, value='Jan2026')
        entries.cell(row=3, column=4, value=20)
        entries.cell(row=3, column=6, value=1000)
        entries.cell(row=3, column=8, value='Cash')

        # Contractor: fixed 50000 gross, no deduction -> 50000 net
        entries.cell(row=4, column=1, value='W010')
        entries.cell(row=4, column=2, value='ACME Contractors')
        entries.cell(row=4, column=3, value='Jan2026')
        entries.cell(row=4, column=6, value=0)
        entries.cell(row=4, column=8, value='Bank')
        wb.save(path)

        sol = calc(path)
        assert cell_value(sol, path, 'PayEntries', 'E3') == 10000
        assert cell_value(sol, path, 'PayEntries', 'G3') == 9000
        assert cell_value(sol, path, 'PayEntries', 'E4') == 50000
        assert cell_value(sol, path, 'PayEntries', 'G4') == 50000
        assert cell_value(sol, path, 'PayEntries', 'K3') == 'W001|Jan2026'
        print("test_payentries_formulas PASSED")
```

Update `__main__` to also call `test_payentries_formulas()`.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_payslip.py`
Expected: `KeyError` or similar — `PayEntries` sheet does not exist yet.

- [ ] **Step 3: Implement the PayEntries sheet**

Add module-level constants (after `WORKERS_LAST_ROW`):

```python
PAYENTRIES_HEADER_ROW = 2
PAYENTRIES_FIRST_ROW = 3
PAYENTRIES_LAST_ROW = 502  # 500-entry capacity
```

Add the function:

```python
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
```

Note: column 8 (Payment Mode) is written as a formula default here, but nothing prevents the user from typing over it in Excel with a static value for a specific row — that satisfies the spec's "defaults from Workers, overridable per period."

Note: every `Workers` lookup is wrapped in `IFERROR(...,"Invalid Worker ID")` — without a dropdown constraining PayEntries' Worker ID column, a typo would otherwise cascade a raw `#N/A` into Gross Pay and Net Pay silently. Net Pay (column G) additionally checks for that literal string so it shows blank rather than a `#VALUE!` from subtracting text.

Wire it into `build()`:

```python
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

    wb.save(path)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_payslip.py`
Expected: both tests print `PASSED`.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_payslip.py scripts/test_build_payslip.py
git commit -m "Add PayEntries sheet with Gross/Net Pay formulas"
```

---

### Task 3: Payslip template sheet

**Files:**
- Modify: `scripts/build_payslip.py`
- Modify: `scripts/test_build_payslip.py`

**Interfaces:**
- Consumes: Workers sheet columns A–C (Task 1), PayEntries sheet columns D, E, F, G, H, K (Task 2).
- Produces: `PAYSLIP_WORKER_CELL = 'B3'`, `PAYSLIP_MONTH_CELL = 'B4'` (the two dropdown input cells), `PAYSLIP_WORKER_ID_CELL = 'B6'`, `PAYSLIP_NET_PAY_CELL = 'B16'`, `PAYSLIP_PAYMENT_MODE_CELL = 'B17'` — the VBA macro in Task 4 reads these exact cells.
- Produces: `_add_payslip_sheet(ws)` wired into `build()`.

- [ ] **Step 1: Write the failing test**

Add to `scripts/test_build_payslip.py`:

```python
def test_payslip_lookup_and_no_entry_message():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_payslip.build(path)

        import openpyxl
        wb = openpyxl.load_workbook(path)
        workers = wb['Workers']
        workers.cell(row=3, column=1, value='W001')
        workers.cell(row=3, column=2, value='Ramesh')
        workers.cell(row=3, column=3, value='Labourer')
        workers.cell(row=3, column=4, value='Per Day')
        workers.cell(row=3, column=5, value=500)
        workers.cell(row=3, column=6, value='Cash')
        workers.cell(row=3, column=7, value='Y')

        entries = wb['PayEntries']
        entries.cell(row=3, column=1, value='W001')
        entries.cell(row=3, column=2, value='Ramesh')
        entries.cell(row=3, column=3, value='Jan2026')
        entries.cell(row=3, column=4, value=20)
        entries.cell(row=3, column=6, value=1000)
        entries.cell(row=3, column=8, value='Cash')

        payslip = wb[build_payslip.PAYSLIP_SHEET_NAME]
        payslip[build_payslip.PAYSLIP_WORKER_CELL] = 'Ramesh'
        payslip[build_payslip.PAYSLIP_MONTH_CELL] = 'Jan2026'
        wb.save(path)

        sol = calc(path)
        assert cell_value(sol, path, 'Payslip', build_payslip.PAYSLIP_NET_PAY_CELL) == 9000
        assert cell_value(sol, path, 'Payslip', build_payslip.PAYSLIP_PAYMENT_MODE_CELL) == 'Cash'

        # No PayEntries row for Feb2026 -> should show the "no entry" message, not an error
        wb2 = openpyxl.load_workbook(path)
        wb2[build_payslip.PAYSLIP_SHEET_NAME][build_payslip.PAYSLIP_MONTH_CELL] = 'Feb2026'
        wb2.save(path)
        sol2 = calc(path)
        assert cell_value(sol2, path, 'Payslip', 'B12') == 'No pay entry found for this period'
        assert cell_value(sol2, path, 'Payslip', build_payslip.PAYSLIP_NET_PAY_CELL) == ''
        print("test_payslip_lookup_and_no_entry_message PASSED")
```

Update `__main__` to also call `test_payslip_lookup_and_no_entry_message()`.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_payslip.py`
Expected: `AttributeError: module 'build_payslip' has no attribute 'PAYSLIP_SHEET_NAME'`

- [ ] **Step 3: Implement the Payslip sheet**

Add module-level constants (after `PAYENTRIES_LAST_ROW`):

```python
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
```

Add the function:

```python
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
```

Add the Worker/Month dropdowns (needs `from openpyxl.worksheet.datavalidation import DataValidation` added to the imports at the top of the file):

```python
def _add_payslip_dropdowns(ws):
    active_name_range = "='Workers'!$H$%d:$H$%d" % (WORKERS_FIRST_ROW, WORKERS_LAST_ROW)
    dv_worker = DataValidation(type='list', formula1=active_name_range, allow_blank=True)
    ws.add_data_validation(dv_worker)
    dv_worker.add(PAYSLIP_WORKER_CELL)

    month_list = '"' + ','.join(MONTH_SHEETS) + '"'
    dv_month = DataValidation(type='list', formula1=month_list, allow_blank=True)
    ws.add_data_validation(dv_month)
    dv_month.add(PAYSLIP_MONTH_CELL)
```

Wire both into `build()`:

```python
    if 'Payslip' in wb.sheetnames:
        del wb['Payslip']
    payslip_ws = wb.create_sheet(PAYSLIP_SHEET_NAME)
    _add_payslip_sheet(payslip_ws)
    _add_payslip_dropdowns(payslip_ws)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_payslip.py`
Expected: all three tests print `PASSED`.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_payslip.py scripts/test_build_payslip.py
git commit -m "Add Payslip lookup/display template sheet with Worker+Month dropdowns"
```

---

### Task 4: Styling, print area, and VBA macro source

**Files:**
- Modify: `scripts/build_payslip.py`
- Modify: `scripts/test_build_payslip.py`
- Create: `scripts/GeneratePayslip.bas`

**Interfaces:**
- Consumes: all cell constants from Tasks 1–3 (`PAYSLIP_*`, `WORKERS_*`, `PAYENTRIES_*`).
- Produces: nothing further downstream — this is the final scripted task. The `.bas` file is a manual-install deliverable, not something `build()` touches.

- [ ] **Step 1: Write the failing test — styling must not change any formula/value**

Add to `scripts/test_build_payslip.py`:

```python
def test_styling_preserves_formulas():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_payslip.build(path)

        import openpyxl
        wb = openpyxl.load_workbook(path)
        payslip = wb[build_payslip.PAYSLIP_SHEET_NAME]

        assert payslip[build_payslip.PAYSLIP_NET_PAY_CELL].value.startswith('=IF(')
        assert payslip.print_area is not None
        print("test_styling_preserves_formulas PASSED")
```

Update `__main__` to also call `test_styling_preserves_formulas()`.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_payslip.py`
Expected: `AssertionError` (`payslip.print_area is None` — no print area set yet).

- [ ] **Step 3: Add styling and print area**

Add to the imports at the top of `scripts/build_payslip.py`:

```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
```

Add the function, reusing the workbook's existing palette (slate `37474F` for headers, matching the Dashboard and Categories sheets):

```python
def _style_payslip(ws):
    title_font = Font(bold=True, size=14)
    label_font = Font(bold=True)
    value_font = Font(size=12)
    header_fill = PatternFill(start_color='37474F', end_color='37474F', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')
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

    ws.column_dimensions['Z'].hidden = True
```

Wire it into `build()`, right after `_add_payslip_dropdowns(payslip_ws)`:

```python
    _add_payslip_dropdowns(payslip_ws)
    _style_payslip(payslip_ws)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_payslip.py`
Expected: all four tests print `PASSED`.

- [ ] **Step 5: Write the VBA macro source**

Create `scripts/GeneratePayslip.bas` (no `Attribute VB_Name` line — that syntax is only valid on `File > Import`, not on copy-paste into the VBA editor, per the lesson learned building `ExportDashboardToPDF`):

```vba
Sub GeneratePayslip()
    Dim payslipWs As Worksheet
    Dim payEntriesWs As Worksheet
    Dim targetWs As Worksheet
    Dim workerId As String
    Dim monthName As String
    Dim netPay As Double
    Dim mainCategory As String
    Dim subCategory As String
    Dim paymentMode As String
    Dim role As String
    Dim entryRow As Long
    Dim payRowRefCell As Range
    Dim targetRow As Long
    Dim r As Long
    Dim exportPath As String
    Dim fileName As String

    Set payslipWs = ThisWorkbook.Sheets("Payslip")
    Set payEntriesWs = ThisWorkbook.Sheets("PayEntries")

    workerId = payslipWs.Range("B6").Value
    monthName = payslipWs.Range("B4").Value
    role = payslipWs.Range("B7").Value
    netPay = payslipWs.Range("B16").Value
    paymentMode = payslipWs.Range("B17").Value

    If workerId = "" Or monthName = "" Or payslipWs.Range("B16").Value = "" Then
        MsgBox "No pay entry found for this Worker/Month. Fill in PayEntries first.", vbExclamation
        Exit Sub
    End If

    ' Find the matching PayEntries row (Worker ID + Month) to read/write Payments Row Ref
    entryRow = 0
    For r = 3 To 502
        If payEntriesWs.Cells(r, 1).Value = workerId And payEntriesWs.Cells(r, 3).Value = monthName Then
            entryRow = r
            Exit For
        End If
    Next r

    If entryRow = 0 Then
        MsgBox "No matching PayEntries row found. Fill in PayEntries first.", vbExclamation
        Exit Sub
    End If

    If role = "Labourer" Then
        mainCategory = "Labour & Contractors"
        subCategory = "Labour Wages"
    Else
        mainCategory = "Labour & Contractors"
        subCategory = "Contractor Payments"
    End If

    Set targetWs = ThisWorkbook.Sheets(monthName)
    Set payRowRefCell = payEntriesWs.Cells(entryRow, 9) ' Payments Row Ref column

    If payRowRefCell.Value <> "" Then
        ' Regenerating: update the existing Payments row in place
        Dim refParts() As String
        refParts = Split(payRowRefCell.Value, "K")
        targetRow = CLng(refParts(1))
    Else
        ' First generation: find the first blank row in the Payments table (rows 3-100)
        targetRow = 0
        For r = 3 To 100
            If targetWs.Cells(r, 11).Value = "" Then ' column K = Amount
                targetRow = r
                Exit For
            End If
        Next r
        If targetRow = 0 Then
            MsgBox "Payments table for " & monthName & " is full (rows 3-100 all used).", vbCritical
            Exit Sub
        End If
    End If

    targetWs.Cells(targetRow, 11).Value = netPay              ' Amount (K)
    targetWs.Cells(targetRow, 12).Value = mainCategory         ' Main Category (L)
    targetWs.Cells(targetRow, 13).Value = subCategory          ' Sub-Category (M)
    targetWs.Cells(targetRow, 14).Value = paymentMode          ' Payment Mode (N)
    targetWs.Cells(targetRow, 15).Value = "Payslip: " & payslipWs.Range("B3").Value & " - " & monthName ' Comments (O)

    payRowRefCell.Value = monthName & "!K" & targetRow

    fileName = "Payslip_" & workerId & "_" & monthName & ".pdf"
    exportPath = ThisWorkbook.Path & Application.PathSeparator & fileName
    payslipWs.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Payslip generated and payment recorded in " & monthName & "!K" & targetRow & vbCrLf & _
           "PDF: " & exportPath, vbInformation
End Sub
```

- [ ] **Step 6: Commit**

```bash
git add scripts/build_payslip.py scripts/test_build_payslip.py scripts/GeneratePayslip.bas
git commit -m "Add Payslip styling, print area, and GeneratePayslip macro source"
```

---

### Task 5: Final integration against the real workbook

**Files:**
- Modify: `scripts/build_payslip.py`
- Modify: `scripts/test_build_payslip.py`
- Modify: `expense_tracker_2026.xlsm` (the actual output of running the finished script)

**Interfaces:**
- Consumes: everything from Tasks 1–4.
- Produces: nothing further downstream — this is the final task.

- [ ] **Step 1: Write the failing test — regression check against the real workbook**

Add to `scripts/test_build_payslip.py`:

```python
def test_builds_cleanly_against_real_workbook():
    with tempfile.TemporaryDirectory() as tmp:
        import shutil
        path = os.path.join(tmp, 'real_copy.xlsm')
        shutil.copyfile(REAL_WORKBOOK, path)
        build_payslip.build(path)

        import openpyxl
        wb = openpyxl.load_workbook(path, keep_vba=True)
        assert wb.vba_archive is not None, "macro was dropped by build_payslip.build()"
        assert 'Workers' in wb.sheetnames
        assert 'PayEntries' in wb.sheetnames
        assert 'Payslip' in wb.sheetnames

        sol = calc(path)
        # No workers/pay entries exist yet -> Payslip should show the "no entry" message cleanly
        assert cell_value(sol, path, 'Payslip', 'B12') in ('No pay entry found for this period', '')
        print("test_builds_cleanly_against_real_workbook PASSED")
```

Update `__main__` to also call `test_builds_cleanly_against_real_workbook()`.

- [ ] **Step 2: Run test to verify it passes**

Run: `python3 scripts/test_build_payslip.py`
Expected: all five tests print `PASSED`. If it fails, that indicates a bug from an earlier task; fix `build_payslip.py` until it passes before proceeding.

- [ ] **Step 3: Confirm the real workbook is not open in Excel**

```bash
ls -la expense_tracker_2026.xlsm 2>&1
ls -la "~\$expense_tracker_2026.xlsm" 2>&1
```

Expected: the second command errors with "No such file or directory" (no Excel lock file present). If it does exist, stop and ask the user to close the file in Excel first — writing to an open workbook risks data loss.

- [ ] **Step 4: Build into the real workbook**

```bash
python3 scripts/build_payslip.py expense_tracker_2026.xlsm
```

- [ ] **Step 5: Spot-check the real workbook**

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('expense_tracker_2026.xlsm', keep_vba=True)
print('vba_archive present:', wb.vba_archive is not None)
print('sheetnames:', wb.sheetnames)
print('Workers header:', wb['Workers']['A2'].value)
print('Payslip Net Pay formula:', wb['Payslip']['B16'].value)
"
```

Expected: `vba_archive present: True`; `Workers`, `PayEntries`, `Payslip` all present in `sheetnames`; `Workers header: Worker ID`; `Payslip Net Pay formula` shows an `=IF(...)` formula.

- [ ] **Step 6: Commit**

```bash
git add scripts/build_payslip.py scripts/test_build_payslip.py expense_tracker_2026.xlsm
git commit -m "Build Workers/PayEntries/Payslip sheets into the real workbook"
```

---

## Self-Review Notes

- **Spec coverage:** Workers roster with Active-only filtering (Task 1), PayEntries pay-period log with Gross/Net Pay formulas and the Match Key helper (Task 2), Payslip lookup template with dropdowns and the "no entry found" message (Task 3), styling/print area (Task 4), VBA macro source implementing both the Payments write-back and PDF export (Task 4), and final real-workbook integration (Task 5) — every section of the spec maps to a task.
- **Placeholder scan:** no TBD/TODO; every step has complete, runnable code (VBA included).
- **Type/name consistency:** `WORKERS_FIRST_ROW/LAST_ROW`, `PAYENTRIES_FIRST_ROW/LAST_ROW`, and every `PAYSLIP_*_CELL` constant are each defined exactly once (Tasks 1–3) and referenced identically by name in every later task, the VBA macro's hardcoded cell references (`B6`, `B4`, `B16`, `B17`, `B3`, `B7`), and the styling pass — no renames across tasks. The VBA macro cannot import Python constants, so its cell references (`Range("B6")`, `Range("B16")`, etc.) are manually kept in sync with `PAYSLIP_WORKER_ID_CELL`, `PAYSLIP_NET_PAY_CELL`, etc. — if a future change moves those cells, the macro must be updated by hand alongside it.
- **VBA is not unit-tested** — there is no way to execute VBA from this environment. Task 4/5 rely on the Python-buildable structure being correct (verified by tests) and the macro being manually pasted and manually verified by the user in Excel, same as `ExportDashboardToPDF` was.
- **Ambiguity caught during self-review:** PayEntries' Worker ID column (A) has no dropdown constraining it to valid Workers rows, so a typo would otherwise cascade a raw `#N/A` into Gross Pay and Net Pay. Fixed inline in Task 2 by wrapping every `Workers` lookup in `IFERROR(...,"Invalid Worker ID")`, with Net Pay additionally checking for that literal string so it shows blank instead of `#VALUE!` from subtracting text.
