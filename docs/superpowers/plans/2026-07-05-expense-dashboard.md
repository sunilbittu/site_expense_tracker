# Expense Tracker Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a live "Dashboard" sheet to `expense_tracker_2026.xlsx` with KPI cards, aggregation tables, and native Excel charts, per `docs/superpowers/specs/2026-07-05-expense-dashboard-design.md`.

**Architecture:** A single script, `scripts/build_dashboard.py`, exposes a `build(path)` function that (re)creates a "Dashboard" sheet in the workbook at `path`, using only formulas that reference the 12 existing month sheets — no raw data is duplicated onto the Dashboard sheet. The script reads the category list dynamically from the existing `Categories` sheet (single source of truth) rather than hardcoding it a second time.

**Tech Stack:** Python 3, `openpyxl` (build workbook structure/formulas/charts), `formulas` (evaluate formulas in tests to assert real computed values, since openpyxl cannot evaluate formulas itself).

## Global Constraints

- Workbook under edit: `expense_tracker_2026.xlsx` (repo root).
- Existing month sheet layout (already built, do not change): Receipts table columns A–F (`SNO, Date, Receipts, Amount, Category, Payment Mode`), Payments table columns I–O (`Date, Payments, Amount, Main Category, Sub-Category, Payment Mode, Comments`), data rows 3–100, Summary block at rows 102–106 with `Opening Balance` at B103, `Total Receipts` at B104, `Total Payments` at B105, `Net / Closing Balance` at B106.
- `MONTH_SHEETS = ['Jan2026','Feb2026','Mar2026','Apr2026','May2026','Jun2026','Jul2026','Aug2026','Sep2026','Oct2026','Nov2026','Dec2026']` — exact order, do not resort.
- Categories sheet layout (already built, do not change): Income Categories at `A3:A6`, Unique Main Categories at `F3:F12`, per-category sub-category grid with headers at row 2 columns H onward and values starting row 3 in each column.
- The Dashboard sheet must be rebuildable (idempotent): running `build()` twice must not create duplicate sheets or leftover data.
- No new top-level dependencies beyond `openpyxl` and `formulas` (see `requirements.txt`).

---

### Task 1: Script scaffold, category reader, and Monthly Trend table

**Files:**
- Create: `scripts/build_dashboard.py`
- Create: `scripts/test_helpers.py`
- Create: `scripts/test_build_dashboard.py`
- Test: `scripts/test_build_dashboard.py` (same file, runs as a plain script)

**Interfaces:**
- Produces: `build_dashboard.MONTH_SHEETS` (list[str], the 12 month sheet names, in order).
- Produces: `build_dashboard.read_categories(wb) -> (income_categories: list[str], main_categories: list[str], main_sub: dict[str, list[str]])` — reads from `wb['Categories']`.
- Produces: `build_dashboard.build(path: str) -> None` — top-level entry point, idempotently (re)builds the `Dashboard` sheet in the workbook at `path` and saves it.
- Produces: `build_dashboard.MONTHLY_TREND_FIRST_ROW = 45`, `MONTHLY_TREND_LAST_ROW = 56`, columns `A` (Month), `B` (Receipts), `C` (Payments), `D` (Closing Balance) — later tasks (KPIs, charts) read from this table.
- Produces: `test_helpers.sample_workbook(tmp_dir: str) -> str` — copies `expense_tracker_2026.xlsx` into `tmp_dir`, injects one sample Receipts row and one sample Payments row into `Jan2026`, saves, and returns the new path.
- Consumes: nothing (first task).

- [ ] **Step 1: Install dependencies**

```bash
pip3 install -r requirements.txt
```

- [ ] **Step 2: Write `scripts/test_helpers.py`**

```python
import shutil
import os
import openpyxl

SOURCE_WORKBOOK = os.path.join(os.path.dirname(__file__), '..', 'expense_tracker_2026.xlsx')


def sample_workbook(tmp_dir):
    """Copy the real workbook into tmp_dir and inject one sample Receipts
    row and one sample Payments row into Jan2026. Returns the new path."""
    dest = os.path.join(tmp_dir, 'sample.xlsx')
    shutil.copyfile(SOURCE_WORKBOOK, dest)

    wb = openpyxl.load_workbook(dest)
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
```

- [ ] **Step 3: Write the failing test for `read_categories` and the Monthly Trend table**

```python
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


if __name__ == '__main__':
    test_read_categories()
    test_monthly_trend_table()
```

- [ ] **Step 4: Run test to verify it fails**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `ModuleNotFoundError: No module named 'build_dashboard'` (the module doesn't exist yet).

- [ ] **Step 5: Write `scripts/build_dashboard.py` (scaffold + category reader + Monthly Trend table)**

```python
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


def build(path):
    wb = openpyxl.load_workbook(path)
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']
    ws = wb.create_sheet('Dashboard', 0)

    ws.cell(row=1, column=1, value='Site Expense Dashboard')
    _add_monthly_trend_table(ws)

    wb.save(path)


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else 'expense_tracker_2026.xlsx')
```

- [ ] **Step 6: Run test to verify it passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `test_read_categories PASSED` then `test_monthly_trend_table PASSED`.

- [ ] **Step 7: Commit**

```bash
git add requirements.txt scripts/build_dashboard.py scripts/test_helpers.py scripts/test_build_dashboard.py
git commit -m "Add Dashboard sheet scaffold with category reader and Monthly Trend table"
```

---

### Task 2: KPI cards

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `scripts/test_build_dashboard.py`

**Interfaces:**
- Consumes: `MONTH_SHEETS`, `MONTHLY_TREND_FIRST_ROW`, `MONTHLY_TREND_LAST_ROW` from Task 1. Monthly Trend table columns `B` (Receipts), `C` (Payments), `D` (Closing Balance), rows 45–56.
- Produces: `build_dashboard.KPI_LABEL_ROW = 3`, `KPI_VALUE_ROW = 4`, and `KPI_CELLS = {'balance': 'A4', 'receipts': 'C4', 'payments': 'E4', 'net': 'G4'}` — later tasks (charts, integration test) may reference these.
- Produces: `_add_kpi_cards(ws)` wired into `build()`.

- [ ] **Step 1: Write the failing test**

Add to `scripts/test_build_dashboard.py` (above the `if __name__ == '__main__':` block):

```python
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
```

And update the `__main__` block to also call `test_kpi_cards()`.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `AttributeError: module 'build_dashboard' has no attribute 'KPI_CELLS'`

- [ ] **Step 3: Implement KPI cards in `scripts/build_dashboard.py`**

Add module-level constants (near the top, after `MONTHLY_TREND_LAST_ROW`):

```python
KPI_LABEL_ROW = 3
KPI_VALUE_ROW = 4
KPI_CELLS = {'balance': 'A4', 'receipts': 'C4', 'payments': 'E4', 'net': 'G4'}
```

Add the function:

```python
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
```

Wire it into `build()`, right after `_add_monthly_trend_table(ws)`:

```python
    _add_monthly_trend_table(ws)
    _add_kpi_cards(ws)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: all four `PASSED` lines print, including `test_kpi_cards PASSED`.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_dashboard.py scripts/test_build_dashboard.py
git commit -m "Add KPI cards to Dashboard sheet"
```

---

### Task 3: Category aggregation tables (Main Category spend, Income by Category, Bank vs Cash)

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `scripts/test_build_dashboard.py`
- Modify: `scripts/test_helpers.py`

**Interfaces:**
- Consumes: `MONTH_SHEETS`, `read_categories(wb)` from Task 1.
- Produces: `MAIN_CAT_FIRST_ROW = 45`, `MAIN_CAT_LAST_ROW` (`45 + len(main_categories) - 1`, computed at build time and also exposed as a module constant sized for the current 10 categories: `MAIN_CAT_LAST_ROW = 54`), columns `F` (label), `G` (amount).
- Produces: `INCOME_CAT_FIRST_ROW = 45`, `INCOME_CAT_LAST_ROW = 48`, columns `I` (label), `J` (amount).
- Produces: `BANKCASH_FIRST_ROW = 45`, `BANKCASH_LAST_ROW = 46`, columns `L` (label), `M` (amount).
- Produces: `_add_category_tables(ws, income_categories, main_categories)` wired into `build()`.

- [ ] **Step 1: Extend `test_helpers.sample_workbook` with a second transaction**

The Task 1 sample only covers `Materials`/`Cement`/`Sales`. To distinguish category totals from each other, add a second Receipts row (different income category) and confirm the Payments row's Payment Mode is exercised. Replace the body of `sample_workbook` in `scripts/test_helpers.py` with:

```python
def sample_workbook(tmp_dir):
    """Copy the real workbook into tmp_dir and inject sample rows into
    Jan2026 that exercise every aggregation dimension. Returns the new path."""
    dest = os.path.join(tmp_dir, 'sample.xlsx')
    shutil.copyfile(SOURCE_WORKBOOK, dest)

    wb = openpyxl.load_workbook(dest)
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
```

(Unchanged from Task 1 — confirmed still fits this task's assertions, since `Sales` and `Materials`/`Cement` are each distinct from every other category, so their totals are unambiguous against the zeroed-out rest.)

- [ ] **Step 2: Write the failing test**

Add to `scripts/test_build_dashboard.py`:

```python
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
```

Update `__main__` to also call `test_category_tables()`.

- [ ] **Step 3: Run test to verify it fails**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `AttributeError: module 'build_dashboard' has no attribute 'MAIN_CAT_FIRST_ROW'`

- [ ] **Step 4: Implement category aggregation tables in `scripts/build_dashboard.py`**

Add module-level constants:

```python
MAIN_CAT_HEADER_ROW = 44
MAIN_CAT_FIRST_ROW = 45
MAIN_CAT_LAST_ROW = 54  # 10 main categories

INCOME_CAT_HEADER_ROW = 44
INCOME_CAT_FIRST_ROW = 45
INCOME_CAT_LAST_ROW = 48  # 4 income categories

BANKCASH_HEADER_ROW = 44
BANKCASH_FIRST_ROW = 45
BANKCASH_LAST_ROW = 46  # Bank, Cash
```

Add the function:

```python
def _main_category_formula(main_cat_cell):
    terms = ["SUMIFS(%s!K3:K100,%s!L3:L100,%s)" % (m, m, main_cat_cell) for m in MONTH_SHEETS]
    return "=" + "+".join(terms)


def _income_category_formula(income_cat_cell):
    terms = ["SUMIFS(%s!D3:D100,%s!E3:E100,%s)" % (m, m, income_cat_cell) for m in MONTH_SHEETS]
    return "=" + "+".join(terms)


def _payment_mode_formula(mode_cell):
    terms = []
    for m in MONTH_SHEETS:
        terms.append("SUMIFS(%s!D3:D100,%s!F3:F100,%s)" % (m, m, mode_cell))
        terms.append("SUMIFS(%s!K3:K100,%s!N3:N100,%s)" % (m, m, mode_cell))
    return "=" + "+".join(terms)


def _add_category_tables(ws, income_categories, main_categories):
    ws.cell(row=MAIN_CAT_HEADER_ROW, column=6, value='Main Category')
    ws.cell(row=MAIN_CAT_HEADER_ROW, column=7, value='Amount')
    for i, cat in enumerate(main_categories):
        row = MAIN_CAT_FIRST_ROW + i
        ws.cell(row=row, column=6, value=cat)
        ws.cell(row=row, column=7, value=_main_category_formula("$F%d" % row))

    ws.cell(row=INCOME_CAT_HEADER_ROW, column=9, value='Income Category')
    ws.cell(row=INCOME_CAT_HEADER_ROW, column=10, value='Amount')
    for i, cat in enumerate(income_categories):
        row = INCOME_CAT_FIRST_ROW + i
        ws.cell(row=row, column=9, value=cat)
        ws.cell(row=row, column=10, value=_income_category_formula("$I%d" % row))

    ws.cell(row=BANKCASH_HEADER_ROW, column=12, value='Payment Mode')
    ws.cell(row=BANKCASH_HEADER_ROW, column=13, value='Amount')
    for i, mode in enumerate(['Bank', 'Cash']):
        row = BANKCASH_FIRST_ROW + i
        ws.cell(row=row, column=12, value=mode)
        ws.cell(row=row, column=13, value=_payment_mode_formula("$L%d" % row))
```

Update `build()` to read categories and wire the new function in:

```python
def build(path):
    wb = openpyxl.load_workbook(path)
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']
    ws = wb.create_sheet('Dashboard', 0)
    income_categories, main_categories, main_sub = read_categories(wb)

    ws.cell(row=1, column=1, value='Site Expense Dashboard')
    _add_monthly_trend_table(ws)
    _add_kpi_cards(ws)
    _add_category_tables(ws, income_categories, main_categories)

    wb.save(path)
```

- [ ] **Step 5: Run test to verify it passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: all five `PASSED` lines print, including `test_category_tables PASSED`.

- [ ] **Step 6: Commit**

```bash
git add scripts/build_dashboard.py scripts/test_build_dashboard.py scripts/test_helpers.py
git commit -m "Add Main Category, Income Category, and Bank/Cash aggregation tables"
```

---

### Task 4: Sub-Category Detail table (Excel Table with autofilter)

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `scripts/test_build_dashboard.py`

**Interfaces:**
- Consumes: `MONTH_SHEETS` from Task 1; `main_sub` dict from `read_categories()`.
- Produces: `SUBCAT_HEADER_ROW = 60`, `SUBCAT_FIRST_ROW = 61`, `SUBCAT_LAST_ROW` (`61 + 68 - 1` = `128`), columns `A` (Main Category), `B` (Sub-Category), `C` (Amount).
- Produces: `_add_subcategory_detail_table(ws, main_sub)` wired into `build()`, creates an `openpyxl.worksheet.table.Table` named `SubCategoryDetail` over `A60:C128`.

- [ ] **Step 1: Write the failing test**

Add to `scripts/test_build_dashboard.py`:

```python
def test_subcategory_detail_table():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_dashboard.build(path)
        sol = calc(path)

        import openpyxl
        wb = openpyxl.load_workbook(path)
        dash = wb['Dashboard']

        target_row = None
        for r in range(build_dashboard.SUBCAT_FIRST_ROW, build_dashboard.SUBCAT_LAST_ROW + 1):
            if dash.cell(row=r, column=1).value == 'Materials' and dash.cell(row=r, column=2).value == 'Cement':
                target_row = r
        assert target_row is not None
        assert cell_value(sol, path, 'Dashboard', 'C%d' % target_row) == 400

        assert 'SubCategoryDetail' in dash.tables
        print("test_subcategory_detail_table PASSED")
```

Update `__main__` to also call `test_subcategory_detail_table()`.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `AttributeError: module 'build_dashboard' has no attribute 'SUBCAT_FIRST_ROW'`

- [ ] **Step 3: Implement the Sub-Category Detail table**

Add module-level constants:

```python
SUBCAT_HEADER_ROW = 60
SUBCAT_FIRST_ROW = 61
SUBCAT_LAST_ROW = 128  # 68 sub-categories across all main categories
```

Add the function (needs `from openpyxl.worksheet.table import Table, TableStyleInfo` added to the imports at the top of the file):

```python
def _subcategory_formula(main_cell, sub_cell):
    terms = [
        "SUMIFS(%s!K3:K100,%s!L3:L100,%s,%s!M3:M100,%s)" % (m, m, main_cell, m, sub_cell)
        for m in MONTH_SHEETS
    ]
    return "=" + "+".join(terms)


def _add_subcategory_detail_table(ws, main_sub):
    ws.cell(row=SUBCAT_HEADER_ROW, column=1, value='Main Category')
    ws.cell(row=SUBCAT_HEADER_ROW, column=2, value='Sub-Category')
    ws.cell(row=SUBCAT_HEADER_ROW, column=3, value='Amount')

    row = SUBCAT_FIRST_ROW
    for main, subs in main_sub.items():
        for sub in subs:
            ws.cell(row=row, column=1, value=main)
            ws.cell(row=row, column=2, value=sub)
            ws.cell(row=row, column=3, value=_subcategory_formula("$A%d" % row, "$B%d" % row))
            row += 1

    table = Table(displayName='SubCategoryDetail', ref='A%d:C%d' % (SUBCAT_HEADER_ROW, SUBCAT_LAST_ROW))
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(table)
```

Wire it into `build()`, right after `_add_category_tables(...)`:

```python
    _add_category_tables(ws, income_categories, main_categories)
    _add_subcategory_detail_table(ws, main_sub)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: all six `PASSED` lines print, including `test_subcategory_detail_table PASSED`.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_dashboard.py scripts/test_build_dashboard.py
git commit -m "Add Sub-Category Detail table as a sortable Excel Table"
```

---

### Task 5: Charts

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `scripts/test_build_dashboard.py`

**Interfaces:**
- Consumes: Monthly Trend table (Task 1), Main Category table (Task 3), Income Category table (Task 3) — cell ranges only, no new formulas.
- Produces: `_add_charts(ws)` wired into `build()`. No new module constants (charts are display-only, nothing downstream reads them).

- [ ] **Step 1: Write the failing test**

Add to `scripts/test_build_dashboard.py` (this test is structural — `formulas` can't evaluate chart objects, so it inspects the chart objects directly via openpyxl):

```python
def test_charts():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_dashboard.build(path)

        import openpyxl
        from openpyxl.chart import LineChart, BarChart, PieChart
        wb = openpyxl.load_workbook(path)
        dash = wb['Dashboard']

        assert len(dash._charts) == 4
        chart_types = [type(c) for c in dash._charts]
        assert chart_types.count(LineChart) == 1
        assert chart_types.count(BarChart) == 2
        assert chart_types.count(PieChart) == 1
        print("test_charts PASSED")
```

Update `__main__` to also call `test_charts()`.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `AssertionError` (`len(dash._charts) == 0`, since no charts exist yet).

- [ ] **Step 3: Implement charts in `scripts/build_dashboard.py`**

Add to the imports at the top of the file:

```python
from openpyxl.chart import LineChart, BarChart, PieChart, Reference
```

Add the function:

```python
def _add_charts(ws):
    months_ref = Reference(ws, min_col=1, min_row=MONTHLY_TREND_FIRST_ROW, max_row=MONTHLY_TREND_LAST_ROW)

    balance_chart = LineChart()
    balance_chart.title = 'Balance Trend'
    balance_data = Reference(ws, min_col=4, min_row=MONTHLY_TREND_HEADER_ROW, max_row=MONTHLY_TREND_LAST_ROW)
    balance_chart.add_data(balance_data, titles_from_data=True)
    balance_chart.set_categories(months_ref)
    ws.add_chart(balance_chart, 'A7')

    trend_chart = BarChart()
    trend_chart.type = 'col'
    trend_chart.grouping = 'clustered'
    trend_chart.title = 'Monthly Receipts vs Payments'
    trend_data = Reference(ws, min_col=2, max_col=3, min_row=MONTHLY_TREND_HEADER_ROW, max_row=MONTHLY_TREND_LAST_ROW)
    trend_chart.add_data(trend_data, titles_from_data=True)
    trend_chart.set_categories(months_ref)
    ws.add_chart(trend_chart, 'J7')

    main_cat_chart = BarChart()
    main_cat_chart.type = 'bar'
    main_cat_chart.title = 'Spend by Main Category'
    main_cat_data = Reference(ws, min_col=7, min_row=MAIN_CAT_HEADER_ROW, max_row=MAIN_CAT_LAST_ROW)
    main_cat_categories = Reference(ws, min_col=6, min_row=MAIN_CAT_FIRST_ROW, max_row=MAIN_CAT_LAST_ROW)
    main_cat_chart.add_data(main_cat_data, titles_from_data=True)
    main_cat_chart.set_categories(main_cat_categories)
    ws.add_chart(main_cat_chart, 'A24')

    income_chart = PieChart()
    income_chart.title = 'Income by Source'
    income_data = Reference(ws, min_col=10, min_row=INCOME_CAT_HEADER_ROW, max_row=INCOME_CAT_LAST_ROW)
    income_categories_ref = Reference(ws, min_col=9, min_row=INCOME_CAT_FIRST_ROW, max_row=INCOME_CAT_LAST_ROW)
    income_chart.add_data(income_data, titles_from_data=True)
    income_chart.set_categories(income_categories_ref)
    ws.add_chart(income_chart, 'J24')
```

Wire it into `build()`, right after `_add_subcategory_detail_table(...)`:

```python
    _add_subcategory_detail_table(ws, main_sub)
    _add_charts(ws)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: all seven `PASSED` lines print, including `test_charts PASSED`.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_dashboard.py scripts/test_build_dashboard.py
git commit -m "Add Dashboard charts: balance trend, monthly trend, category spend, income split"
```

---

### Task 6: Styling pass, final integration against the real workbook, and regression check

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `scripts/test_build_dashboard.py`
- Modify: `expense_tracker_2026.xlsx` (the actual output of running the finished script)

**Interfaces:**
- Consumes: everything from Tasks 1–5.
- Produces: nothing further downstream — this is the final task.

- [ ] **Step 1: Write the failing test — full-structure regression check against the real (currently empty) workbook**

Add to `scripts/test_build_dashboard.py`:

```python
def test_builds_cleanly_against_real_workbook():
    with tempfile.TemporaryDirectory() as tmp:
        import shutil
        path = os.path.join(tmp, 'real_copy.xlsx')
        shutil.copyfile(REAL_WORKBOOK, path)
        build_dashboard.build(path)
        sol = calc(path)

        # With no data entered anywhere, every KPI must resolve to 0, not an error.
        for cell in build_dashboard.KPI_CELLS.values():
            value = cell_value(sol, path, 'Dashboard', cell)
            assert value == 0, "%s resolved to %r, expected 0" % (cell, value)

        print("test_builds_cleanly_against_real_workbook PASSED")
```

Update `__main__` to also call `test_builds_cleanly_against_real_workbook()`.

- [ ] **Step 2: Run test to verify it fails or passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: this one should already PASS, since Tasks 1–5 are functionally complete — it's a regression guard, not new behavior. If it fails, that indicates a bug introduced in an earlier task; fix `build_dashboard.py` until it passes before proceeding.

- [ ] **Step 3: Apply consistent styling to the Dashboard sheet**

Add to the imports at the top of the file:

```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
```

Add a styling function reusing the workbook's existing palette (green `2E7D32` for receipts, red `C62828` for payments, slate `37474F` for headers, light grey `ECEFF1` for labels — the same colors already used on the month sheets and Categories sheet):

```python
def _style_dashboard(ws):
    title_font = Font(bold=True, size=14)
    kpi_label_font = Font(bold=True, color='FFFFFF')
    kpi_label_fill = PatternFill(start_color='37474F', end_color='37474F', fill_type='solid')
    kpi_value_font = Font(bold=True, size=13)
    table_header_font = Font(bold=True, color='FFFFFF')
    table_header_fill = PatternFill(start_color='37474F', end_color='37474F', fill_type='solid')
    thin = Side(style='thin', color='B0B0B0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws['A1'].font = title_font

    for cell in KPI_CELLS.values():
        col = cell[0]
        label_cell = ws['%s%d' % (col, KPI_LABEL_ROW)]
        label_cell.font = kpi_label_font
        label_cell.fill = kpi_label_fill
        label_cell.alignment = Alignment(horizontal='center')
        value_cell = ws[cell]
        value_cell.font = kpi_value_font
        value_cell.number_format = '#,##0.00'
        value_cell.border = border

    header_rows_cols = [
        (MONTHLY_TREND_HEADER_ROW, [1, 2, 3, 4]),
        (MAIN_CAT_HEADER_ROW, [6, 7]),
        (INCOME_CAT_HEADER_ROW, [9, 10]),
        (BANKCASH_HEADER_ROW, [12, 13]),
    ]
    for row, cols in header_rows_cols:
        for col in cols:
            c = ws.cell(row=row, column=col)
            c.font = table_header_font
            c.fill = table_header_fill
            c.alignment = Alignment(horizontal='center')
            c.border = border

    for col_letter in ['B', 'C', 'D', 'G', 'J', 'M']:
        for row in range(MONTHLY_TREND_FIRST_ROW, SUBCAT_LAST_ROW + 1):
            ws['%s%d' % (col_letter, row)].number_format = '#,##0.00'

    for col_letter, width in [('A', 22), ('B', 18), ('C', 18), ('D', 22),
                               ('F', 22), ('G', 16), ('I', 22), ('J', 16),
                               ('L', 14), ('M', 16)]:
        ws.column_dimensions[col_letter].width = width
```

Wire it into `build()`, right after `_add_charts(ws)`:

```python
    _add_charts(ws)
    _style_dashboard(ws)
```

- [ ] **Step 4: Run the full test suite to verify nothing broke**

Run: `python3 scripts/test_build_dashboard.py`
Expected: all eight `PASSED` lines print (styling doesn't change any formula, only presentation, so every prior assertion still holds).

- [ ] **Step 5: Build the Dashboard sheet into the real workbook**

```bash
python3 scripts/build_dashboard.py expense_tracker_2026.xlsx
```

- [ ] **Step 6: Manually spot-check the real workbook**

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('expense_tracker_2026.xlsx')
print(wb.sheetnames)
ws = wb['Dashboard']
print('charts:', len(ws._charts))
print('A1:', ws['A1'].value)
print('KPI receipts formula:', ws['C4'].value)
"
```

Expected: `Dashboard` appears first in `wb.sheetnames`; `charts: 4`; `A1: Site Expense Dashboard`; `KPI receipts formula` shows a `=SUM(...)` formula.

- [ ] **Step 7: Commit**

```bash
git add scripts/build_dashboard.py scripts/test_build_dashboard.py expense_tracker_2026.xlsx
git commit -m "Add styling pass and build the Dashboard sheet into the real workbook"
```

---

## Self-Review Notes

- **Spec coverage:** Monthly Trend (Task 1), KPI cards (Task 2), Main/Income/Bank-Cash aggregation (Task 3), Sub-Category Detail table (Task 4), 4 charts (Task 5), styling + real-workbook build (Task 6) — every section of the spec maps to a task. The spec's original `LOOKUP`-based Current Cash Balance was simplified to a direct `Dec2026!B106` reference during planning (see spec correction commit `f266014`); Task 2 implements the corrected version.
- **Placeholder scan:** no TBD/TODO; every step has complete, runnable code.
- **Type/name consistency:** `MONTH_SHEETS`, `MONTHLY_TREND_FIRST_ROW/LAST_ROW`, `KPI_CELLS`, `MAIN_CAT_FIRST_ROW/LAST_ROW`, `INCOME_CAT_FIRST_ROW/LAST_ROW`, `BANKCASH_FIRST_ROW/LAST_ROW`, `SUBCAT_FIRST_ROW/LAST_ROW` are each defined exactly once (Tasks 1–4) and referenced identically by name in every later task and in the styling pass — no renames across tasks.
