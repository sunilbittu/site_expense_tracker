# Cash at Hand / Cash at Bank KPIs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two new Dashboard KPI cards, "Cash at Hand" and "Cash at Bank," per `docs/superpowers/specs/2026-07-05-cash-position-kpis-design.md`.

**Architecture:** Extend `scripts/build_dashboard.py`'s existing KPI-card section with two more formula-only cards (`I3`/`I4` and `K3`/`K4`), computed the same way every other Dashboard aggregate is: a `SUMIFS`-per-month-sheet formula summed across all 12 months, filtered by Payment Mode instead of Category. No new sheets, input cells, or dependencies.

**Tech Stack:** Python 3, `openpyxl` (build workbook formulas/styling), `formulas` (evaluate formulas in tests to assert real computed values) — same stack already used by `build_dashboard.py`/`test_build_dashboard.py`.

## Global Constraints

- Workbook under edit: `expense_tracker_2026.xlsm` (repo root). Always load with `openpyxl.load_workbook(path, keep_vba=True)` — this workbook carries a real VBA macro (`GeneratePayslip`, `ExportDashboardToPDF`) that must survive every edit.
- `build_dashboard.build(path)` only deletes and recreates the `Dashboard` sheet — it never touches `Workers`, `PayEntries`, or `Payslip`, so it is safe to run directly against the real workbook without wiping the payroll data already entered there.
- Both Cash at Hand and Cash at Bank start the year at zero (per the existing `Jan2026!B103` Opening Balance of literal `0`) — no new input cells or opening-balance split are needed; both KPIs are pure formulas.
- `Cash at Hand + Cash at Bank` must always equal the existing `Current Cash Balance` KPI (`KPI_CELLS['balance']`) — this is a direct consequence of the formula design, not something to special-case.
- The existing "Bank vs Cash split" supporting-data table (`BANKCASH_*` constants, `_add_category_tables`, `_payment_mode_formula`) is a different metric (total transaction volume per mode) and must not be modified.
- Test fixture: use `scripts/test_helpers.sample_workbook(tmp_dir)` for any test that needs a workbook copy — it sets `Jan2026!D3=1000` (Receipt, Payment Mode `Bank`) and `Jan2026!K3=400` (Payment, Payment Mode `Cash`, Category Materials/Cement). Do not write a second/competing fixture helper.
- No new dependencies beyond `openpyxl` and `formulas` (already in `requirements.txt`).

---

### Task 1: Add Cash at Hand / Cash at Bank KPI formulas and styling

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `scripts/test_build_dashboard.py`

**Interfaces:**
- Consumes: `MONTH_SHEETS` (module constant, already defined at the top of `build_dashboard.py`); `KPI_CELLS` dict (already defined at `scripts/build_dashboard.py:19`).
- Produces: `KPI_CELLS['cash_hand'] = 'I4'`, `KPI_CELLS['cash_bank'] = 'K4'` — later tasks and the existing `test_builds_cleanly_against_real_workbook` regression test (which loops over `KPI_CELLS.values()`) pick these up automatically, no changes needed there.

- [ ] **Step 1: Write the failing test**

Add to `scripts/test_build_dashboard.py`, directly after `test_kpi_cards`:

```python
def test_cash_hand_and_bank_kpis():
    with tempfile.TemporaryDirectory() as tmp:
        path = sample_workbook(tmp)
        build_dashboard.build(path)
        sol = calc(path)
        # Fixture: Jan2026 has a 1000 Bank receipt and a 400 Cash payment.
        # Cash at Hand = Cash receipts (0) - Cash payments (400) = -400
        # Cash at Bank = Bank receipts (1000) - Bank payments (0) = 1000
        assert cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['cash_hand']) == -400
        assert cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['cash_bank']) == 1000
        # Cash at Hand + Cash at Bank must always equal Current Cash Balance
        balance = cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['balance'])
        cash_hand = cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['cash_hand'])
        cash_bank = cell_value(sol, path, 'Dashboard', build_dashboard.KPI_CELLS['cash_bank'])
        assert cash_hand + cash_bank == balance
        print("test_cash_hand_and_bank_kpis PASSED")
```

Update the `if __name__ == '__main__':` block at the bottom of `scripts/test_build_dashboard.py` to also call it, directly after `test_kpi_cards()`:

```python
    test_read_categories()
    test_monthly_trend_table()
    test_kpi_cards()
    test_cash_hand_and_bank_kpis()
    test_category_tables()
    test_subcategory_detail_table()
    test_charts()
    test_builds_cleanly_against_real_workbook()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `KeyError: 'cash_hand'` (or similar) — `KPI_CELLS` has no `cash_hand`/`cash_bank` entries yet.

- [ ] **Step 3: Implement the KPI formulas**

In `scripts/build_dashboard.py`, update the `KPI_CELLS` dict at line 19:

```python
KPI_CELLS = {
    'balance': 'A4', 'receipts': 'C4', 'payments': 'E4', 'net': 'G4',
    'cash_hand': 'I4', 'cash_bank': 'K4',
}
```

Add a new helper function directly after `_payment_mode_formula` (around line 116):

```python
def _mode_balance_formula(mode):
    quoted = '"%s"' % mode
    terms = [
        "SUMIFS(%s!D3:D100,%s!F3:F100,%s)-SUMIFS(%s!K3:K100,%s!N3:N100,%s)" % (m, m, quoted, m, m, quoted)
        for m in MONTH_SHEETS
    ]
    return "=" + "+".join(terms)
```

In `_add_kpi_cards`, add the two new cards after the existing `G4` card:

```python
    ws['G3'] = 'Net Position (YTD)'
    ws['G4'] = "=%s-%s" % (KPI_CELLS['receipts'], KPI_CELLS['payments'])

    ws['I3'] = 'Cash at Hand'
    ws['I4'] = _mode_balance_formula('Cash')

    ws['K3'] = 'Cash at Bank'
    ws['K4'] = _mode_balance_formula('Bank')
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 scripts/test_build_dashboard.py`
Expected: `test_cash_hand_and_bank_kpis PASSED`, and all other tests still print `PASSED`.

- [ ] **Step 5: Add styling for the new KPI cards**

In `_style_dashboard`, update the `kpi_label_colors` dict (around line 216) to add the two new columns, both slate (matching `Current Cash Balance` and `Net Position`, since these are neutral/summary cards, not receipts- or payments-colored):

```python
    kpi_label_colors = {
        'A': '37474F',  # Current Cash Balance - slate (neutral)
        'C': '2E7D32',  # YTD Total Receipts - green
        'E': 'C62828',  # YTD Total Payments - red
        'G': '37474F',  # Net Position (YTD) - slate (neutral/summary)
        'I': '37474F',  # Cash at Hand - slate (neutral/summary)
        'K': '37474F',  # Cash at Bank - slate (neutral/summary)
    }
```

This loop already styles both the label cell (row 3) and value cell (row 4, including `number_format = '#,##0.00'`) for every column in the dict, so no other styling code changes are needed.

Update the column width list (around line 286) to size column `K` (column `I` is already sized at 22, reused from the Income Category table):

```python
    for col_letter, width in [('A', 22), ('B', 18), ('C', 18), ('D', 22),
                               ('F', 22), ('G', 16), ('I', 22), ('J', 16),
                               ('K', 16), ('L', 14), ('M', 16)]:
        ws.column_dimensions[col_letter].width = width
```

- [ ] **Step 6: Run test to verify styling didn't break anything**

Run: `python3 scripts/test_build_dashboard.py`
Expected: all tests print `PASSED`.

- [ ] **Step 7: Commit**

```bash
git add scripts/build_dashboard.py scripts/test_build_dashboard.py
git commit -m "Add Cash at Hand / Cash at Bank KPI cards to Dashboard"
```

---

### Task 2: Rebuild the real workbook's Dashboard

**Files:**
- Modify: `expense_tracker_2026.xlsm` (the actual output of running the finished script)

**Interfaces:**
- Consumes: everything from Task 1.
- Produces: nothing further downstream — this is the final task.

- [ ] **Step 1: Confirm the real workbook is not open in Excel**

```bash
ls -la "expense_tracker_2026.xlsm" 2>&1
ls -la "~\$expense_tracker_2026.xlsm" 2>&1
```

Expected: the second command errors with "No such file or directory" (no Excel lock file present). If it does exist, stop and ask the user to close the file in Excel first — writing to an open workbook risks data loss.

- [ ] **Step 2: Build into the real workbook**

```bash
python3 scripts/build_dashboard.py expense_tracker_2026.xlsm
```

- [ ] **Step 3: Spot-check the real workbook**

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('expense_tracker_2026.xlsm', keep_vba=True)
print('vba_archive present:', wb.vba_archive is not None)
print('sheetnames:', wb.sheetnames)
dash = wb['Dashboard']
print('I3:', dash['I3'].value)
print('I4:', dash['I4'].value)
print('K3:', dash['K3'].value)
print('K4:', dash['K4'].value)
"
```

Expected: `vba_archive present: True`; `Workers`, `PayEntries`, `Payslip` all still present in `sheetnames`; `I3` = `Cash at Hand`, `K3` = `Cash at Bank`, `I4`/`K4` hold the `SUMIFS`-based formula strings from Task 1.

- [ ] **Step 4: Commit**

```bash
git add expense_tracker_2026.xlsm
git commit -m "Build Cash at Hand / Cash at Bank KPIs into the real workbook"
```

---

## Self-Review Notes

- **Spec coverage:** Two new KPI cards (Task 1's formulas + styling), applied to the real workbook (Task 2) — both spec sections ("Design" and the KPI values themselves) map to tasks. The "Bank vs Cash split" table is explicitly left untouched per the Global Constraints and is not referenced by any task's code changes.
- **Placeholder scan:** no TBD/TODO; every step has complete, runnable code.
- **Type/name consistency:** `KPI_CELLS['cash_hand']`/`KPI_CELLS['cash_bank']` are defined once in Task 1 Step 3 and referenced identically (by dict key, never by hardcoded cell address) in the Task 1 test and Task 2's spot-check. `_mode_balance_formula` is defined once and used twice (`'Cash'`, `'Bank'`) with no renames across steps.
