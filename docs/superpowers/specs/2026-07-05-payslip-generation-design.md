# Payslip Generation — Design

Date: 2026-07-05
Workbook: `expense_tracker_2026.xlsm`

## Purpose

Let the workbook maintainer generate a printable payslip (PDF) for a site
labourer or contractor for a given month, and have the resulting net pay
automatically recorded as a Payments-table entry — without duplicating data
entry or breaking the existing month-sheet/Dashboard structure.

## Scope

Covers both daily-wage labourers (paid by days worked × daily rate) and
contractors (paid a fixed agreed amount per period), for monthly pay
periods aligned to the workbook's existing month sheets (Jan2026 …
Dec2026). No day-by-day attendance grid — just a total "days worked" figure
per labourer per month.

## Data Model

Three new sheets are added to `expense_tracker_2026.xlsm`:

### Workers sheet

One row per person, the roster / single source of truth for who gets paid:

| Column | Meaning |
|---|---|
| Worker ID | Short unique code (e.g. `W001`) — used as the join key everywhere else, so renaming a worker or having two workers share a name never causes ambiguity. |
| Name | Display name. |
| Role | `Labourer` or `Contractor` — determines which existing Sub-Category a payslip maps to (`Labour Wages` or `Contractor Payments`, both already under Main Category `Labour & Contractors` — no new categories needed). |
| Rate Type | `Per Day` or `Fixed Amount`. |
| Rate or Agreed Amount | Daily rate (Labourer) or the standard agreed payment (Contractor). |
| Default Payment Mode | `Bank` or `Cash` — pre-fills PayEntries/Payslip, can be overridden per period. |
| Active | `Y`/`N` — only Active workers appear in the Payslip sheet's Worker dropdown. |

### PayEntries sheet

One row per Worker+Month pay period — the actual data-entry point and the
running payslip history:

| Column | Meaning |
|---|---|
| Worker ID | Joins to Workers. |
| Worker Name | Denormalized for readability (not the join key). |
| Month | One of the 12 month names. |
| Days Worked | Entered for Labourers; left blank for Contractors. |
| Gross Pay | Formula: `Days Worked × Rate` (Labourer) or the Agreed Amount (Contractor) — never typed directly. |
| Advance/Deduction | Entered manually — money already drawn mid-period that reduces the payout. |
| Net Pay | Formula: `Gross Pay − Advance/Deduction`. |
| Payment Mode | Defaults from Workers, overridable per period. |
| Payments Row Ref | Hidden bookkeeping cell (e.g. `Jan2026!K37`) recording which Payments-table row this entry last wrote to, so regenerating the same payslip updates that row instead of creating a duplicate. |
| Comments | Free text. |

This is where the maintainer actually records "worker X worked 26 days in
March" each pay period — not on the Payslip sheet itself.

### Payslip sheet

A single printable template, not a data store:

- Two dropdown cells at the top: **Worker** (sourced from `Workers!Name`,
  Active rows only) and **Month** (the 12 month names).
- Everything below is formula-driven off those two cells: worker details
  via `INDEX/MATCH` against Workers; pay-period figures via `INDEX/MATCH`
  against PayEntries for the selected Worker+Month.
- If no PayEntries row exists yet for that combination, the template shows
  "No pay entry found for this period" rather than blanks, zeros, or
  `#N/A` — so it's obvious data entry is needed first, instead of
  accidentally exporting an empty payslip.
- Layout: site/project name header, pay period, worker name + role, an
  earnings line (Days Worked × Rate = Gross, or Agreed Amount for
  contractors), Advance/Deduction line, Net Pay total, Payment Mode, and a
  generation date stamp. Styled with the workbook's existing palette
  (slate headers, bordered tables), matching the Dashboard's visual
  language.

## Macro Behavior

One button, "Generate Payslip," on the Payslip sheet runs a macro that:

1. **Writes/updates the Payments row.** Reads the selected Worker+Month,
   finds the matching PayEntries row. If that row's `Payments Row Ref`
   already points at a Payments-table cell (a prior generation), updates
   that row in place. Otherwise writes to the first blank row in the
   target month sheet's Payments table (rows 3–100, the same range every
   other Payments entry uses) with: Amount = Net Pay, Main Category =
   `Labour & Contractors`, Sub-Category = `Labour Wages` or `Contractor
   Payments` per the worker's Role, Payment Mode = the entry's Payment
   Mode, Comments = `"Payslip: <Worker Name> - <Month>"`. Writes the new
   row's address back into `Payments Row Ref`.
2. **Exports to PDF.** Same `ExportAsFixedFormat` pattern as the existing
   Dashboard export macro, print area scoped to just the Payslip
   template's cells, filename `Payslip_<WorkerID>_<Month>.pdf`.

If no PayEntries row exists for the selected Worker+Month, the macro stops
with a message directing the maintainer to fill in PayEntries first —
it never writes a zero/blank payment.

## Edge Cases

- **Regenerating the same payslip** (e.g. correcting a days-worked number)
  updates the existing Payments row via `Payments Row Ref` instead of
  adding a duplicate expense.
- **Inactive workers** are excluded from the Payslip dropdown, but their
  historical PayEntries/Payments rows are untouched.
- **Contractors** have no Days Worked value; Gross Pay comes directly from
  the Agreed Amount.
- **Payments table full** (all 100 rows used for a given month): the macro
  shows a clear error rather than silently overwriting a real row or
  failing with an obscure cell-reference error.

## Out of Scope

- Day-by-day attendance tracking (only a total days-worked figure per
  period).
- Weekly/fortnightly pay cycles (monthly only, aligned to existing month
  sheets).
- Bulk-generating payslips for every worker at once (one at a time via
  the dropdown picker).
- Statutory deductions (tax, PF, ESI, etc.) — only a single generic
  Advance/Deduction line.
