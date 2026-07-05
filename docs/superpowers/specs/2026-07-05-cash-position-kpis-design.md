# Cash at Hand / Cash at Bank KPIs — Design

Date: 2026-07-05
Workbook: `expense_tracker_2026.xlsm`

## Purpose

Give the workbook maintainer a single-glance view of how much of the
current cash position (already shown as the combined "Current Cash
Balance" KPI) sits as physical cash on hand versus in the bank account.

## Background

Each month sheet tags every Receipts/Payments row with a Payment Mode
(`Cash` or `Bank`), but no running balance is split by mode anywhere —
only a single combined Opening/Closing Balance exists per month
(`Jan2026!B103` starts at literal `0`; each following month's Opening
Balance references the prior month's Closing Balance).

The Dashboard's existing "Bank vs Cash split" supporting-data table
(`BANKCASH_*` rows) is a different metric: it sums total transaction
*volume* per mode (receipts + payments combined) across the year — it
answers "how much moved through each mode," not "how much is left in
each mode." That table is unrelated to this feature and is not modified.

Cash on hand and the bank balance both started the year at zero (per the
existing Jan2026 Opening Balance of `0`) — there is no separate real
starting split to account for.

## Design

Two new KPI cards, styled and positioned like the existing four
(`Current Cash Balance`, `YTD Total Receipts`, `YTD Total Payments`, `Net
Position`), extending the same top-of-sheet KPI row:

- **Cash at Hand** (`I3` label / `I4` value): for each month sheet, sum
  `SUMIFS(<month>!D3:D100, <month>!F3:F100, "Cash")` (Cash-tagged
  receipts) minus `SUMIFS(<month>!K3:K100, <month>!N3:N100, "Cash")`
  (Cash-tagged payments), summed across all 12 months.
- **Cash at Bank** (`K3` label / `K4` value): identical formula, filtered
  to `"Bank"` instead of `"Cash"`.

Both are pure formulas — no new input cells, since both modes start at
zero. By construction, `Cash at Hand + Cash at Bank` always equals the
existing `Current Cash Balance` KPI, since that KPI is itself the
combined Opening Balance (`0`) plus total receipts minus total payments.

`KPI_CELLS` gains two entries (`cash_hand`: `I4`, `cash_bank`: `K4`) so
existing test patterns (e.g. the "every KPI resolves to a number, never
an error, against the real workbook" regression test) automatically cover
the new cards.

## Styling

Same slate (`37474F`) neutral color as `Current Cash Balance` and `Net
Position` — these are informational/summary cards, not receipts- or
payments-colored like the green/red cards.

## Edge Cases

- Empty months compute to 0 via SUMIFS, same as every other Dashboard
  formula — no error path needed.
- A mode's running balance can go negative in test/fixture data if
  payments in that mode aren't offset by receipts in the same mode
  (mathematically valid; not specifically guarded against, matching how
  the existing Net Position card also permits negative values).

## Out of Scope

- No change to the existing "Bank vs Cash split" volume table.
- No separate real-world starting cash/bank balances (both are zero at
  Jan2026 per the existing Opening Balance).
- No mode-level monthly trend chart (only the two current-position KPI
  cards are requested).
