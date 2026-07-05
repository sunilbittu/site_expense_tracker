# Expense Tracker Dashboard — Design

Date: 2026-07-05
Workbook: `expense_tracker_2026.xlsx`

## Purpose

Give the workbook maintainer a single-glance view of cash position, spend
by category, income sources, and the monthly receipts/payments trend,
without leaving Excel or maintaining a second tool. Personal use, not for
external sharing.

## Approach

Add a new **Dashboard** sheet to the existing workbook. It holds no raw
data — only formulas that aggregate the 12 month sheets (`Jan2026` …
`Dec2026`), plus native Excel charts built from those formulas. This keeps
the dashboard live: as the maintainer fills in Receipts/Payments rows
through the year, the dashboard updates automatically on open/recalc.

Native Excel PivotTables were considered but ruled out — the tooling used
to build this workbook (openpyxl) cannot create pivot caches/PivotTables,
only read them. SUMIFS-based aggregation + native charts achieves the same
live-update behavior without that limitation.

## Data & Aggregation

All aggregation tables live on the Dashboard sheet itself, below the
charts, in a section labeled "Supporting Data (auto-calculated — do not
edit)".

1. **Monthly Trend** (12 rows, one per month): reuses each month sheet's
   existing Summary block directly — `='Jan2026'!B104` (Total Receipts),
   `!B105` (Total Payments), `!B106` (Closing Balance). No new formulas
   needed for the per-month totals themselves.
2. **Spend by Main Category** (10 rows): for each of the 10 main
   categories, sum `SUMIFS(<month>!K3:K100, <month>!L3:L100, category)`
   across all 12 month sheets (12 SUMIFS terms added per row, generated
   programmatically).
3. **Income by Category** (4 rows): same SUMIFS pattern over each month's
   Receipts table (Amount column D, Category column E).
4. **Bank vs Cash split** (2 rows): SUMIFS over Payment Mode columns in
   both Receipts and Payments tables, across all 12 sheets, combined per
   mode (Bank / Cash).
5. **Sub-Category Detail** (68 rows): Main Category + Sub-Category +
   summed Amount across all 12 sheets, built as a real Excel Table
   (ListObject with autofilter) so the maintainer can sort/filter by
   amount once real numbers exist. Kept in natural Main→Sub grouping
   order at build time (sorting by amount now would be meaningless since
   all values are 0 pre-data-entry).

## KPI Cards

Four cards at the top of the sheet:

- **Current Cash Balance** — closing balance of the *last month with any
  activity*, not simply Dec2026. Found via a `LOOKUP(2, 1/((receipts<>0)
  + (payments<>0)), balances)` formula over the Monthly Trend table,
  wrapped in `IFERROR(..., 0)` so it shows 0 cleanly before any data
  exists instead of `#N/A`.
- **YTD Total Receipts** — sum of the Monthly Trend Receipts column.
- **YTD Total Payments** — sum of the Monthly Trend Payments column.
- **Net Position YTD** — YTD Receipts − YTD Payments.

## Charts

Four native Excel charts in a 2×2 grid, all sourced from the Supporting
Data tables:

1. **Balance Trend** — line chart, 12 months, Closing Balance series.
2. **Monthly Receipts vs Payments** — clustered bar chart, 12 months, two
   series (Receipts, Payments).
3. **Spend by Main Category** — bar chart, 10 categories.
4. **Income by Source** — pie chart, 4 income categories.

Sub-category data is not charted (68 categories is too many to read as a
chart) — it's available as a sortable table instead.

## Styling

Reuse the workbook's existing palette for visual consistency: green
(`2E7D32`) for receipts-related elements, red (`C62828`) for
payments-related elements, slate (`37474F`) for headers/KPI card accents,
light grey (`ECEFF1`) for labels — matching the Categories sheet and
per-month Summary blocks already in the workbook.

## Edge Cases

- Empty months (no rows entered yet) compute to 0 via SUM/SUMIFS — no
  `#DIV/0!` or `#N/A`, since no ratios are computed anywhere.
- Before any data exists for the year, the Current Cash Balance KPI falls
  back to 0 via `IFERROR` rather than showing an error.

## Out of Scope

- No external/shareable formatting pass (audience is the maintainer only).
- No budget-vs-actual comparison (no budget data exists in the workbook).
- No cross-year (2027+) rollup — out of scope until a 2027 workbook exists.
