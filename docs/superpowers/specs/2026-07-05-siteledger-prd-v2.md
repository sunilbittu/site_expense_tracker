# SiteLedger — Product Requirements Document (v2)

**Status:** Target spec for a greenfield build (v1 PRD was written "as-built," but no such app exists in this repository or any known deployment — this v2 treats it as the build target)
**Last updated:** 2026-07-05
**Supersedes:** v1 PRD (2026-07-04)

## 0. Changes from v1 (gap fixes)

| # | Gap in v1 | Fix in v2 |
|---|---|---|
| 1 | Labeled "as-built" but the app doesn't exist | Reframed as target spec; development plan covers full build + deployment |
| 2 | **Security:** `VITE_RESEND_API_KEY` in client build — any visitor could extract the key and send email as you | Resend key is server-side only (`RESEND_API_KEY`, Vercel serverless env). No `VITE_` prefix, never in the client bundle |
| 3 | Receipt photo UI existed with no storage destination | Supabase Storage private bucket `receipts` with RLS, client-side compression, signed URLs for viewing (§5.3) |
| 4 | Locked entries permanently immutable — mistakes require offsetting "General" entries, corrupting category totals | Admin **void** mechanism: soft-void with reason, audit-preserving, excluded from totals (§6.1) |
| 5 | `projects` table stored a month-to-date total (stale-aggregate risk, no defined update path) | Computed on read via a SQL view; no stored aggregates (§8) |
| 6 | No timezone rules — "today," edit windows, and the 7 PM IST report could disagree on date boundaries | All date-boundary logic pinned to `Asia/Kolkata` (§6.4) |
| 7 | No self-service password recovery — forgotten password meant a manual Supabase dashboard reset | "Forgot password" email reset via Supabase Auth (§4.1) |
| 8 | Daily report recipient hardcoded to one email in code | `REPORT_RECIPIENTS` env var (comma-separated); manual re-trigger endpoint for missed runs (§7.2) |
| 9 | PDF-only export | CSV export added alongside PDF (§7.1) |
| 10 | Validation (amounts > 0, lock time) only described at UI level | DB-level CHECK constraints, enums, and a trigger that server-sets `locked_at` — client cannot forge it (§6.1, §8) |
| 11 | Edits within the 24 h window silently overwrote history — weak audit trail for an audit-focused product | Trigger-based `entries_audit` table records prior row on every UPDATE/DELETE (§8) |
| 12 | No test suite, no linter | Vitest + Testing Library, SQL-based RLS tests, ESLint + Prettier — required from Phase 0 (§12) |
| 13 | Flaky-network double submits (duplicate entries) undefined | Client-generated UUID primary key makes retried inserts idempotent; submit disabled while in flight (§6.5) |
| 14 | No environments, migration workflow, CI, monitoring, or backup story | Staging + production Supabase projects, `supabase` CLI migrations, GitHub Actions CI, Sentry, backup policy (§11–§12) |
| 15 | Realtime + RLS interaction unstated | Realtime authorization configured so subscriptions respect RLS (§6.2) |

Explicitly **kept** as non-goals for v1.0 (deliberate, not oversights): offline entry queueing, in-app user/assignment management, in-app report scheduling UI. See §9.

## 1. Overview

SiteLedger is a mobile-first web application for tracking day-to-day construction site expenses. Site supervisors log costs (labor, equipment, contractor fees, materials) against a project as they occur; an admin reviews, corrects (via void), and exports spend across all projects. It is a single-page React app backed by Supabase (Postgres + Auth + Realtime + Storage), deployed on Vercel, with a daily automated email summary.

## 2. Problem Statement

Construction sites accumulate cash expenses (daily labor wages, JCB rental hours, contractor payments, ad-hoc costs) that are traditionally tracked on paper or in scattered messages, making it hard for an owner/admin to see same-day spend, catch errors, or produce a reliable audit trail. SiteLedger gives supervisors a fast structured entry form for each expense type and gives the admin a consolidated, exportable view across all sites.

## 3. Users & Roles

Two roles, stored in `profiles.role`:

### Supervisor
- Assigned to one or more projects via `project_assignments`.
- Logs into a project picker, then a per-project home screen.
- Can add expense entries for assigned projects only, and edit/delete their own entries within a 24-hour window.
- Cannot see other supervisors' projects or the cross-project admin view.

### Admin
- Bypasses per-project assignment restrictions (`is_admin()` Postgres function used in RLS policies).
- Lands directly on the Admin Dashboard after login.
- Views and exports entries across all projects for any date range.
- Cannot create or edit entry contents, but **can void** any entry (including locked ones) with a mandatory reason (§6.1).

Roles, users, and project assignments are managed by the admin directly in the Supabase dashboard for v1.0, following a written runbook (part of the development plan). No self-service signup.

## 4. Core User Flows

### 4.1 Authentication
1. User signs in with email + password (Supabase Auth, `signInWithPassword`).
2. A **"Forgot password?"** link triggers Supabase's email reset flow (`resetPasswordForEmail` → magic link → new-password screen).
3. If `profiles.must_change_password` is true (default for newly created users), the user is forced through **Set a new password** before anything else — 8+ characters, at least one number, confirmation field must match.
4. On session, `DataContext` loads the user's profile, assigned projects (all projects for admin), plots grouped by project, and reference tables (worker types, contractors, JCB operators, work types).
5. Routing branches on role: admin → Admin Dashboard; supervisor → Project Picker (skipped straight to Home if only one assignment).

### 4.2 Supervisor flow
1. **Project Picker** — assigned projects only, with name, location, and month-to-date spend (computed, IST month).
2. **Home** — today's entries (IST) for the selected project, grouped by category with running totals, a category filter, and buttons to Add Expense, View History, or Switch Project. Live-updates via Supabase Realtime.
3. **Add Expense** — two-step sheet:
   - Step 1: pick a category (ranked by today's usage) or "Repeat last" to prefill the most recent entry.
   - Step 2: category-specific form (§5), plot selection, payment mode (cash / UPI / bank transfer), optional receipt photo, optional notes, live-calculated total. Entry date defaults to today (IST) and may be backdated at most 1 day (yesterday) — see §6.4.
4. **Entry Detail** — full entry view with edit-window countdown ("Edit window: Xh Ym" or "Locked"), receipt photo (signed URL), void banner if voided. Edit/delete only for the creator while unlocked.
5. **History** — all project entries grouped by date (descending) with per-date category subtotals, paginated (30 days per page).

### 4.3 Admin flow
1. **Admin Dashboard** — the admin's landing screen.
2. Date range picker (defaults to today, IST).
3. Summary card: grand total (voided entries excluded), entry count, stacked category-split bar.
4. Entries grouped by date → project, each with a category breakdown (count + amount). Voided entries shown struck-through with reason on tap.
5. **Void action** on any entry: requires a reason (min 10 chars), confirmation dialog; recorded as `voided_at`/`voided_by`/`void_reason`.
6. Export buttons: **PDF** and **CSV** for the selected range (§7.1).

## 5. Expense Categories

Every entry belongs to exactly one category (Postgres enum), with category-specific fields in `entries.details` JSONB:

| Category | Purpose | `details` fields | Total formula |
|---|---|---|---|
| **NMR (Labor)** | Daily-wage manual labor | `worker_type_id`, `worker_gender` (male/female), `worker_count`, `wage_per_worker`, `wage_overridden` | worker_count × wage_per_worker |
| **JCB** | Excavator/equipment rental | `operator_id`, `hours` (≥0.5), `hourly_rate` | hours × hourly_rate |
| **Contractor Fee** | Daily fee paid to a contractor | `contractor_id`, `fee`, optional `tip` | fee + tip |
| **Labor Contract** | Contract-based labor payment | `contractor_id`, `work_type`, `amount`, `payment_stage` (advance/running/final/other) | amount |
| **General** | Any other site cost | `description`, `amount` | amount |

Every entry requires a plot and a payment mode; all amounts are positive (enforced by DB CHECK on `total_amount` and client validation of `details`). Currency is INR (₹), formatted with Indian digit grouping.

### 5.1 Gender-based wages
Worker types carry male (`default_wage`) and female (`default_wage_female`) default wages. Selecting worker type + gender auto-fills the wage (female default seeded as male − ₹200); manual override recorded via `wage_overridden`.

### 5.2 Contractor filtering
Contractors are typed (`daily_fee`, `labor_contract`, `both`). Contractor Fee form offers `daily_fee`/`both`; Labor Contract form offers `labor_contract`/`both`.

### 5.3 Receipt photos
- Private Supabase Storage bucket `receipts`; object path `{project_id}/{entry_id}.jpg`.
- Client compresses to max 1600 px longest edge / ~300 KB JPEG before upload.
- Storage RLS mirrors entry access: creator may INSERT while entry is unlocked; project members and admin may SELECT.
- Displayed via short-lived signed URLs (never public).
- Upload failure does not block entry save — the entry saves, and the photo can be retried from Entry Detail while unlocked.

## 6. Business Rules

### 6.1 Entry locking and admin void
- A DB trigger sets `locked_at = created_at + interval '24 hours'` at insert — the client cannot supply or forge it.
- RLS permits UPDATE/DELETE only by the entry's creator while `locked_at > now()` (and the entry is not voided).
- After the window, the entry is immutable to the creator. The **admin may void** any entry at any time: an admin-only RLS policy permits updating exactly the void columns (`voided_at`, `voided_by`, `void_reason`), enforced by a trigger that rejects admin changes to any other column. Voids are permanent (no unvoid in v1.0). Voided entries are excluded from all totals and reports' sums but appear struck-through in lists and exports for the audit trail.

### 6.2 Realtime sync
Home and History subscribe to Postgres realtime changes on `entries` filtered to the active project. Supabase Realtime is configured with **RLS authorization enabled** (`private` channel / realtime authorization) so subscribers only receive rows their session can SELECT. Admin Dashboard remains one-shot fetch per date-range change.

### 6.3 Access control
All access control is enforced via Postgres RLS, not just UI. A supervisor's session cannot query another project's rows; an admin session cannot INSERT entries or modify non-void columns. RLS behavior is covered by automated SQL tests run in CI against a shadow database (§12).

### 6.4 Timezone & dates
- All date-boundary logic (today's entries, month-to-date, report "today," edit-window display) uses **Asia/Kolkata**.
- `entries.entry_date` is a `date` (no time component); `created_at`/`locked_at` are `timestamptz`.
- A DB CHECK constrains `entry_date` to between yesterday and today (IST) at insert time; supervisors may backdate exactly one day to cover late logging.
- The daily email runs at 13:30 UTC = 19:00 IST and reports the IST calendar day.

### 6.5 Idempotent writes
Entry `id` is a client-generated UUID (v4) sent with the INSERT. A retried insert after a network timeout hits the primary-key conflict and is treated as success — no duplicate entries. Submit buttons disable while a request is in flight.

## 7. Reporting

### 7.1 Exports (admin-triggered, client-side)
- **PDF** (`jsPDF` + `jspdf-autotable`): header (title, range, generated-at IST), line-item table (date, project, plot, category, description, payment mode, amount, voided marker), subtotals by category and by day, grand total (voided excluded). Filename `SiteLedger-Report-{start}-to-{end}.pdf`.
- **CSV**: same line items, one row per entry incl. void columns, UTF-8 with BOM (Excel-safe). Filename `SiteLedger-Report-{start}-to-{end}.csv`.

### 7.2 Daily automated email
Vercel cron (`api/daily-report.js`) at 13:30 UTC daily:
- Auth: request must carry `Authorization: Bearer ${CRON_SECRET}`.
- Uses `SUPABASE_SERVICE_ROLE_KEY` (server-side only) to bypass RLS.
- Sends via Resend (`RESEND_API_KEY`, server-side only) to every address in `REPORT_RECIPIENTS` (comma-separated env var).
- Content: subject with IST date + total; today's total + entry count; month-to-date total; per-project category breakdown. Voided entries excluded from sums.
- **Idempotent & re-runnable:** accepts optional `?date=YYYY-MM-DD` to regenerate any day's report (manual backfill for missed runs via `curl` with the secret). Failures return non-200 so they're visible in Vercel logs and Sentry.

## 8. Data Model (summary)

| Table | Purpose |
|---|---|
| `profiles` | One row per auth user: name, email, role enum (`supervisor`/`admin`), `must_change_password`. Auto-created by trigger on `auth.users` insert. |
| `projects` | Construction sites (id, name, location, active flag). **No stored aggregates** — month-to-date spend comes from the `project_month_totals` view. |
| `plots` | Sub-areas within a project (villas/plots, or a "Site-wide" catch-all). |
| `project_assignments` | Supervisor ↔ project mapping. |
| `entries` | Core transactions: client-UUID id, project, plot, category enum, `total_amount` numeric CHECK (> 0), payment mode enum, `receipt_path`, notes, `entry_date` date (CHECK: yesterday…today IST), `created_by`, `created_at`, `locked_at` (trigger-set), `voided_at`/`voided_by`/`void_reason`, `details` JSONB. Indexed on `(project_id, entry_date)` and `(entry_date)`. |
| `entries_audit` | Trigger-populated: prior row snapshot + actor + action on every entries UPDATE/DELETE. Insert-only; readable by admin. |
| `worker_types` | NMR reference: name, male/female default wages. |
| `contractors` | Reference: name, type enum (`daily_fee`/`labor_contract`/`both`). |
| `jcb_operators` | Reference: name, default hourly rate. |
| `work_types` | Labor-contract work categories (plastering, brickwork, …). |
| Storage: `receipts` bucket | Private; RLS mirrors entries access (§5.3). |

Schema, constraints, and RLS live in `supabase/migrations/` and are applied via the Supabase CLI (never hand-edited in the dashboard).

## 9. Non-Goals (v1.0 — deliberate)

- No self-service signup or in-app user/role/assignment management (Supabase dashboard + runbook).
- No offline entry queueing (forms preserve state on failed submit and support retry, but require connectivity to save).
- No in-app configuration of report recipients/schedule (env vars).
- No entry "unvoid," no editing of locked entries even by admin (void + re-enter is the correction path).
- No native mobile app (mobile-first responsive web, installable as a PWA shortcut only — no service worker).

## 10. Tech Stack

- **Frontend:** React 18 + Vite 6, manual `route` state machine in `App.jsx` (no router lib), CSS variables in a shadcn-style theme.
- **Backend:** Supabase (Postgres, Auth, Realtime, Storage, RLS).
- **Exports:** jsPDF + jspdf-autotable (PDF); hand-rolled CSV.
- **Email:** Resend from a Vercel serverless function.
- **Hosting/CI:** Vercel (static build + serverless + cron in `vercel.json`); GitHub Actions for lint/test/migration checks.
- **Quality:** ESLint + Prettier, Vitest + React Testing Library, SQL RLS tests, Sentry (frontend + serverless).

## 11. Environments & Configuration

Two Supabase projects (**staging**, **production**) and matching Vercel environments (preview → staging Supabase; production → production Supabase).

| Variable | Scope | Notes |
|---|---|---|
| `VITE_SUPABASE_URL` | Client build | Per-environment Supabase URL |
| `VITE_SUPABASE_ANON_KEY` | Client build | Anon key (RLS-protected) |
| `VITE_SENTRY_DSN` | Client build | Optional in staging |
| `SUPABASE_URL` | Serverless only | Same project as client |
| `SUPABASE_SERVICE_ROLE_KEY` | Serverless only | Never in client bundle |
| `RESEND_API_KEY` | Serverless only | **Moved out of client (was `VITE_RESEND_API_KEY` in v1 — security fix)** |
| `REPORT_RECIPIENTS` | Serverless only | Comma-separated emails |
| `CRON_SECRET` | Serverless only | Vercel-generated; also used for manual report re-runs |

## 12. Quality, Operations & Deployment (new in v2)

- **Migrations:** every schema change is a timestamped file in `supabase/migrations/`, applied with `supabase db push` (staging first, then production). No dashboard schema edits.
- **CI (GitHub Actions):** on every PR — ESLint, Prettier check, Vitest, `supabase db start` + apply migrations + run SQL RLS tests. Vercel builds previews per PR against staging Supabase.
- **Testing:** Vitest + Testing Library for form math, validation, and routing logic; SQL tests asserting each RLS policy (supervisor cross-project denial, creator-only edit, lock expiry, admin void-only update); a manual staging smoke checklist before each production deploy.
- **Monitoring:** Sentry on frontend and the cron function; Vercel cron execution logs checked by the daily email itself (missing email = alarm); Supabase dashboard alerts for DB health.
- **Backups:** Supabase automated daily backups (Pro plan) or a scheduled `pg_dump` GitHub Action to private storage on the free tier — decided at deploy time by plan tier; the development plan includes both paths.
- **Deployment:** documented runbook (in the development plan) covering Supabase project creation, migration push, seed data, Vercel project setup, env vars, custom domain, cron verification, and first-admin bootstrap.
