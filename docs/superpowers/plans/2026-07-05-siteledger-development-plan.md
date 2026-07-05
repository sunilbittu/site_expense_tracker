# SiteLedger Web App — Master Development Plan

> **For agentic workers:** This is a **master plan** — the product spans multiple subsystems, so each phase below is executed as its own detailed implementation plan. At the start of each phase, use superpowers:writing-plans to expand that phase into bite-sized TDD tasks, then superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement it. Phase checkboxes (`- [ ]`) track phase-level completion.

**Goal:** Build and deploy SiteLedger v1.0 per `docs/superpowers/specs/2026-07-05-siteledger-prd-v2.md` — a mobile-first React + Supabase expense tracker for construction sites, hosted on Vercel with a daily email report.

**Architecture:** SPA (React 18 + Vite) talking directly to Supabase (Postgres/Auth/Realtime/Storage) with all authorization in RLS; one Vercel serverless function for the daily Resend email, triggered by Vercel cron. Two environments (staging/production), schema managed exclusively by Supabase CLI migrations.

**Tech Stack:** React 18, Vite 6, Supabase JS v2, jsPDF + jspdf-autotable, Resend, Vercel, ESLint + Prettier, Vitest + React Testing Library, Sentry, GitHub Actions.

## Global Constraints

- Spec of record: `docs/superpowers/specs/2026-07-05-siteledger-prd-v2.md`. Where this plan and the spec disagree, the spec wins.
- **New repository `siteledger`** (this repo stays the Excel tool). If the user prefers a subfolder here instead, set Vercel "Root Directory" to it — everything else is identical.
- All schema changes via files in `supabase/migrations/` applied with `supabase db push`. Never edit schema in the Supabase dashboard.
- Server-secret env vars (`SUPABASE_SERVICE_ROLE_KEY`, `RESEND_API_KEY`, `CRON_SECRET`, `REPORT_RECIPIENTS`) are configured only in Vercel serverless scope — never with a `VITE_` prefix.
- All date-boundary logic uses `Asia/Kolkata`; `Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata' })` is the single client helper for "IST today" (`src/lib/istDate.js`); SQL uses `(now() at time zone 'Asia/Kolkata')::date`.
- Currency formatting: `new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 })` in one shared helper.
- Entry `id` is a client-generated UUID v4 (`crypto.randomUUID()`); inserts treat a 23505 primary-key conflict as success (idempotent retry).
- Every phase ends green: lint + Vitest + RLS tests pass in CI before the next phase starts. TDD within each phase's detailed plan.
- No routing library, no state library, no CSS framework — manual `route` state machine, React context, CSS variables (per spec §10). Resist adding dependencies.

## File Structure (target)

```
siteledger/
├── index.html
├── package.json
├── vite.config.js            # + vitest config
├── vercel.json               # cron schedule + SPA rewrites
├── .github/workflows/ci.yml
├── api/
│   └── daily-report.js       # Vercel serverless fn (cron + manual re-run)
├── supabase/
│   ├── config.toml
│   ├── migrations/
│   │   ├── 001_initial_schema.sql      # tables, enums, indexes, triggers
│   │   ├── 002_rls_policies.sql        # all RLS + is_admin()
│   │   ├── 003_reference_seed.sql      # worker types, work types (idempotent)
│   │   └── 004_storage_receipts.sql    # bucket + storage RLS
│   └── tests/rls.test.sql              # RLS assertions (run in CI)
├── src/
│   ├── main.jsx
│   ├── App.jsx               # route state machine + role branch
│   ├── lib/
│   │   ├── supabase.js       # client singleton
│   │   ├── istDate.js        # IST today/month helpers
│   │   ├── money.js          # INR formatting
│   │   ├── categories.js     # category defs, detail schemas, total formulas
│   │   └── csv.js            # CSV export builder
│   ├── context/DataContext.jsx
│   ├── screens/
│   │   ├── Login.jsx  SetPassword.jsx  ForgotPassword.jsx
│   │   ├── ProjectPicker.jsx  Home.jsx  History.jsx  EntryDetail.jsx
│   │   ├── AddExpense.jsx    # 2-step sheet, delegates to forms/
│   │   └── AdminDashboard.jsx
│   ├── forms/                # one file per category form
│   │   ├── NmrForm.jsx  JcbForm.jsx  ContractorFeeForm.jsx
│   │   ├── LaborContractForm.jsx  GeneralForm.jsx
│   ├── components/           # shared UI (Sheet, AmountInput, PhotoCapture, …)
│   └── pdf/report.js         # jsPDF report builder
└── tests/                    # Vitest specs mirroring src/
```

---

## Phase 0: Repo scaffold, tooling, environments — `- [ ]`

**Deliverable:** empty-but-deployed app; CI green; both Supabase projects exist.

1. `npm create vite@latest siteledger -- --template react`; add ESLint (flat config, react hooks plugin), Prettier, Vitest + `@testing-library/react` + jsdom; one placeholder test; scripts: `dev`, `build`, `lint`, `test`.
2. Create GitHub repo; add `.github/workflows/ci.yml`:

```yaml
name: CI
on: [push, pull_request]
jobs:
  ci:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: 22, cache: npm }
      - run: npm ci
      - run: npm run lint
      - run: npx prettier --check .
      - run: npm test -- --run
  rls:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: supabase/setup-cli@v1
      - run: supabase db start
      - run: supabase db reset          # applies all migrations
      - run: psql "$(supabase status -o json | jq -r '.DB_URL // .services.db.url')" -v ON_ERROR_STOP=1 -f supabase/tests/rls.test.sql
```

3. Create **two** Supabase projects (`siteledger-staging`, `siteledger-prod`); `supabase init` + `supabase link` to staging locally. Record both project refs in the repo README.
4. Create Vercel project from the GitHub repo. Set env vars: Preview → staging Supabase keys, Production → prod keys (client vars only for now). Confirm a hello-world deploy on both.
5. Add Sentry (`@sentry/react`) behind `VITE_SENTRY_DSN` (no-op when unset).

**Exit criteria:** PR merge triggers CI (green) and a Vercel preview; production URL serves the placeholder app.

## Phase 1: Database schema, RLS, migrations, RLS tests — `- [ ]`

**Deliverable:** full schema live on staging; every RLS rule proven by a SQL test in CI.

Migration `001_initial_schema.sql` — key definitions (complete enums/tables per spec §8):

```sql
create type user_role as enum ('supervisor', 'admin');
create type entry_category as enum ('nmr', 'jcb', 'contractor_fee', 'labor_contract', 'general');
create type payment_mode as enum ('cash', 'upi', 'bank_transfer');
create type contractor_type as enum ('daily_fee', 'labor_contract', 'both');

-- profiles auto-created from auth.users by trigger (security definer)
create table profiles (
  id uuid primary key references auth.users on delete cascade,
  name text not null default '',
  email text not null,
  role user_role not null default 'supervisor',
  must_change_password boolean not null default true
);

create table entries (
  id uuid primary key,                          -- client-generated (idempotent retry)
  project_id uuid not null references projects,
  plot_id uuid not null references plots,
  category entry_category not null,
  total_amount numeric(12,2) not null check (total_amount > 0),
  payment_mode payment_mode not null,
  receipt_path text,
  notes text,
  entry_date date not null,
  details jsonb not null default '{}',
  created_by uuid not null references profiles default auth.uid(),
  created_at timestamptz not null default now(),
  locked_at timestamptz not null,               -- trigger-set, never client-set
  voided_at timestamptz, voided_by uuid references profiles, void_reason text,
  check ((voided_at is null) = (voided_by is null) and (voided_at is null) = (void_reason is null))
);
create index on entries (project_id, entry_date);
create index on entries (entry_date);

create or replace function set_entry_defaults() returns trigger language plpgsql as $$
begin
  new.created_by := auth.uid();
  new.created_at := now();
  new.locked_at  := now() + interval '24 hours';
  new.voided_at := null; new.voided_by := null; new.void_reason := null;
  if new.entry_date not between ((now() at time zone 'Asia/Kolkata')::date - 1)
                            and  (now() at time zone 'Asia/Kolkata')::date then
    raise exception 'entry_date must be today or yesterday (IST)';
  end if;
  return new;
end $$;
create trigger entries_defaults before insert on entries
  for each row execute function set_entry_defaults();

-- audit: snapshot prior row on every UPDATE/DELETE
create table entries_audit (
  id bigint generated always as identity primary key,
  entry_id uuid not null, action text not null,
  actor uuid, at timestamptz not null default now(),
  prior jsonb not null
);
create or replace function audit_entries() returns trigger language plpgsql security definer as $$
begin
  insert into entries_audit (entry_id, action, actor, prior)
  values (old.id, tg_op, auth.uid(), to_jsonb(old));
  return coalesce(new, old);
end $$;
create trigger entries_audit_trg after update or delete on entries
  for each row execute function audit_entries();

-- month-to-date per project, IST month (replaces v1's stored aggregate)
create view project_month_totals with (security_invoker = true) as
  select project_id, sum(total_amount) as mtd_total
  from entries
  where voided_at is null
    and date_trunc('month', entry_date) = date_trunc('month', (now() at time zone 'Asia/Kolkata')::date)
  group by project_id;
```

Migration `002_rls_policies.sql` — the load-bearing policies:

```sql
create or replace function is_admin() returns boolean language sql stable security definer as
  $$ select exists (select 1 from profiles where id = auth.uid() and role = 'admin') $$;

create or replace function is_assigned(p uuid) returns boolean language sql stable security definer as
  $$ select exists (select 1 from project_assignments where project_id = p and user_id = auth.uid()) $$;

alter table entries enable row level security;

create policy entries_select on entries for select
  using (is_admin() or is_assigned(project_id));

create policy entries_insert on entries for insert
  with check (not is_admin() and is_assigned(project_id) and created_by = auth.uid());

-- creator edit/delete inside the unlock window only
create policy entries_update_creator on entries for update
  using (created_by = auth.uid() and locked_at > now() and voided_at is null)
  with check (created_by = auth.uid());

create policy entries_delete_creator on entries for delete
  using (created_by = auth.uid() and locked_at > now() and voided_at is null);

-- admin void: UPDATE allowed, but a trigger rejects changes to non-void columns
create policy entries_void_admin on entries for update
  using (is_admin()) with check (is_admin());

create or replace function enforce_void_only() returns trigger language plpgsql as $$
begin
  if is_admin() and new.created_by = old.created_by then
    if (to_jsonb(new) - 'voided_at' - 'voided_by' - 'void_reason')
       is distinct from (to_jsonb(old) - 'voided_at' - 'voided_by' - 'void_reason') then
      raise exception 'admin may only set void fields';
    end if;
    if new.voided_at is not null and old.voided_at is null then
      if length(coalesce(new.void_reason, '')) < 10 then
        raise exception 'void_reason must be at least 10 characters';
      end if;
      new.voided_by := auth.uid(); new.voided_at := now();
    elsif old.voided_at is not null then
      raise exception 'void is permanent';
    end if;
  end if;
  return new;
end $$;
create trigger entries_void_guard before update on entries
  for each row execute function enforce_void_only();
```

`supabase/tests/rls.test.sql` must assert at minimum (using `set local role authenticated; set local request.jwt.claims ...` per test):
supervisor sees only assigned projects' entries; supervisor cannot insert into unassigned project; creator can update before `locked_at`, cannot after (simulate by updating `locked_at` as service role); non-creator supervisor on same project cannot update; admin cannot insert; admin cannot change `total_amount`; admin void succeeds with reason and fails with short reason; second void fails; audit row written on update/delete; `project_month_totals` excludes voided entries.

**Exit criteria:** `supabase db reset` clean; all RLS tests pass locally and in CI; `supabase db push` applied to staging.

## Phase 2: Auth & app shell — `- [ ]`

**Deliverable:** login → forced password change → role-branched routing, with DataContext loading reference data.

- `src/lib/supabase.js` singleton; `DataContext` loads profile, projects (admin: all; supervisor: via assignments), plots, reference tables once per session.
- Screens: `Login` (signInWithPassword + error states), `ForgotPassword` (resetPasswordForEmail + Supabase email template configured), `SetPassword` (8+ chars, ≥1 digit, confirm match; clears `must_change_password`).
- `App.jsx` route state machine: unauthenticated → Login; `must_change_password` → SetPassword; admin → AdminDashboard; supervisor → ProjectPicker (auto-skip to Home when exactly one assignment).
- Vitest: password validator, routing branch logic (mock session/profile), DataContext load shaping.

**Exit criteria:** on staging, a seeded supervisor and admin can each log in and land on the right (still-stub) screen; forced-change and forgot-password flows work end-to-end.

## Phase 3: Supervisor core — picker, home, add-expense, detail, history — `- [ ]`

**Deliverable:** the entire §4.2 supervisor flow, minus photos.

- `src/lib/categories.js`: per-category detail schema, validation, and total formula (pure functions — fully unit-tested: NMR count×wage, JCB hours≥0.5×rate, fee+tip, amount passthroughs; rejection of zero/negative amounts).
- ProjectPicker with `project_month_totals` join; Home with IST-today query, category grouping/totals/filter, Realtime subscription (private channel, RLS-authorized) with unsubscribe on project switch; AddExpense 2-step sheet (category ranking by today's usage, "Repeat last" prefill, plot + payment mode + notes, live total, UUID id, 23505-as-success, in-flight disable, backdate-by-1 date picker); EntryDetail with edit-window countdown, creator-only edit/delete while unlocked; History grouped by date desc with per-date category subtotals, 30-day pagination.
- Vitest: category math, ranking, repeat-last prefill, lock countdown boundaries, reducer/grouping logic. Realtime handler tested as a pure reducer (event in → list out).

**Exit criteria:** staging smoke: two supervisor accounts on different projects cannot see each other's data; entry add/edit/delete round-trips live on two devices.

## Phase 4: Receipt photos — `- [ ]`

**Deliverable:** capture → compress → upload → signed-URL display.

- Migration `004_storage_receipts.sql`: private `receipts` bucket; storage.objects RLS: INSERT by entry creator while entry unlocked (path `{project_id}/{entry_id}.jpg` parsed with `storage.foldername`), SELECT for assigned members + admin.
- `PhotoCapture` component: `<input type="file" accept="image/*" capture="environment">`, canvas resize to 1600 px / JPEG q≈0.7; upload after entry insert; `receipt_path` saved on entry; failure leaves entry intact with retry from EntryDetail.
- Display in EntryDetail and AdminDashboard via `createSignedUrl` (5-min TTL).

**Exit criteria:** photo taken on a phone appears for admin; unassigned supervisor's signed-URL attempt fails (RLS test added).

## Phase 5: Admin dashboard, void, PDF + CSV export — `- [ ]`

**Deliverable:** the entire §4.3 admin flow.

- Date-range picker (IST defaults), summary card (grand total excl. voided, count, stacked category bar), date→project grouping with category breakdowns, struck-through voided entries with reason.
- Void action: reason dialog (min 10 chars) → UPDATE of void fields → refetch.
- `src/pdf/report.js` (jsPDF/autotable per spec §7.1) and `src/lib/csv.js` (UTF-8 BOM, one row per entry incl. void columns). Both pure functions over the fetched entry list — unit-tested for subtotal math and CSV escaping (commas, quotes, newlines in notes).

**Exit criteria:** staging: admin voids a locked entry (reason enforced), totals update, exports open correctly in a PDF viewer and Excel.

## Phase 6: Daily email cron — `- [ ]`

**Deliverable:** 7 PM IST summary email, re-runnable for any date.

`vercel.json`:

```json
{
  "rewrites": [{ "source": "/((?!api/).*)", "destination": "/index.html" }],
  "crons": [{ "path": "/api/daily-report", "schedule": "30 13 * * *" }]
}
```

`api/daily-report.js` skeleton (Node runtime):

```js
import { createClient } from '@supabase/supabase-js';
import { Resend } from 'resend';

export default async function handler(req, res) {
  if (req.headers.authorization !== `Bearer ${process.env.CRON_SECRET}`)
    return res.status(401).json({ error: 'unauthorized' });

  const istToday = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata' }).format(new Date());
  const date = /^\d{4}-\d{2}-\d{2}$/.test(req.query.date ?? '') ? req.query.date : istToday;

  const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);
  // fetch day's entries + month-to-date (voided excluded), group per project/category,
  // build HTML, then:
  const resend = new Resend(process.env.RESEND_API_KEY);
  const to = process.env.REPORT_RECIPIENTS.split(',').map((s) => s.trim());
  const { error } = await resend.emails.send({ from: 'SiteLedger <reports@YOURDOMAIN>', to, subject, html });
  if (error) return res.status(502).json({ error: error.message });
  return res.status(200).json({ date, sent: to.length });
}
```

- Vercel cron sends `Authorization: Bearer ${CRON_SECRET}` automatically when `CRON_SECRET` is set — the same check covers manual re-runs: `curl -H "Authorization: Bearer $CRON_SECRET" "https://<prod>/api/daily-report?date=2026-07-04"`.
- Resend domain verification (SPF/DKIM) for the from-address, or use Resend's onboarding domain in staging.
- Unit-test the aggregation/formatting as pure functions; hit the endpoint manually on staging preview before enabling the prod cron.

**Exit criteria:** staging manual trigger delivers a correct email for a seeded day; missing/wrong secret → 401; Resend failure → 502 visible in Vercel logs.

## Phase 7: Deployment hardening & production launch — `- [ ]`

**Deliverable:** production live with monitoring, backups, and runbooks.

**Deployment runbook (execute in order):**

1. **Production Supabase:** `supabase link --project-ref <prod-ref>` → `supabase db push` → verify tables/policies via `supabase db diff` (empty). Apply `003_reference_seed.sql` data. Configure Auth: disable public signups, set Site URL to prod domain, customize reset-password email template.
2. **Bootstrap admin:** create the admin user in Supabase Auth dashboard → set `profiles.role='admin'`, `must_change_password=true`. Create supervisor users + `project_assignments` rows per the **user-management runbook** (a `docs/runbooks/users.md` with exact SQL snippets — written in this phase).
3. **Projects/plots/reference data:** insert real projects, plots (incl. "Site-wide" per project), contractors, JCB operators via SQL snippets in the runbook.
4. **Vercel production env:** `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY`, `VITE_SENTRY_DSN` (client); `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`, `RESEND_API_KEY`, `REPORT_RECIPIENTS`, `CRON_SECRET` (serverless). Confirm no `VITE_` variable contains a secret (`grep VITE_ .env* + Vercel dashboard review`).
5. **Domain & delivery:** attach custom domain in Vercel (HTTPS automatic); verify Resend domain DNS (SPF, DKIM) and switch from-address to it.
6. **Cron:** confirm the cron appears in Vercel → Settings → Cron Jobs after deploy; trigger one manual run with `?date=<yesterday>`; verify email received.
7. **Backups:** Supabase Pro → enable PITR/daily backups; free tier → add `.github/workflows/backup.yml` running nightly `pg_dump` (connection string in GitHub secret) uploading to a private artifact/bucket, 30-day retention. Do a **restore drill** once: restore dump into a scratch database and row-count `entries`.
8. **Monitoring:** Sentry DSNs set for prod (frontend + serverless); Vercel log retention checked; the daily email doubles as a heartbeat — "no 7 PM email" is the documented alarm, with the manual re-run command in the runbook.
9. **Staging smoke checklist** (run before every prod deploy, documented in `docs/runbooks/release.md`): login both roles, forced password change, add one entry per category, edit + delete within window, realtime on second device, photo upload + admin view, admin void, PDF + CSV export, manual report trigger.
10. **Launch:** production deploy from `main`, run checklist steps 1–2 against prod with real accounts, hand supervisors the URL + first-login instructions (in `docs/runbooks/onboarding.md`).

**Exit criteria:** production URL live on custom domain; first real daily email received at 7 PM IST; backup + restore drill done; runbooks committed.

---

## Ordering, risk & rollback

- Phases are strictly sequential except Phase 4 (photos) and Phase 5 (admin) which can proceed in parallel after Phase 3.
- **Highest-risk items are front-loaded:** RLS correctness (Phase 1, SQL-tested in CI) and the void-only trigger — everything else is conventional UI work.
- Rollback: Vercel instant rollback to a previous deployment for frontend issues; migrations are forward-only — every migration must be additive or ship with a companion down-path noted in its header comment; schema changes deploy to staging ≥1 day before prod.
- Estimated effort: Phases 0–2 ≈ 1 week, Phase 3 ≈ 1–1.5 weeks, Phases 4–6 ≈ 1 week combined, Phase 7 ≈ 2–3 days (calendar, single developer + agent assistance).

## Open decisions (defaults chosen; change before Phase 1 if desired)

1. **Backdating window** — defaulted to 1 day (yesterday). Widen the CHECK in `set_entry_defaults` if supervisors log later than that.
2. **Supabase plan tier** — free tier works for v1.0 scale, but backups then rely on the `pg_dump` Action, and paused-on-inactivity applies; Pro removes both concerns.
3. **Repo location** — defaulted to a new `siteledger` repo; subfolder-of-this-repo also works (set Vercel root directory).
4. **Void visibility to supervisors** — defaulted to visible (struck-through) on Home/History so site staff know a correction happened.
