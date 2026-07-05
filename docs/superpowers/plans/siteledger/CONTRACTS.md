# SiteLedger — Shared Implementation Contracts

Single source of truth for names, signatures, schema, and conventions used by ALL phase plans (`phase-0` … `phase-7`). Phase plans must use these **verbatim** — never rename, never invent parallel helpers. If a phase plan and this file disagree, this file wins.

Spec of record: `docs/superpowers/specs/2026-07-05-siteledger-prd-v2.md`
Master plan: `docs/superpowers/plans/2026-07-05-siteledger-development-plan.md`

## 1. Repository & file tree

New repo `siteledger` (target tree — phases create these files and no others unless a phase plan says so):

```
siteledger/
├── index.html
├── package.json
├── vite.config.js              # vite + vitest config in one file
├── vercel.json
├── .env.example
├── .github/workflows/ci.yml
├── api/daily-report.js
├── supabase/
│   ├── config.toml
│   ├── migrations/
│   │   ├── 20260706000001_initial_schema.sql
│   │   ├── 20260706000002_rls_policies.sql
│   │   ├── 20260706000003_reference_seed.sql
│   │   └── 20260706000004_storage_receipts.sql
│   └── tests/rls.test.sql
├── scripts/test-rls.sh
├── src/
│   ├── main.jsx
│   ├── App.jsx
│   ├── styles.css
│   ├── lib/
│   │   ├── supabase.js  istDate.js  money.js  categories.js
│   │   ├── entriesStore.js  reportData.js  csv.js  receipts.js
│   ├── pdf/report.js
│   ├── context/DataContext.jsx
│   ├── screens/
│   │   ├── Login.jsx  SetPassword.jsx  ForgotPassword.jsx
│   │   ├── ProjectPicker.jsx  Home.jsx  AddExpense.jsx
│   │   ├── EntryDetail.jsx  History.jsx  AdminDashboard.jsx
│   ├── forms/
│   │   ├── NmrForm.jsx  JcbForm.jsx  ContractorFeeForm.jsx
│   │   ├── LaborContractForm.jsx  GeneralForm.jsx
│   └── components/
│       ├── Sheet.jsx  AmountInput.jsx  PhotoCapture.jsx
├── tests/                      # Vitest specs, mirroring src/ (tests/lib/categories.test.js etc.)
└── docs/runbooks/              # users.md  release.md  onboarding.md (Phase 7)
```

## 2. Dependencies (package.json)

Runtime: `react ^18.3`, `react-dom ^18.3`, `@supabase/supabase-js ^2.45`, `jspdf ^3.0`, `jspdf-autotable ^5.0`, `@sentry/react ^9` . Serverless only: `resend ^4` (regular dependency; used by `api/`).
Dev: `vite ^6`, `@vitejs/plugin-react ^4`, `vitest ^3`, `jsdom ^26`, `@testing-library/react ^16`, `@testing-library/jest-dom ^6`, `eslint ^9` (flat config) + `eslint-plugin-react-hooks ^5` + `eslint-plugin-react-refresh`, `prettier ^3` (default settings).

Scripts: `dev`, `build`, `preview`, `lint` (`eslint .`), `format:check` (`prettier --check .`), `test` (`vitest`), `test:rls` (`bash scripts/test-rls.sh`).

## 3. Environment variables

| Variable | Scope |
|---|---|
| `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY`, `VITE_SENTRY_DSN` (optional) | client build |
| `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`, `RESEND_API_KEY`, `REPORT_RECIPIENTS`, `CRON_SECRET` | Vercel serverless only |

Never a secret behind a `VITE_` prefix. `.env.example` lists all with placeholder values.

Steps a human must do in a dashboard (Supabase/Vercel/Resend/DNS) are marked **[MANUAL]** in phase plans, with exact click-paths.

## 4. Database schema (verbatim — Phase 1 ships these migrations; all other phases read column names from here)

### 4.1 `20260706000001_initial_schema.sql`

```sql
create extension if not exists pgcrypto;

create type user_role as enum ('supervisor', 'admin');
create type entry_category as enum ('nmr', 'jcb', 'contractor_fee', 'labor_contract', 'general');
create type payment_mode as enum ('cash', 'upi', 'bank_transfer');
create type contractor_type as enum ('daily_fee', 'labor_contract', 'both');

create table profiles (
  id uuid primary key references auth.users on delete cascade,
  name text not null default '',
  email text not null,
  role user_role not null default 'supervisor',
  must_change_password boolean not null default true
);

create table projects (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  location text not null default '',
  active boolean not null default true
);

create table plots (
  id uuid primary key default gen_random_uuid(),
  project_id uuid not null references projects on delete cascade,
  name text not null
);

create table project_assignments (
  project_id uuid not null references projects on delete cascade,
  user_id uuid not null references profiles on delete cascade,
  primary key (project_id, user_id)
);

create table worker_types (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  default_wage numeric(10,2) not null check (default_wage > 0),
  default_wage_female numeric(10,2) not null check (default_wage_female > 0)
);

create table contractors (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  type contractor_type not null default 'both'
);

create table jcb_operators (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  default_hourly_rate numeric(10,2) not null check (default_hourly_rate > 0)
);

create table work_types (
  id uuid primary key default gen_random_uuid(),
  name text not null unique
);

create table entries (
  id uuid primary key,                          -- client-generated UUID v4 (idempotent retry)
  project_id uuid not null references projects,
  plot_id uuid not null references plots,
  category entry_category not null,
  total_amount numeric(12,2) not null check (total_amount > 0),
  payment_mode payment_mode not null,
  receipt_path text,
  notes text not null default '',
  entry_date date not null,
  details jsonb not null default '{}',
  created_by uuid not null references profiles,
  created_at timestamptz not null default now(),
  locked_at timestamptz not null,               -- trigger-set, never client-set
  voided_at timestamptz,
  voided_by uuid references profiles,
  void_reason text,
  check ((voided_at is null) = (voided_by is null)),
  check ((voided_at is null) = (void_reason is null))
);
create index entries_project_date_idx on entries (project_id, entry_date);
create index entries_date_idx on entries (entry_date);

create table entries_audit (
  id bigint generated always as identity primary key,
  entry_id uuid not null,
  action text not null,
  actor uuid,
  at timestamptz not null default now(),
  prior jsonb not null
);

-- profile auto-creation from auth.users
create or replace function public.handle_new_user() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into public.profiles (id, name, email)
  values (new.id, coalesce(new.raw_user_meta_data->>'name', ''), new.email);
  return new;
end $$;
create trigger on_auth_user_created after insert on auth.users
  for each row execute function public.handle_new_user();

create or replace function ist_today() returns date language sql stable as
  $$ select (now() at time zone 'Asia/Kolkata')::date $$;

-- server-set system fields; entry_date window (today or yesterday, IST)
create or replace function set_entry_defaults() returns trigger language plpgsql as $$
begin
  new.created_by := coalesce(auth.uid(), new.created_by);
  new.created_at := now();
  new.locked_at  := now() + interval '24 hours';
  new.voided_at := null; new.voided_by := null; new.void_reason := null;
  if new.entry_date not between (ist_today() - 1) and ist_today() then
    raise exception 'entry_date must be today or yesterday (IST)';
  end if;
  return new;
end $$;
create trigger entries_defaults before insert on entries
  for each row execute function set_entry_defaults();

-- audit snapshot of prior row on every UPDATE/DELETE
create or replace function audit_entries() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into entries_audit (entry_id, action, actor, prior)
  values (old.id, tg_op, auth.uid(), to_jsonb(old));
  return coalesce(new, old);
end $$;
create trigger entries_audit_trg after update or delete on entries
  for each row execute function audit_entries();

-- update guard: admin may only void (once, with reason >= 10 chars);
-- creator edits cannot touch system/void fields or move projects
create or replace function enforce_entry_update_rules() returns trigger language plpgsql as $$
begin
  if auth.uid() is null then
    return new;  -- service role (seeds, test setup) unrestricted
  end if;
  if is_admin() then
    if (to_jsonb(new) - 'voided_at' - 'voided_by' - 'void_reason')
       is distinct from (to_jsonb(old) - 'voided_at' - 'voided_by' - 'void_reason') then
      raise exception 'admin may only set void fields';
    end if;
    if old.voided_at is not null then
      raise exception 'void is permanent';
    end if;
    if new.voided_at is null then
      raise exception 'admin update must void the entry';
    end if;
    if length(coalesce(new.void_reason, '')) < 10 then
      raise exception 'void_reason must be at least 10 characters';
    end if;
    new.voided_at := now();
    new.voided_by := auth.uid();
  else
    if new.voided_at is not null or new.voided_by is not null or new.void_reason is not null then
      raise exception 'only admin may void';
    end if;
    if new.project_id is distinct from old.project_id then
      raise exception 'entry cannot move between projects';
    end if;
    if new.entry_date not between (ist_today() - 1) and ist_today() then
      raise exception 'entry_date must be today or yesterday (IST)';
    end if;
    new.created_by := old.created_by;
    new.created_at := old.created_at;
    new.locked_at  := old.locked_at;
  end if;
  return new;
end $$;
create trigger entries_update_guard before update on entries
  for each row execute function enforce_entry_update_rules();
```

Note: `enforce_entry_update_rules` references `is_admin()`, defined in migration 2 — migration 1 must therefore also contain a forward declaration OR the trigger creation moves to migration 2. **Resolution (binding):** move the `create trigger entries_update_guard` statement and its function into migration 2, after `is_admin()` is defined. Phase 1's plan reflects this.

### 4.2 `20260706000002_rls_policies.sql`

```sql
create or replace function is_admin() returns boolean
language sql stable security definer set search_path = public as
  $$ select exists (select 1 from profiles where id = auth.uid() and role = 'admin') $$;

create or replace function is_assigned(p uuid) returns boolean
language sql stable security definer set search_path = public as
  $$ select exists (select 1 from project_assignments where project_id = p and user_id = auth.uid()) $$;

-- (enforce_entry_update_rules function + entries_update_guard trigger created here — see note above)

alter table profiles enable row level security;
alter table projects enable row level security;
alter table plots enable row level security;
alter table project_assignments enable row level security;
alter table worker_types enable row level security;
alter table contractors enable row level security;
alter table jcb_operators enable row level security;
alter table work_types enable row level security;
alter table entries enable row level security;
alter table entries_audit enable row level security;

create policy profiles_select on profiles for select using (id = auth.uid() or is_admin());
create policy profiles_update_self on profiles for update
  using (id = auth.uid()) with check (id = auth.uid());

create or replace function enforce_profile_guard() returns trigger language plpgsql as $$
begin
  if auth.uid() is not null and not is_admin() and new.role is distinct from old.role then
    raise exception 'role can only be changed by an admin';
  end if;
  return new;
end $$;
create trigger profiles_update_guard before update on profiles
  for each row execute function enforce_profile_guard();

create policy projects_select on projects for select using (is_admin() or is_assigned(id));
create policy plots_select on plots for select using (is_admin() or is_assigned(project_id));
create policy assignments_select on project_assignments for select
  using (user_id = auth.uid() or is_admin());

create policy worker_types_select on worker_types for select to authenticated using (true);
create policy contractors_select on contractors for select to authenticated using (true);
create policy jcb_operators_select on jcb_operators for select to authenticated using (true);
create policy work_types_select on work_types for select to authenticated using (true);

create policy entries_select on entries for select using (is_admin() or is_assigned(project_id));
create policy entries_insert on entries for insert
  with check (not is_admin() and is_assigned(project_id) and created_by = auth.uid());
create policy entries_update_creator on entries for update
  using (created_by = auth.uid() and locked_at > now() and voided_at is null)
  with check (created_by = auth.uid());
create policy entries_delete_creator on entries for delete
  using (created_by = auth.uid() and locked_at > now() and voided_at is null);
create policy entries_void_admin on entries for update
  using (is_admin()) with check (is_admin());

create policy entries_audit_select on entries_audit for select using (is_admin());

create view project_month_totals with (security_invoker = true) as
  select project_id, sum(total_amount) as mtd_total
  from entries
  where voided_at is null
    and date_trunc('month', entry_date) = date_trunc('month', ist_today())
  group by project_id;

alter publication supabase_realtime add table entries;
```

### 4.3 `20260706000003_reference_seed.sql` (idempotent)

```sql
insert into worker_types (name, default_wage, default_wage_female) values
  ('Mason', 900, 700), ('Helper', 600, 400), ('Carpenter', 950, 750), ('Bar Bender', 900, 700)
on conflict (name) do nothing;

insert into work_types (name) values
  ('Plastering'), ('Brickwork'), ('Concrete'), ('Flooring'), ('Painting'), ('Other')
on conflict (name) do nothing;
```

(Contractors, JCB operators, projects, plots are real data — inserted at deploy time per the Phase 7 users runbook, not seeded.)

### 4.4 `20260706000004_storage_receipts.sql`

Object path convention: `{project_id}/{entry_id}.jpg` in private bucket `receipts`.

```sql
insert into storage.buckets (id, name, public) values ('receipts', 'receipts', false)
on conflict (id) do nothing;

create policy receipts_insert on storage.objects for insert to authenticated
  with check (
    bucket_id = 'receipts'
    and exists (
      select 1 from public.entries e
      where e.project_id::text = (storage.foldername(name))[1]
        and (e.id::text || '.jpg') = storage.filename(name)
        and e.created_by = auth.uid()
        and e.locked_at > now()
    )
  );

create policy receipts_update on storage.objects for update to authenticated
  using (
    bucket_id = 'receipts'
    and exists (
      select 1 from public.entries e
      where e.project_id::text = (storage.foldername(name))[1]
        and (e.id::text || '.jpg') = storage.filename(name)
        and e.created_by = auth.uid()
        and e.locked_at > now()
    )
  );

create policy receipts_select on storage.objects for select to authenticated
  using (
    bucket_id = 'receipts'
    and (public.is_admin() or public.is_assigned(((storage.foldername(name))[1])::uuid))
  );
```

## 5. RLS test conventions (`supabase/tests/rls.test.sql`)

Runs via `scripts/test-rls.sh` against the local stack (`supabase db start` + `supabase db reset` first). Impersonation pattern:

```sql
-- seed as postgres (bypasses RLS):
insert into auth.users (id, email) values ('00000000-0000-0000-0000-000000000001', 'sup1@test.local');
-- (profiles auto-created by trigger; then e.g.)
update profiles set role = 'admin' where id = '...';

-- impersonate:
set local role authenticated;
set local request.jwt.claims to '{"sub":"00000000-0000-0000-0000-000000000001","role":"authenticated"}';

-- assert visible rows:
do $$ begin
  assert (select count(*) from entries) = 1, 'sup1 sees exactly their project entries';
end $$;

-- assert denial (RLS insert violation = SQLSTATE 42501):
do $$ begin
  begin
    insert into entries (id, project_id, ...) values (...);
    raise exception 'FAIL: insert should have been denied';
  exception when insufficient_privilege then null;
  end;
end $$;

-- back to superuser between sections:
reset role;
```

UPDATE/DELETE denials do not error — they affect 0 rows; assert by re-reading the row as postgres. Simulate lock expiry as postgres: `update entries set locked_at = now() - interval '1 hour' where id = ...;`. The whole file runs in one transaction (`begin;` … `rollback;`) so it never dirties the DB, and every `do` block failure aborts with non-zero exit (ON_ERROR_STOP).

`scripts/test-rls.sh`:

```bash
#!/usr/bin/env bash
set -euo pipefail
DB_URL=$(supabase status --output json | python3 -c 'import json,sys; d=json.load(sys.stdin); print(d.get("DB_URL") or d["services"]["db"]["url"])' 2>/dev/null || echo "postgresql://postgres:postgres@127.0.0.1:54322/postgres")
psql "$DB_URL" -v ON_ERROR_STOP=1 -f supabase/tests/rls.test.sql
echo "RLS tests passed"
```

## 6. JavaScript module contracts (exact exports)

### `src/lib/supabase.js`
```js
import { createClient } from '@supabase/supabase-js';
export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);
```

### `src/lib/istDate.js`
```js
export function istToday()            // 'YYYY-MM-DD' for Asia/Kolkata now
export function istDaysAgo(n)         // 'YYYY-MM-DD', n days before istToday()
export function formatDisplayDate(s)  // 'YYYY-MM-DD' -> '5 Jul 2026'
```
Implementation basis: `new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata' }).format(d)`; `istDaysAgo` subtracts `n * 86400000` ms from `Date.now()` before formatting (IST has no DST).

### `src/lib/money.js`
```js
export function formatINR(n) // ₹ with en-IN grouping, no decimals: 12345 -> '₹12,345'
```

### `src/lib/categories.js`
```js
export const CATEGORY_ORDER = ['nmr', 'jcb', 'contractor_fee', 'labor_contract', 'general'];
export const CATEGORY_LABELS = {
  nmr: 'NMR (Labor)', jcb: 'JCB', contractor_fee: 'Contractor Fee',
  labor_contract: 'Labor Contract', general: 'General',
};
export const PAYMENT_MODES = ['cash', 'upi', 'bank_transfer'];
export const PAYMENT_MODE_LABELS = { cash: 'Cash', upi: 'UPI', bank_transfer: 'Bank Transfer' };
export const PAYMENT_STAGES = ['advance', 'running', 'final', 'other'];

export function computeTotal(category, details)      // number, or null if incomplete/invalid
export function validateDetails(category, details)   // string[] of error messages; [] = valid
export function describeEntry(entry, refs)           // one-line description for lists/reports
export function rankCategories(todayEntries)         // all 5 keys, today-usage desc, ties by CATEGORY_ORDER
```
`refs = { workerTypesById, contractorsById, jcbOperatorsById, workTypesById }` (plain objects keyed by id).

`details` shapes (JSONB, snake_case — stored exactly as below):
- `nmr`: `{ worker_type_id, worker_gender: 'male'|'female', worker_count (int ≥1), wage_per_worker (>0), wage_overridden (bool) }` — total = `worker_count * wage_per_worker`
- `jcb`: `{ operator_id, hours (≥0.5, 0.5 steps), hourly_rate (>0) }` — total = `hours * hourly_rate`
- `contractor_fee`: `{ contractor_id, fee (>0), tip (≥0, optional) }` — total = `fee + (tip || 0)`
- `labor_contract`: `{ contractor_id, work_type_id, amount (>0), payment_stage }` — total = `amount`
- `general`: `{ description (non-empty), amount (>0) }` — total = `amount`

`describeEntry` outputs: nmr `'5 × Mason (male)'`; jcb `'Ramesh — 3.5 h'`; contractor_fee `'<contractor name>'` (+ `' + tip'` when tip > 0); labor_contract `'<contractor> — <work type> (<stage>)'`; general `details.description`.

### `src/lib/entriesStore.js` (pure functions — realtime + grouping logic lives here, unit-tested)
```js
export function applyRealtimeEvent(entries, payload) // payload = supabase postgres_changes payload; returns new array
export function groupByCategory(entries)             // { [category]: Entry[] } insertion order = CATEGORY_ORDER
export function groupByDate(entries)                 // [['YYYY-MM-DD', Entry[]], ...] date desc
export function categoryTotals(entries)              // { [category]: number } voided excluded
export function activeTotal(entries)                 // sum of total_amount, voided excluded
```

### `src/lib/reportData.js`
```js
export function buildReportRows(entries, ctx)
// ctx = { projectsById, plotsById, refs }  (refs as in categories.js)
// -> [{ date, project, plot, category (label), description, paymentMode (label),
//       amount (number), voided (bool), voidReason (string) }] sorted date asc, then created_at asc
export function summarize(rows)
// -> { grandTotal, count, byCategory: {label: amount}, byDay: {date: amount} }  voided excluded everywhere
```

### `src/lib/csv.js`
```js
export function toCsv(rows)  // reportData rows -> CSV string, UTF-8 BOM prefix '﻿'
// header: Date,Project,Plot,Category,Description,Payment Mode,Amount,Voided,Void Reason
// RFC 4180 quoting (fields containing , " or newline get quoted, inner " doubled)
export function downloadFile(filename, text, mime)   // Blob + object URL + <a download> click
```

### `src/pdf/report.js`
```js
export function buildReportPdf(rows, summary, range) // rows/summary from reportData; range = {start, end}
// -> jsPDF instance (caller calls .save(`SiteLedger-Report-${start}-to-${end}.pdf`))
```

### `src/lib/receipts.js`
```js
export async function compressImage(file)                    // File -> JPEG Blob, max 1600px long edge, quality 0.7
export async function uploadReceipt(projectId, entryId, blob) // upload w/ upsert:true -> returns path `${projectId}/${entryId}.jpg`
export async function getReceiptUrl(path)                     // createSignedUrl(path, 300) -> url string or null
```

## 7. React contracts

### `src/context/DataContext.jsx`
```js
export function DataProvider({ session, children })
export function useData() // ->
// {
//   profile,                    // profiles row for session user
//   projects,                   // visible projects (active only), name asc
//   plotsByProject,             // { [projectId]: plots[] name asc }
//   workerTypes, contractors, jcbOperators, workTypes,   // arrays, name asc
//   refs,                       // { workerTypesById, contractorsById, jcbOperatorsById, workTypesById }
//   mtdByProject,               // { [projectId]: number } from project_month_totals
//   loading, error,
//   reload,                     // async () => refetch everything
//   clearMustChangePassword,    // async () => update own profile flag + local state
// }
```

### `src/App.jsx` — route state machine
```js
// route = { name, params }
// names: 'login' | 'forgot-password' | 'set-password' | 'project-picker' |
//        'home' | 'add-expense' | 'entry-detail' | 'history' | 'admin'
// navigate(name, params = {}) — plain setState; no URL sync (per spec §10)
```
Branch logic (in order): no session → `login`; `profile.must_change_password` → `set-password`; `profile.role === 'admin'` → `admin`; supervisor with exactly 1 project → `home` (that project auto-selected); else → `project-picker`. `App` owns `activeProjectId`; screens receive `{ navigate, ...route.params }` plus what they need. `entry-detail` params: `{ entryId }`.

### Components
```js
<Sheet open onClose title>{children}</Sheet>        // bottom sheet, backdrop click = onClose
<AmountInput label value onChange min placeholder /> // value: string; sanitizes to digits + one '.', inputMode="decimal"
<PhotoCapture value onChange />                      // value: Blob|null; <input type="file" accept="image/*" capture="environment">, calls compressImage
```

### Category forms (`src/forms/*.jsx`) — uniform interface
```js
<NmrForm details onChange refs />   // details = current details object; onChange(nextDetails)
// same props for JcbForm, ContractorFeeForm, LaborContractForm, GeneralForm
// forms render fields only; AddExpense owns plot/payment/date/notes/photo/total/submit
```

### Entry insert (AddExpense) — exact payload
```js
const payload = {
  id: crypto.randomUUID(),
  project_id, plot_id, category, details,
  total_amount: computeTotal(category, details),
  payment_mode, notes, entry_date,          // 'YYYY-MM-DD', istToday() or istDaysAgo(1)
};
const { error } = await supabase.from('entries').insert(payload);
if (error && error.code !== '23505') throw error;   // 23505 (duplicate id) = retried success
```
`created_by`, `created_at`, `locked_at` are server-set — never send them.

### Admin void — exact call
```js
await supabase.from('entries')
  .update({ voided_at: new Date().toISOString(), voided_by: profile.id, void_reason: reason })
  .eq('id', entryId);
// server trigger overwrites voided_at/voided_by and enforces reason >= 10 chars; client validates too
```

### Realtime subscription (Home, History)
```js
const channel = supabase
  .channel(`entries-${projectId}`)
  .on('postgres_changes',
    { event: '*', schema: 'public', table: 'entries', filter: `project_id=eq.${projectId}` },
    (payload) => setEntries((prev) => applyRealtimeEvent(prev, payload)))
  .subscribe();
// cleanup: supabase.removeChannel(channel)
```
DELETE payloads carry only `old.id` — `applyRealtimeEvent` removes by id. Postgres Changes respects RLS for authenticated tokens; the publication was set in migration 2.

## 8. Conventions

- **Tests:** Vitest, files under `tests/` mirroring `src/` (`tests/lib/categories.test.js`). jsdom environment, `globals: true`, setup file `tests/setup.js` importing `@testing-library/jest-dom/vitest`. Pure logic gets exhaustive tests; screens get render + interaction tests for logic-bearing behavior only (no snapshot tests).
- **TDD:** every task = failing test → run to see fail → minimal implementation → run to see pass → commit.
- **Commits:** conventional prefix (`feat:`, `fix:`, `test:`, `chore:`, `docs:`), imperative, ≤ 72 chars.
- **Style:** Prettier defaults; ESLint flat config with react-hooks rules; no TypeScript (plain JSX per spec).
- **Timezone:** never call `toISOString().slice(0,10)` for a display/business date — always `istToday()`/`istDaysAgo()`. `entry_date` and all grouping keys are IST date strings.
- **Money:** amounts are plain JS numbers (DB `numeric` arrives as string from PostgREST — coerce with `Number(...)` at fetch boundaries, in DataContext/screens, before math).
- **Voided entries:** excluded from every sum; still rendered (struck-through) in lists and exports.
- **Currency display:** only via `formatINR`.
