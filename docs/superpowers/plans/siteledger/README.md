# SiteLedger — Implementation Plan Set

Ready-to-implement plans for building and deploying SiteLedger v1.0.

**Read first:**
- Spec of record: [`../../specs/2026-07-05-siteledger-prd-v2.md`](../../specs/2026-07-05-siteledger-prd-v2.md)
- Master plan (phasing, risk, rollback, open decisions): [`../2026-07-05-siteledger-development-plan.md`](../2026-07-05-siteledger-development-plan.md)
- **[`CONTRACTS.md`](CONTRACTS.md)** — binding shared contracts (schema SQL, module signatures, component APIs, conventions). Every phase plan uses these verbatim; if a phase plan and CONTRACTS.md disagree, CONTRACTS.md wins.

## Execution order

| Plan | Scope | Depends on |
|---|---|---|
| [phase-0-scaffold.md](phase-0-scaffold.md) | Repo, Vite/ESLint/Prettier/Vitest, CI, Supabase staging+prod projects, Vercel, Sentry | — |
| [phase-1-database.md](phase-1-database.md) | Migrations 001–004, RLS policies, full SQL RLS test suite, staging push | 0 |
| [phase-2-auth-shell.md](phase-2-auth-shell.md) | Login, forgot/forced password, DataContext, route state machine | 1 |
| [phase-3-supervisor.md](phase-3-supervisor.md) | Categories engine, project picker, home + realtime, add-expense (5 forms), entry detail, history | 2 |
| [phase-4-photos.md](phase-4-photos.md) | Receipts bucket + storage RLS, capture/compress/upload, signed-URL display | 3 (‖ 5) |
| [phase-5-admin.md](phase-5-admin.md) | Admin dashboard, void flow, PDF + CSV export | 3 (‖ 4) |
| [phase-6-email.md](phase-6-email.md) | Daily 7 PM IST report: Vercel cron + Resend, manual re-run | 1 (‖ 4, 5) |
| [phase-7-deployment.md](phase-7-deployment.md) | Runbooks (users/release/onboarding), prod setup, domain + DNS, backups + restore drill, monitoring, launch | 0–6 |

## How to execute a phase

Each plan is self-contained: exact file paths, complete code, TDD steps with `- [ ]` checkboxes, exact commands with expected output, one commit per task. Execute with superpowers:subagent-driven-development (recommended) or superpowers:executing-plans. A phase is done when its exit criteria pass and CI is green.

Steps marked **[MANUAL]** need a human in a dashboard (Supabase/Vercel/Resend/DNS) and include exact click-paths.
