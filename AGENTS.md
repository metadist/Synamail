---
name: Synamail
description: Outlook add-in for Synaplan (Vue 3 + TypeScript + Office.js)
---

# Synamail Development Guide

Outlook add-in (taskpane + ribbon commands) that connects the user's Outlook client to a Synaplan workspace.

**Stack:** Vue 3 + TypeScript + Vite + Office.js. **No** PHP, **no** Docker for runtime — the add-in is a static SPA hosted on HTTPS and sideloaded into Outlook via a manifest.

**Toolchain:** Node ≥ 22 (Maintenance LTS through April 2027). Active LTS is Node 24 (`.nvmrc`); CI exercises both. `engines.node` in `package.json` is the source of truth.

**Authoritative docs:** [`planning/PLAN.md`](planning/PLAN.md), [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md), [`docs/FEATURES.md`](docs/FEATURES.md), [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md), [`planning/STEPS.md`](planning/STEPS.md), [`planning/GUI_DEFINITIONS.md`](planning/GUI_DEFINITIONS.md), [`docs/GLOSSARY.md`](docs/GLOSSARY.md).

## Critical Rules

### Merge Conflicts — NEVER Accept One Side Blindly

When resolving merge conflicts:

1. **Manually merge both sides** — understand what each adds/changes.
2. **NEVER** use `git checkout --ours` or `git checkout --theirs` for code files.
3. **Preserve ALL functionality** from both branches unless explicitly instructed.
4. **If unsure, ASK** — throwing away code is worse than asking.

### No Attribution in Commits

- Use [Conventional Commits](https://www.conventionalcommits.org/) (`feat:`, `fix:`, `chore:`, `docs:`, `refactor:`, `test:`, `build:`, `ci:`, `perf:`).
- **NEVER** add "Generated with Claude Code", "Co-Authored-By: Claude", AI tool footers, or similar.

### MANDATORY Pre-Commit Gate — Run Tests BEFORE Every Commit

**You MUST run and pass ALL of these before committing or allowing a commit.** This matches what CI runs on GitHub. If any step fails, fix the issue before committing. No exceptions.

```bash
# Step 1: Lint (Prettier + ESLint) — Vue + TS
make lint

# Step 2: Static type check — vue-tsc -b (catches errors ESLint misses!)
make check-types

# Step 3: Unit + component tests (Vitest)
make test

# Step 4: Manifest validation (office-addin-manifest)
make validate

# Step 5: Production build (Vite, fails on bundle-size budget breach)
make build
```

**Or run everything in one shot:**

```bash
make ci-local
```

`make ci-local` is the **local CI** — it runs the exact same steps as the
GitHub `build` job (lint → check-types → test → validate → build → bundle
budget). If it's green locally, CI is green. If it's red locally, CI will be
red — do not push.

**Rules:**

- Run the FULL test suite (`make test`), not just a subset.
- If `make check-types` fails, fix the type errors before committing — no `// @ts-ignore` escape hatches without a tracked issue + comment.
- E2E (Playwright + sideload) is **not** required on every commit — it runs nightly + before sprint review. But if you touched `useAuth.ts`, `synaplan-client.ts`, or any view, run `make test-e2e` locally before the PR.
- After changing Synaplan's OpenAPI surface or pulling new schemas, run `make generate-schemas` then re-run `make check-types`.
- **NEVER** commit with failing tests — this blocks the entire CI/CD pipeline and the AppSource release.

### Hook enforcement — the local gate is not optional

The local gate is wired so it **cannot be silently skipped**:

- `npm install` runs a `prepare` script that sets `git config core.hooksPath .githooks`. Hooks are enabled automatically on every fresh clone — you never have to remember `make bootstrap`.
- **`pre-commit`** runs `make ci-local` on any non-docs commit.
- **`pre-push`** runs `make ci-local` again (the last line of defence before GitHub). A commit made by a tool that bypassed `pre-commit` — IDE auto-commit, `git commit --no-verify` — is still caught here.
- **NEVER** use `--no-verify` to bypass the gate on a branch that will be pushed to `main`. If the gate is wrong, fix the gate, not the bypass.

If you change anything here, run `git config core.hooksPath` and confirm it prints `.githooks`.

### Tests MUST be deterministic and environment-independent

The single most common cause of "green locally, red in CI" is a test that
depends on ambient state. CI runs with a **clean checkout and no `.env*`
files**, Node 22 **and** 24, in `mode: test` (where `import.meta.env.DEV` is
`true`).

- **NEVER** let a test's outcome depend on `.env.local` / `.env*` or any `import.meta.env.VITE_*` value. If a test needs a specific env, set it explicitly with `vi.stubEnv(...)` and restore it with `vi.unstubAllEnvs()`. (This is exactly what bit `buildDialogUrl` — the test asserted the real-flow URL while the code defaults to the mock relay in dev/test.)
- Don't rely on machine locale, timezone, network, or wall-clock time. Inject or stub them.
- A test that only passes because you have `.env.local` set is a broken test — it will fail in CI.

### Cross-repo PR coordination

Sprint 2 includes one **cross-repo change**: the `/addin/connect` Vue route added inside `@/wwwroot/synaplan/frontend/` (router entry + view). It is **not** a separate "synaplan-website" repo — that doesn't exist. The bridge page builds into the same Docker image as the rest of the app; production picks it up via `synaplan-platform` pulling the new `ghcr.io/metadist/synaplan:latest`. See `docs/SYNAPLAN_INTEGRATION.md`.

When you open that PR (in `@/wwwroot/synaplan/`):

```bash
# Synaplan's house pre-commit gate — non-negotiable:
make lint && make -C backend phpstan && make test \
  && docker compose exec -T frontend npm run check:types
```

Run it from the `synaplan` repo, not from this one.

## Essential Commands

```bash
# Discover
make help

# Development
make dev               # Vite dev server on https://localhost:3000
make sideload          # Sideload manifest.xml into Outlook on the Web

# Quality (ALWAYS before committing)
make ci-local          # lint + check-types + test + validate + build
make lint              # ESLint + Prettier
make check-types       # vue-tsc -b
make test              # Vitest
make test-e2e          # Playwright (requires sideload-ready manifest)
make validate          # office-addin-manifest validate manifest.xml

# Build
make build             # Vite production build
make build-manifest    # Generate manifest.unified.json from manifest.xml

# Schemas
make generate-schemas  # Regenerate Zod schemas from Synaplan OpenAPI spec
```

## Key Constraints

### Frontend Runtime Config

- **NO** `VITE_*` env vars for runtime config — the manifest is shipped to users, and the Synaplan base URL is configurable at runtime via Settings.
- The Synaplan base URL is read from `Office.context.roamingSettings.baseUrl`, with a fallback to a build-time default (`https://web.synaplan.com`).
- **NEVER** hardcode `http://localhost:8000` or any specific Synaplan host outside the build-time default.

### Manifest

- `manifest.xml` is the source of truth in v1; `manifest.unified.json` is generated from it (Sprint 4).
- Every URL in the manifest **must** be HTTPS.
- `AppDomains` must list every external host the taskpane loads (Synaplan API host, Synaplan login origin).
- Run `make validate` after any manifest change.

### Office.js

- Always wrap entry points in `Office.onReady`.
- Use **requirement-set feature-detection** — minimum Mailbox 1.8, prefer 1.10–1.15 features behind `Office.context.requirements.isSetSupported(...)`.
- The add-in must work on Outlook on the Web, new Outlook for Windows, classic Outlook 2024, and Outlook on Mac.

### Auth

- API key lives in `Office.context.roamingSettings` (encrypted at rest by Exchange, per-mailbox).
- Never log the key.
- On 401 → clear the key → bounce to `SignIn.vue`.

### i18n

- All UI text through `vue-i18n`.
- **Always update BOTH** `src/locales/en.json` AND `src/locales/de.json`.

## Code Style Quick Reference

| Area       | Standard                                                                  |
| ---------- | ------------------------------------------------------------------------- |
| TypeScript | Strict mode, no `any`, single quotes, **no semicolons**.                  |
| Vue        | Composition API, `<script setup lang="ts">`, TypeScript.                  |
| CSS        | Fluent UI design tokens via CSS variables (no Tailwind, no Fluent React). |
| Imports    | Sorted by `eslint-plugin-import` + organize-imports.                      |
| HTTP       | `fetch` + Zod-validated schemas, generated from Synaplan OpenAPI.         |

## Architecture Patterns

- **Views** (`src/taskpane/views/`): top-level routed surfaces.
- **Composables** (`src/taskpane/composables/`): reusable logic (auth, roaming, Outlook item, Synaplan client).
- **Components** (`src/taskpane/components/`): pure UI, under 200 lines each.
- **Shared** (`src/shared/`): cross-cutting helpers (client, schemas, prompts, types).

## Boundaries

### Ask First Before

- Adding dependencies (`npm`).
- Modifying `manifest.xml` requirement sets or `Permissions`.
- Touching `.github/workflows/ci.yml` (gate-job semantics).
- Changing the auth flow or roaming-settings schema.
- Adding a new view that talks to a Synaplan endpoint not yet listed in `docs/FEATURES.md` §6.

### Never Do

- Commit secrets, `.env` files with credentials, or signed manifests with production GUIDs to a non-prod branch.
- Edit `node_modules/`.
- Commit `dist/` directories or generated `manifest.unified.json` artefacts.
- Push directly to `main`.
- Force push to `main` or any branch with an open PR.
- Skip the pre-commit gate.

## Detailed Documentation

- [`planning/PLAN.md`](planning/PLAN.md) — the 4-sprint plan.
- [`planning/STATUS.md`](planning/STATUS.md) — current build / sprint status.
- [`planning/STEPS.md`](planning/STEPS.md) — step-by-step execution.
- [`planning/GUI_DEFINITIONS.md`](planning/GUI_DEFINITIONS.md) — UI spec + assets.
- [`planning/APPSOURCE_CHECKLIST.md`](planning/APPSOURCE_CHECKLIST.md) — Microsoft submission gate.
- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — auth, API surface, security.
- [`docs/FEATURES.md`](docs/FEATURES.md) — feature contract.
- [`docs/GLOSSARY.md`](docs/GLOSSARY.md) — canonical terminology.
- [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) — commits, branches, reviews, releases.
- [`docs/CONTRIBUTING.md`](docs/CONTRIBUTING.md) — contributor flow.
- [`docs/SYNAPLAN_INTEGRATION.md`](docs/SYNAPLAN_INTEGRATION.md) — cross-repo integration map.
