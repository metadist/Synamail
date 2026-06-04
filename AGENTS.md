---
name: Synamail
description: Outlook add-in for Synaplan (Vue 3 + TypeScript + Office.js)
---

# Synamail Development Guide

Outlook add-in (taskpane + ribbon commands) that connects the user's Outlook client to a Synaplan workspace.

**Stack:** Vue 3 + TypeScript + Vite + Office.js. **No** PHP, **no** Docker for runtime — the add-in is a static SPA hosted on HTTPS and sideloaded into Outlook via a manifest.

**Toolchain:** Node ≥ 22 (Maintenance LTS through April 2027). Active LTS is Node 24 (`.nvmrc`); CI exercises both. `engines.node` in `package.json` is the source of truth.

**Authoritative docs:** [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md), [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md), [`docs/FEATURES.md`](docs/FEATURES.md), [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md), [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md), [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md), [`docs/GLOSSARY.md`](docs/GLOSSARY.md).

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

### Authentication / Sign-in flow — read the doc FIRST

The sign-in / sign-out flow spans **two repos** (this add-in + `synaplan/frontend`),
is subtle, and has regressed repeatedly (most painfully: "Connect says success but
the dialog never closes"). **Before changing anything in the auth path, read and
follow [`docs/AUTH_FLOW.md`](docs/AUTH_FLOW.md)** — it is the authoritative,
field-tested reference.

Non-negotiable invariants (full list + rationale in the doc):

- **No mock mode.** Sign-in is always a real round-trip to the server URL the user
  picks. Never reintroduce a mock relay, `mock-key-` auto-selection, or a
  `MockSynaplanClient`.
- **The login round-trip MUST preserve the `redirect` (relay) param.** In
  `synaplan/frontend/src/views/AddinConnectView.vue`, the unauthenticated branch
  must use `route.fullPath` — reconstructing a subset drops `redirect` and breaks
  the desktop dialog close.
- **`buildDialogUrl` must send `redirect=relayUrl()`**, and the relay
  (`src/dialog/auth-relay.html`) must be same-origin as the taskpane **and load
  Office.js**.
- The `state` nonce must round-trip and is re-validated — a mismatch is a correct
  rejection, not a bug to remove.
- Switching local ↔ live is **Settings → Reset saved settings**, then sign in to a
  different server URL. Local bridge = `https://localhost:5174`, live =
  `https://web.synaplan.com`.

If you change any file in the flow, re-verify every step in `docs/AUTH_FLOW.md`
and update that doc in the same change.

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

### Styling & form readability (MANDATORY — must be consistent + dark-mode safe)

Every UI surface must be readable and visually consistent with the existing
views in **both light and dark mode**. The recurring bug is form fields that
render as a white box with black text on the dark pane because they never set a
background/color.

- **Use ONLY the `--syn-*` design tokens** (`tokens.css`) for color, background,
  border, spacing, radius, and font. Never hardcode hex colors, and never use a
  raw/browser-default color that isn't a token.
- **Form controls inherit the shared baseline in `app.css`** (`input` /
  `textarea` / `select` → `--syn-bg` background, `--syn-text` color, tokenized
  border, `--syn-muted` placeholder, brand focus ring, `--syn-surface` disabled).
  In a component you may tweak **layout only** (width, rows, flex). **NEVER**
  set `background`/`color` on a field to anything other than a `--syn-*` token,
  and don't leave them unset expecting a default — the global rule already
  handles it.
- **Match existing patterns**: reuse `.syn-card`, `.syn-card-title`,
  `.syn-card-sub`, `.syn-view-header`, `.syn-row`, `.syn-stack`, `.syn-muted`,
  and `ActionButton`/`Toast` rather than re-inventing per view.
- **Verify dark mode** for any new view/field (DevTools → emulate
  `prefers-color-scheme: dark`) before considering UI work done. Placeholder and
  disabled text must stay legible (use `--syn-muted`, never a faint custom grey).
- If a token is missing for a need, **add it to `tokens.css` (light + dark)** —
  don't inline a one-off color.

## Code Style Quick Reference

| Area       | Standard                                                                                                                                                          |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| TypeScript | Strict mode, no `any`, single quotes, **no semicolons**.                                                                                                          |
| Vue        | Composition API, `<script setup lang="ts">`, TypeScript.                                                                                                          |
| CSS        | Fluent UI design tokens via CSS variables (no Tailwind, no Fluent React). Form controls inherit the shared `app.css` baseline — see "Styling & form readability". |
| Imports    | Sorted by `eslint-plugin-import` + organize-imports.                                                                                                              |
| HTTP       | `fetch` + Zod-validated schemas, generated from Synaplan OpenAPI.                                                                                                 |

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

- [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md) — the 4-sprint plan.
- [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md) — current build / sprint status.
- [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md) — step-by-step execution.
- [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md) — UI spec + assets.
- [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md) — Microsoft submission gate.
- [`docs/AUTH_FLOW.md`](docs/AUTH_FLOW.md) — **authoritative sign-in / sign-out flow** (read before touching auth).
- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — auth, API surface, security.
- [`docs/FEATURES.md`](docs/FEATURES.md) — feature contract.
- [`docs/GLOSSARY.md`](docs/GLOSSARY.md) — canonical terminology.
- [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) — commits, branches, reviews, releases.
- [`docs/CONTRIBUTING.md`](docs/CONTRIBUTING.md) — contributor flow.
- [`docs/SYNAPLAN_INTEGRATION.md`](docs/SYNAPLAN_INTEGRATION.md) — cross-repo integration map.
