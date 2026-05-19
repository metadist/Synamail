# Synamail — Execution Steps

Step-by-step plan aligned to the 4 sprints in [`PLAN.md`](PLAN.md). Every step has a **Build** section, an **Acceptance test**, and a **Done when…** criterion. Tick boxes as work lands.

> Companion docs: [`../docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md), [`../docs/FEATURES.md`](../docs/FEATURES.md), [`GUI_DEFINITIONS.md`](GUI_DEFINITIONS.md), [`STATUS.md`](STATUS.md).

## Sprint 1 — Planning, Design & Test Definition

### Step 1.1 — Lock the architecture & terminology

**Build**

- Confirm `docs/ARCHITECTURE.md` matches the current Synaplan API surface (no inventions).
- Confirm terminology mapping in `PLAN.md` ("RAG keys" → groups, "RULE" → Synapse Routing rules) is reflected in every other doc.

**Acceptance test**

- Grep `STEPS.md`, `PLAN.md`, `FEATURES.md`, `GUI_DEFINITIONS.md` for "RAG key" — every hit is either inside a glossary entry or annotated with "= RAG group".
- Every endpoint mentioned in `FEATURES.md` exists in `synaplan/backend/src/Controller/` or is flagged "confirm during wire-up".

**Done when** the doc-grep returns no untranslated terms and the endpoint check passes.

### Step 1.2 — Lock the feature specs

**Build**

- `docs/FEATURES.md` documents every feature from the user's brief: AI summarise/translate/draft, RAG ingest, contact knowledge base, RULE integration.
- Each feature lists input, output, endpoint, UI view, edge cases.

**Acceptance test**

- For every row in `FEATURES.md` §6 (coverage matrix), there is a corresponding step in Sprint 2 or Sprint 3 below.

**Done when** §6 of `FEATURES.md` and the steps below are in 1:1 correspondence.

### Step 1.3 — Lock the GUI spec

**Build**

- `planning/GUI_DEFINITIONS.md` lists every view, every reusable component, all icon sizes, all AppSource asset sizes.
- Wireframes are described in text (ASCII or bullet structure) so they can be sanity-checked without a design tool.

**Acceptance test**

- For each view named in `PLAN.md` deliverable D2.4 / D3.6 / D3.7 / D4.4, the GUI spec has a section.

**Done when** all six views (SignIn, ReadMode, ComposeMode, Settings, RuleEditor, ContactKnowledgeBase) are documented and the asset list is complete.

### Step 1.4 — Lock the test plan

**Build**

- The five test layers (static / unit / component / E2E / manifest) are listed with tooling and trigger in `PLAN.md`.
- This file lists the specific test cases each Sprint-2 and Sprint-3 step must add.

**Done when** the layered test plan is in `PLAN.md` and per-step tests are written below.

### Step 1.5 — Identify cross-repo work

**Build**

- `PLAN.md` "Cross-repo dependencies" table names every change to `synaplan` and `synaplan-platform` and assigns it to a sprint. See `docs/SYNAPLAN_INTEGRATION.md` for the full integration map.

**Done when** the only cross-repo Synaplan-side change in v1 is the `/addin/connect` Vue route in `synaplan/frontend/` (Sprint 2), and the deploy loop on `synaplan-platform` is documented (`docker compose pull && re-startweb{1,2,3}.sh`).

### Step 1.6 — Engineering scaffold (commit process + CI + hooks)

Lands the professional commit process _before_ any code, so Sprint 2 starts on a paved road. Mirrors the gate structure used in `@/wwwroot/synaplan/` (lint → types → tests → build → e2e → gate job), adapted for this repo's TypeScript-only stack.

**Build**

- `AGENTS.md` — workspace rules for AI agents (synaplan-style, with this repo's pre-commit gate).
- `docs/CONTRIBUTING.md` — contributor entry point.
- `docs/COMMIT_PROCESS.md` — Conventional Commits spec, branch policy, PR rules, release flow.
- `.github/`:
  - `workflows/ci.yml` — staged CI (`detect` → `docs` + `commitlint` + `manifest` + `build` → `e2e` → `all-checks-passed`). Stages auto-activate as the repo grows.
  - `PULL_REQUEST_TEMPLATE.md` — type + sprint/step + pre-commit-gate ticks + cross-repo link.
  - `ISSUE_TEMPLATE/bug.md` and `feature.md`.
  - `dependabot.yml` — weekly grouped npm + actions updates.
- `Makefile` — `help`, `bootstrap`, `doctor`, `dev`, `sideload`, `lint`, `format`, `check-types`, `test`, `test-e2e`, `validate`, `build`, `build-manifest`, `generate-schemas`, `ci-local`, `clean`. Every npm-dependent target is a no-op while `package.json` is absent, prints a "skip: Sprint 2.1" hint.
- `.githooks/`:
  - `pre-commit` — runs `make ci-local` for code commits, fast docs lint for docs-only commits.
  - `commit-msg` — regex Conventional-Commits validation; auto-delegates to `commitlint` once installed.
  - `pre-push` — warns when auth/client/view/manifest files are in the push range (run `make test-e2e`).
- `commitlint.config.cjs` — used by the hook + CI once npm lands.
- `.editorconfig`, `.gitattributes`, `.gitignore`, `.markdownlint.jsonc`.

**Acceptance test**

- `make doctor` reports the planning-stage state correctly (no `package.json` yet, hooks enabled).
- `git config core.hooksPath` is `.githooks` after `make bootstrap`.
- A commit with a malformed subject (e.g. `Updated stuff`) is rejected by the `commit-msg` hook.
- A docs-only commit passes the pre-commit hook in under 2 s.
- The `docs` and `all-checks-passed` jobs go green on a planning-stage PR (other CI stages are correctly skipped because `package.json` / `manifest.xml` / `src/` are absent).

**Done when** the scaffold is merged to `main`, branch protection allows that and only that PR through, and every subsequent contributor (human or AI) inherits the pre-commit gate by running `make bootstrap`.

---

## Sprint 2 — GUI, Sideload, Live Identification

> **Status as of 2026-05-16:** Steps 2.1, 2.2, 2.4, 2.5, 2.6 (mock relay), 2.7 are code-complete. `make ci-local` exits 0 (lint + types + 30 tests + manifest + build). See [`STATUS.md`](STATUS.md) for full handover notes and the next actions for the human.

### Step 2.1 — Bootstrap project (1 day)

This step lands the actual code scaffold on top of the engineering process that's already in place (Sprint 1 shipped `Makefile`, CI workflow, git hooks, AGENTS.md, COMMIT_PROCESS.md, PR/issue templates). The `make` targets that were no-ops during planning go live here.

**Build**

- `cd /wwwroot/Synamail && npx --package yo --package generator-office -- yo office` → Outlook Add-in, TypeScript, name **Synamail**.
- Replace webpack with **Vite + Vue 3 + TS**; keep Yeoman's `manifest.xml`, `commands/`, dev-cert scripts.
- `package.json` scripts (mirrored from `Makefile`): `dev`, `build`, `lint`, `format`, `check:types`, `test`, `test:e2e`, `validate`, `sideload`, `generate:schemas`.
- Pin: `@types/office-js`, `office-addin-validator`, `office-addin-debugging`, `office-addin-dev-certs`, `vite`, `vue`, `vue-tsc`, `typescript`, `eslint`, `prettier`, `vitest`, `@vue/test-utils`, `@playwright/test`, `vue-i18n`, `zod`, `@commitlint/cli`, `@commitlint/config-conventional`.
- Add `.prettierrc` + `eslint.config.js` aligned with `synaplan/frontend/` conventions (no semicolons, single quotes, no `any`).
- Hook up `commitlint` to the existing `.githooks/commit-msg` (the hook auto-detects `node_modules/.bin/commitlint` and switches from the regex fallback to commitlint once installed).
- Run `make bootstrap` to verify hooks are enabled.

**Acceptance test**

```bash
make doctor         # node + npm + git hooks all OK
make bootstrap      # enables hooks, installs deps
make ci-local       # lint + check-types + test + validate + build all green
```

In CI: the existing `build` job in `.github/workflows/ci.yml` now activates (it was skipped while `package.json` was absent — see `detect` job) and must pass.

**Done when** `make ci-local` exits 0 locally **and** the `build` + `commitlint` CI jobs are green on the PR.

### Step 2.2 — Hello-Synamail taskpane sideloaded into OWA (½ day)

**Build**

- `manifest.xml`: `DisplayName="Synamail"`, `Mailbox 1.8`, `DialogApi 1.1`, `https://localhost:3000/taskpane.html`, ribbon entries for read + compose, `Permissions=ReadWriteItem`, `AppDomains` placeholder for `web.synaplan.com`.
- `src/taskpane/App.vue` shows the title + current subject (read) or "Compose mode" (compose). Uses `Office.onReady`.
- Dev cert via `office-addin-dev-certs install`.
- `npm run sideload` wraps `office-addin-debugging start manifest.xml`.

**Acceptance test (manual)**

1. `npm run dev` serves on `https://localhost:3000`.
2. `npm run sideload` opens Outlook on the Web with the Synamail ribbon.
3. Click on a selected email → subject visible in the taskpane.
4. Open a new message → ribbon button → "Compose mode" placeholder.

**Acceptance test (automated)**

- Vitest unit test on `App.vue` with mocked `Office.context.mailbox.item`.
- CI: lint + types + manifest validate + build.

**Done when** manual + automated tests pass on OWA Edge + OWA Chrome.

### Step 2.3 — CI activation + branch protection (½ day)

The CI workflow already exists from Sprint 1 (`.github/workflows/ci.yml`) with stages that auto-activate as the repo grows (`detect` job picks up `package.json`, `manifest.xml`, `src/`). This step lights up the remaining stages and locks down `main`.

**Build**

- Confirm all CI stages green on a feature PR: `detect`, `docs`, `commitlint` (on PR), `manifest` (after Step 2.2), `build`, `e2e` (after Step 2.6).
- Configure branch protection on `main`:
  - Require status check: `All Checks Passed` (the gate job).
  - Require 1 approving review from CODEOWNERS.
  - Require linear history (squash-merge only).
  - Disallow force-push.
  - Disallow direct push (PR required).
- Add `CODEOWNERS` file mapping `/docs`, `/planning`, `/.github`, `/src/taskpane/views`, `/src/shared/synaplan-client.ts`, `manifest.xml` to the right reviewers.
- Verify dependabot is open on weekly schedule (config already shipped in Sprint 1).

**Acceptance test**

- Open a PR with an intentional lint error → `build` stage fails → `All Checks Passed` fails → PR cannot merge.
- Fix the lint error → `All Checks Passed` green → squash-merge available.
- Open a PR with a malformed commit subject (e.g. `WIP some change`) → `commitlint` stage fails on PR.
- Push attempt to `main` (as a non-admin) is rejected by branch protection.

**Done when** PRs cannot merge with red CI, branch protection rejects direct pushes to `main`, and the `commitlint` job blocks non-Conventional-Commits messages.

### Step 2.4 — Read-mode UI shell + mock client (1 day)

**Build**

- `src/shared/synaplan-client.ts` — fetch wrapper _and_ a `MockSynaplanClient` selected by env var.
- `src/taskpane/views/ReadMode.vue` — five buttons (Summarise, Translate, Draft reply, Classify, Save to RAG, Ask) + contact pill stub.
- `src/taskpane/composables/useOutlookItem.ts` — exposes `subject`, `from`, `to`, `bodyText` (lazy), `attachments`.
- Mock client returns canned strings with 500 ms delay.

**Acceptance test**

- Vitest covers `useOutlookItem`.
- Manual: every read-mode button round-trips through the mock; spinner + error states render.

**Done when** all read-mode buttons demo correctly against the mock.

### Step 2.5 — Compose-mode UI shell + mock client (1 day)

**Build**

- `src/taskpane/views/ComposeMode.vue` — Draft from prompt, Improve selection, Translate selection, Insert from RAG (stub).
- Draft → mock client → `body.setAsync(html, { coercionType: 'html' })`.
- Improve/Translate → `body.getSelectedDataAsync` → mock → `body.setSelectedDataAsync`.

**Acceptance test**

- Vitest on the compose composable.
- Manual: Draft from prompt fills body; Translate replaces selection.

**Done when** all compose buttons demo correctly against the mock.

### Step 2.6 — SignIn view + Office Dialog flow (½ day, mock relay)

**Build**

- `SignIn.vue` — "Sign in to Synaplan" button + "Use a self-hosted instance" link.
- `useAuth.ts` — opens `Office.context.ui.displayDialogAsync(authRelayUrl + '?state=<nonce>')`, handles `DialogMessageReceived` + `DialogEventReceived`, validates the state nonce.
- `dialog/auth-relay.html` (mock): immediately calls `Office.context.ui.messageParent(JSON.stringify({ apiKey: 'fake-test-key', keyId: 1, email: 'demo@synaplan.test', baseUrl: 'https://web.synaplan.com', state }))`.
- `useRoamingSettings.ts` — persists `{ apiKey, keyId, email, baseUrl }`.
- App routing: no key → SignIn; has key + read item → ReadMode; has key + compose item → ComposeMode.

**Acceptance test**

- Manual: SignIn shown on fresh load → click → mock dialog closes → ReadMode/ComposeMode renders → reload Outlook → state survives.
- Vitest: mocked `displayDialogAsync`, state-nonce validation, storage round-trip.

**Done when** the local mock loop works including nonce rejection.

### Step 2.7 — Settings / configuration window + Sign-out (½ day)

**Build**

- `Settings.vue` — shows email + base URL, "Sign out" button, "Use a self-hosted instance" override.
- Sign-out clears roaming settings; the live `apikeys` revoke call is wired in Sprint 3 (Step 3.x).
- Base-URL override applies **before** the next sign-in dialog opens.

**Acceptance test**

- Manual: sign in → see Settings → sign out → SignIn returns.
- Manual: override base URL to `https://example.invalid` → next sign-in dialog opens that URL.

**Done when** sign-out + base-URL override work.

### Step 2.8 — Cross-repo: `/addin/connect` Vue route in `synaplan/frontend/` (1 day)

This is the **only Synaplan-side change** in v1. Done in `@/wwwroot/synaplan/frontend/`. **NOT** in a "synaplan-website" repo — that doesn't exist. `web.synaplan.com` is served by the same Docker image as the API. See `docs/SYNAPLAN_INTEGRATION.md`.

**Build (in `@/wwwroot/synaplan/frontend/`)**

- Router: add a public route in `src/router/index.ts`:

  ```ts
  {
    path: '/addin/connect',
    name: 'addin-connect',
    component: () => import('@/views/AddinConnectView.vue'),
    meta: { requiresAuth: false, public: true, titleKey: 'pageTitles.addinConnect' },
  }
  ```

- New view `src/views/AddinConnectView.vue`:
  - Load Office.js (`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`) inside this page only (the SPA doesn't normally load it).
  - Read `state` and `baseUrl` from query string.
  - If not authenticated (auth store) → redirect to `/login?redirect=/addin/connect&state=...`.
  - If authenticated → render "Connect this Outlook to `<email>`?" with **Connect** button.
  - On Connect: call the existing `apiKeysApi.create(...)` from `src/services/api/apiKeysApi.ts` with `{ name: 'Outlook Add-in (<UA-derived host>)', scopes: ['messages:*','chats:*','files:*','rag:*'] }`.
  - On 201: `Office.context.ui.messageParent(JSON.stringify({ apiKey, keyId, email, baseUrl, state }))`.
  - On error: render error + Try again.
- i18n: add `pageTitles.addinConnect` to `en.json` and `de.json`.
- CORS: nothing to do. `synaplan/backend/config/packages/nelmio_cors.yaml` is already `allow_origin: ['*']`.

**Deploy (in `@/wwwroot/synaplan-platform/`, after the Synaplan PR merges and CI republishes the image)**

Per web node (web1 / web2 / web3):

```bash
cd /wwwroot/synaplan-platform
docker compose pull backend worker
bash re-startweb1.sh    # or re-startweb2.sh / re-startweb3.sh per node
docker compose ps
curl -sf http://localhost/api/health
```

No `synaplan-platform` files change. The new image carries the bridge page.

**Acceptance test (live, against `web.synaplan.com`)**

- `curl -sfL -o /dev/null -w '%{http_code}\n' https://web.synaplan.com/addin/connect` → `200` from each web node.
- Sideloaded Synamail → click Sign in → live `https://web.synaplan.com/addin/connect` opens in the Office dialog → completes OAuth with the user's existing provider → taskpane receives a real `apiKey`.
- Synaplan API-keys page lists the new entry as `Outlook Add-in (<host>)`.
- `synaplan` PR passes the house pre-commit gate:

  ```bash
  make lint && make -C backend phpstan && make test \
    && docker compose exec -T frontend npm run check:types
  ```

**Done when** a fresh sideload of Synamail produces a real API key in `roamingSettings` with no copy-paste, via the live login on all three web nodes.

### Step 2.9 — Wire SignIn to the live bridge (½ day, in Synamail)

**Build**

- Replace the mock `dialog/auth-relay.html` URL with the live `https://web.synaplan.com/addin/connect` URL (configurable per environment).
- Keep the mock available behind an env flag for offline development.

**Acceptance test**

- E2E (Playwright + OWA): full sign-in dance against the live bridge with a dedicated test user. State-nonce rejection case included.

**Done when** Sprint 2's exit criteria from `PLAN.md` are all met.

🟢 **Sprint 2 checkpoint** — the user can sideload Synamail into their own Outlook, sign in via `web.synaplan.com`, and the configuration window works. Read/compose actions still go through the mock client.

---

## Sprint 3 — Functionality & Live API Calls

### Step 3.1 — Replace mock client with real Synaplan client (½ day)

**Build**

- `synaplan-client.ts` — real fetch, `X-API-Key`, exponential-backoff retry for 5xx, 401 → clear key → bounce to SignIn.
- Generate Zod schemas from `https://web.synaplan.com/api/doc.json` and check them in.
- Smoke test on boot: `GET /api/v1/profile` (or equivalent); if 401, clear key.

**Acceptance test**

- Manual: sign in → Network tab shows `GET /profile` returning 200.
- Manual: revoke the key from Synaplan's API-keys page → next action returns 401 → add-in clears state and shows SignIn.
- Unit: header injection, retry on 502, 401 mapping, Zod parse failure surface.

**Done when** every taskpane HTTP call goes through the real client.

### Step 3.2 — Sign-out revokes the key (½ day)

**Build**

- Wire `Settings.vue` "Sign out" to `DELETE /api/v1/apikeys/{keyId}`.
- Clear roaming settings only after a 2xx (or after a 404 — already gone).

**Acceptance test**

- Manual + automated: sign out → key disappears from Synaplan's API-keys page.

**Done when** the key revocation round-trip is verified.

### Step 3.3 — Read-mode AI features wired to live endpoints (2 days)

Implement in this order, each as its own PR.

1. **Summarise** → `POST /api/v1/messages/send` with summarise prompt.
2. **Translate** → `POST /api/v1/messages/send` with translation prompt + target-language picker.
3. **Draft reply** → `messages/send` with reply prompt → `mailbox.displayReplyForm({ htmlBody })`.
4. **Classify** → `messages/send` with classification prompt + JSON output schema.
5. **Ask follow-ups** → `POST /api/v1/chats` per `conversationId` + `POST /api/v1/chats/{id}/messages` for subsequent turns; `GET .../messages` for display.

**Acceptance test (manual + Playwright E2E)**

- Each action produces a real Synaplan response on a seed email.
- Three Ask turns share the same `chatId`; verified in Synaplan chat history.
- Token-overflow path falls back to chunked summarise.
- 401 / 5xx / network-down each surface a human-readable error.

**Done when** every read-mode action passes E2E against the test user.

### Step 3.4 — Save-to-RAG with group picker (1 day)

**Build**

- `Office.context.mailbox.item.getAsFileAsync` → `.eml` upload via `POST /api/v1/files/upload`.
- For each attachment: `getAttachmentsAsync` + `getAttachmentContentAsync` → upload.
- `POST /api/v1/files/{id}/process` with chosen processing level + group id.
- Group picker driven by `GET /api/v1/files/groups`; "Create new group" inline calls `POST /api/v1/files/groups`.
- Pre-select `lastRagGroupId` from roaming; offer `contact:<sender>` as a suggestion chip.

**Acceptance test**

- Manual: upload an email + 2 attachments to an existing group → see them listed in Synaplan.
- E2E: create a new group via the picker → save email → group exists with the email in it.

**Done when** uploads + group create/select work end-to-end.

### Step 3.5 — Contact knowledge base (1.5 days)

**Build**

- `ContactKnowledgeBase.vue` — sender/recipient picker pill, search input, results list, "Save to this contact" action, "Ask about this contact" action.
- "Save to <contact>" creates the `contact:<email>` group on first use, then runs the same upload + process flow.
- Search calls `POST /api/v1/rag/search` with `groups: ["contact:<email>"]`.
- "Ask about <contact>" opens a chat with the contact group as RAG scope (exact endpoint shape confirmed in Step 3.1).

**Acceptance test**

- E2E:
  1. Save an email from `alice@example.com` to her contact group.
  2. Open a different email from her → ContactKnowledgeBase pill shows; click → see the earlier email.
  3. "Ask about alice" → chat answer cites the earlier email.

**Done when** the three-step E2E above passes.

### Step 3.6 — Compose-mode AI features wired (1 day)

**Build**

- Draft from prompt, Improve / Shorten / Translate selection, Insert from RAG — all calling `messages/send` or `rag/search` and writing back via `setAsync` / `setSelectedDataAsync`.

**Acceptance test**

- E2E: open new mail → Draft from prompt fills body; select text → Translate replaces selection; Insert from RAG finds and inserts a real snippet.

**Done when** all three compose flows pass E2E.

### Step 3.7 — RULE integration (Synapse Routing rules) (1–2 days)

**Pre-step audit (must run before any RuleEditor code)**

- Read `synaplan/backend/src/Controller/PromptController.php` and document in this step:
  - Does `GET /api/v1/prompts` return `BSELECTION_RULES`?
  - Is there a user-scoped write method (`PATCH`/`PUT`/`POST`) on a prompt?
  - What scopes / roles does that endpoint require?
- Record the answers in `docs/FEATURES.md` §5.4 before writing code.

**Build**

- `RuleEditor.vue` — list topics + their `BSELECTION_RULES`, add/remove keyword matchers (write affordances depend on the audit above).
- "Test against current email" — calls `POST /api/v1/admin/synapse/dry-run`. **Hidden for non-admin users**; admins see it; non-admins see an inline hint.
- "Apply this email's pattern as a rule" — pre-fills RuleEditor with a candidate rule from the current email (subject substring, sender domain).
- If the audit shows **no user-scoped write endpoint**, ship `RuleEditor.vue` as **read-only** in v1, file a Synaplan-side ticket for the missing endpoint, and document the deferral in `FEATURES.md` §5.4.

**Acceptance test**

- Audit results recorded in `FEATURES.md` §5.4 before any code merges.
- E2E (write path, if available): add a keyword rule via the editor → forward a matching email via the user's `smart+keyword@…` flow → routed topic visible in Synaplan logs.
- E2E (read-only fallback): rules render matches `GET /api/v1/prompts` payload.
- E2E (admin path): admin user sees the dry-run button and gets a routing preview against a sample email.

**Done when** Sprint 3 exit criteria from `PLAN.md` are met. If the user-scoped write endpoint is missing, RULE editing is a tracked v1.1 deferral and the read-only view ships in v1.

### Step 3.8 — i18n (en + de) (½ day)

**Build**

- `vue-i18n` + `en.json` + `de.json` covering every UI string.
- Locale auto-detected from `Office.context.displayLanguage`; manual override in Settings.

**Acceptance test**

- Switch Outlook to German → reload → all visible strings in German.

**Done when** both locale files cover 100% of UI strings (lint rule fails on missing keys).

🟢 **Sprint 3 checkpoint** — every feature in the user's brief works live against `web.synaplan.com`.

---

## Sprint 4 — Release: AppSource

### Step 4.1 — Production hosting (½ day)

**Build**

- Deploy taskpane SPA to `https://addin.synaplan.com` (or `https://web.synaplan.com/addin/`).
- Strict CSP, HTTPS-only, no third-party scripts.
- Synaplan API CORS allow-list updated.
- `manifest.xml` `SourceLocation` + `AppDomains` updated to prod.

**Acceptance test**

- Smoke test: install the production-pointed manifest → full sign-in + summarise round-trip works.

**Done when** the production-pointed manifest passes a full E2E run.

### Step 4.2 — Unified manifest variant (½ day)

**Build**

- Run `office-addin-manifest-converter` on `manifest.xml` → `manifest.unified.json`.
- CI validates both manifests.
- Sideload the unified manifest into new Outlook for Windows.

**Acceptance test**

- `office-addin-validator manifest.xml` and `office-addin-validator manifest.unified.json` both return zero warnings.

**Done when** both manifests validate clean and the unified one sideloads.

### Step 4.3 — Graphical assets (1–2 days)

**Build**

- Icons: 16, 32, 64, 80, 128 PNG (`assets/icon-*.png`).
- Hero: 256, 512 PNG (`assets/store/hero-*.png`).
- Screenshots (1366×768): SignIn, ReadMode, ComposeMode, Settings, ContactKnowledgeBase, RuleEditor (`assets/store/screenshot-*.png`).
- Screencast video ≤ 60 s showing the killer flow (summarise → save to contact → ask) (`assets/store/screencast.mp4`).

**Acceptance test**

- Asset checklist in `planning/APPSOURCE_CHECKLIST.md` is fully ticked.

**Done when** every required asset exists at the correct dimension and quality.

### Step 4.4 — Store copy + legal URLs (½ day)

**Build**

- `assets/store/copy.md` — title, short description, long description, search keywords, support contact, in en + de.
- Privacy policy + Terms of Service + Support pages hosted publicly via `synaplan/frontend/` routes (or `synaplan-docs`).
- Manifest `SupportUrl` + Partner Center privacy/terms URLs all point at live pages.

**Acceptance test**

- `curl -fsSL` of each URL returns 200 from an anonymous network.

**Done when** all three URLs are publicly reachable and linked from the manifest.

### Step 4.5 — Partner Center submission (1 day + Microsoft review)

**Build**

- Create / configure Microsoft Partner Center account (Microsoft 365 and Copilot program).
- Demo Microsoft 365 tenant pre-loaded with a Synaplan account for reviewers.
- Submit via Partner Center.

**Acceptance test**

- AppSource certification policy 1100 checklist fully ticked in `planning/APPSOURCE_CHECKLIST.md`.
- Submission accepted **or** the review comments are tracked as follow-up tasks and a second submission is queued.

**Done when** Microsoft accepts the submission (typically after 1–2 review iterations, 4–8 weeks calendar).

---

## What we need from the user, mapped to sprints

| Sprint | What we need                                                                                                                                        | Why                                                                                |
| ------ | --------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| 1      | Nothing — planning only.                                                                                                                            | All Sprint-1 work is offline document work.                                        |
| 2      | A test Synaplan account on `web.synaplan.com`. Confirm bridge page lives in `synaplan/frontend/` (only option — no `synaplan-website` repo exists). | To wire up the live OAuth dialog flow end-to-end.                                  |
| 2      | Confirmation that `https://localhost:3000` and the future production origin are in the Synaplan API CORS allow-list.                                | Without this, dev sideload and production deploy will hit CORS preflight failures. |
| 3      | Same test account with at least one RAG group seeded; confirmation of Synapse Routing API access for read **and** write.                            | To exercise RAG features and RULE integration end-to-end.                          |
| 4      | Production hostname decision (`addin.synaplan.com` vs `web.synaplan.com/addin/`). Partner Center account access. Demo M365 tenant for reviewers.    | Required by Microsoft for the AppSource submission.                                |

## Continuous test pyramid (runs from Sprint 2 onwards)

| Layer        | Tooling                                                             | Trigger                              |
| ------------ | ------------------------------------------------------------------- | ------------------------------------ |
| Static       | ESLint + Prettier + `vue-tsc -b` + `office-addin-manifest validate` | every commit                         |
| Unit         | Vitest                                                              | every commit                         |
| Component    | Vitest + `@vue/test-utils` (mocked Office.context)                  | every commit                         |
| E2E (OWA)    | Playwright headless + sideload                                      | nightly + before sprint review       |
| Manifest     | `office-addin-validator`                                            | every commit + before AppSource push |
| Manual smoke | new Outlook for Windows, classic Outlook 2024, Outlook on Mac       | end of Sprint 2 and Sprint 3         |
