# Synamail — Project Plan & Sprints

Development plan for the Synamail Outlook add-in, organised into **4 sprints**.

- Sprint 1 = planning, definition, design, test-plan, GUI spec (no code shipped).
- Sprints 2–4 = the **3-sprint development phase** executing against Sprint 1's plan.

Authoritative companion documents:

- [`docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md) — technical architecture, auth flow, API surface, security.
- [`docs/FEATURES.md`](../docs/FEATURES.md) — feature specifications (AI actions, RAG, RULE integration, contact knowledge base).
- [`planning/STEPS.md`](STEPS.md) — step-by-step execution plan aligned to these sprints.
- [`planning/GUI_DEFINITIONS.md`](GUI_DEFINITIONS.md) — UI views, components, design tokens, asset list.

## Cross-repo dependencies (read this first)

| Repo                          | Why Synamail touches it                                                                                                                                                                                                                                                                                                                                                                                                                                 | Sprint     |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| `@/wwwroot/synaplan`          | **Source of all Synaplan-side changes.** Consumer of `/api/v1/messages/send`, `/chats`, `/files/upload`, `/files/{id}/process`, `/files/groups`, `/rag/search`, `/apikeys`. Sprint 2 adds the Vue route `/addin/connect` in `frontend/src/router/index.ts` + `frontend/src/views/AddinConnectView.vue`. No backend changes for v1.                                                                                                                      | 2, 3       |
| `@/wwwroot/synaplan-platform` | **Production deployment config for the 3-node cluster (web1/web2/web3).** Pulls `ghcr.io/metadist/synaplan:latest`. Bind-mounts `plugins/` for the Synaplan plugin system. Sprint 2 → 4: `docker compose pull && bash re-startweb{1,2,3}.sh` once the new image with the bridge page ships. **No structural change required.** Optional Sprint 3.5+: a `plugins/synamail/` plugin for contact-RAG-group hygiene + RULE-preview wrapping (non-blocking). | 2, 4       |
| `@/wwwroot/Synamail`          | This repo — the Outlook add-in (Vue 3 + TypeScript + Office.js).                                                                                                                                                                                                                                                                                                                                                                                        | 1, 2, 3, 4 |

> **Earlier drafts of this plan referenced a `synaplan-website` repo. That was a mistake — there is no such repo. `web.synaplan.com` is served by the same `synaplan` image as the API; the bridge page is a Vue route inside that SPA. The production cluster lives in `synaplan-platform`.**

Synaplan's house pre-commit gate (`make lint && make -C backend phpstan && make test && docker compose exec -T frontend npm run check:types`) applies to the bridge-page PR in Sprint 2 (it touches `synaplan/frontend/`).

**CORS**: `synaplan/backend/config/packages/nelmio_cors.yaml` is currently `allow_origin: ['*']` for all paths, so the Synamail taskpane on **any** origin (dev `https://localhost:3000`, prod `https://addin.synaplan.com`, self-hosted) already passes preflights. No platform-side CORS change required.

## Terminology pinning

The user's brief and Synaplan's docs use different words for the same concepts. Lock these mappings now to avoid drift in the next 3 sprints:

| User's wording           | Synaplan / code term                                                               | Where it lives                            |
| ------------------------ | ---------------------------------------------------------------------------------- | ----------------------------------------- |
| "RAG keys"               | **RAG file groups**                                                                | `GET /api/v1/files/groups`, `BFILEGROUPS` |
| "RULE integration"       | **Synapse Routing Rules** (`BSELECTION_RULES`) — Tier-0 keyword/if-then matchers   | `docs/SYNAPSE_ROUTING.md` in `synaplan/`  |
| "Identification link"    | OAuth via **Office Dialog API** → existing Synaplan login → `POST /api/v1/apikeys` | `docs/ARCHITECTURE.md` §7.5               |
| "Contact knowledge base" | Per-contact RAG group keyed by sender/recipient email (`contact:<email>`)          | New convention; see `docs/FEATURES.md` §3 |

---

## Sprint 1 — Planning, Design & Test Definition

**Goal:** Produce a complete, integration-ready plan with concrete steps, acceptance criteria, tests, GUI definitions, and graphical assets list. **No production code shipped.**

**Duration:** 3–5 working days.

**Deliverables**

- D1.1 — Clean repo structure (`docs/`, `planning/`, `src/`, `assets/`) — **done**.
- D1.2 — `PLAN.md` (this file) — sprint definitions with acceptance criteria.
- D1.3 — `docs/ARCHITECTURE.md` — locked technical architecture (carries forward from original PLAN).
- D1.4 — `docs/FEATURES.md` — explicit specifications for: read-mode AI actions, compose-mode AI actions, RAG ingestion, RULE integration (Synapse Routing rules), contact-based knowledge base (search by sender/recipient).
- D1.5 — `planning/STEPS.md` — 4-sprint step plan, each step has a "Done when…" acceptance criterion and an integration test.
- D1.6 — `planning/GUI_DEFINITIONS.md` — wireframes for: SignIn, ReadMode, ComposeMode, Settings, RuleEditor, ContactKnowledgeBase. Plus design tokens, icon sizes, and Microsoft Store asset list.
- D1.7 — Test plan locked: static / unit / component / E2E / manifest validation layers.
- D1.8 — **Engineering scaffold** (Step 1.6): `AGENTS.md`, `docs/CONTRIBUTING.md`, `docs/COMMIT_PROCESS.md`, `Makefile`, `.github/workflows/ci.yml`, PR + issue templates, dependabot, `.githooks/` (pre-commit + commit-msg + pre-push), `commitlint.config.cjs`, `.editorconfig`, `.gitattributes`, `.gitignore`, `.markdownlint.jsonc`. Mirrors `synaplan/`'s commit-process discipline so Sprint 2 starts on a paved road.

**Acceptance criteria (Sprint 1 exits when all are true)**

1. Every Sprint-2 step has a paragraph in `planning/STEPS.md` with a "Done when…" line that an external engineer could execute against.
2. Every core feature in the user's brief (summarise, translate, answer, RAG ingest, RULE integration, contact knowledge base) is specified in `docs/FEATURES.md` with: input, output, Synaplan endpoint(s) called, UI surface.
3. `docs/ARCHITECTURE.md` references the correct Synaplan endpoints by path; no inventions, no Synaplan-backend code changes assumed for v1.
4. `planning/GUI_DEFINITIONS.md` lists exact icon sizes (16/32/64/80/128) and AppSource asset sizes (256, 512 hero, 1366×768 screenshots).
5. Cross-repo dependencies (above) are explicit and assigned to a sprint.
6. **Engineering scaffold is in place** (D1.8): `make doctor` reports the planning-stage state correctly, `.githooks/commit-msg` rejects non-Conventional-Commits messages, and `.github/workflows/ci.yml` is green on a planning-stage PR.

**Tests this sprint defines (executed in Sprints 2–4)**

| Layer        | Tooling                                                             | Runs on                              |
| ------------ | ------------------------------------------------------------------- | ------------------------------------ |
| Static       | ESLint + Prettier + `vue-tsc -b` + `office-addin-manifest validate` | every commit                         |
| Unit         | Vitest                                                              | every commit                         |
| Component    | Vitest + `@vue/test-utils` with mocked `Office.context`             | every commit                         |
| E2E (OWA)    | Playwright headless + sideload                                      | nightly + before sprint review       |
| Manifest     | `office-addin-validator`                                            | every commit + before AppSource push |
| Manual smoke | New Outlook on Windows, classic Outlook 2024, Outlook on Mac        | end of Sprint 2 and Sprint 3         |

---

## Sprint 2 — GUI, Local Sideload, Live Identification

**Goal:** Ship a sideloadable add-in that the user can install in **their own Outlook**, with the **configuration window** and a working **identification link to `web.synaplan.com`**. Functional features may still call mocked endpoints; the auth + config + sideload loop must be live.

**Duration:** 2–3 weeks.

**Deliverables**

- D2.1 — Scaffolded project (Vite + Vue 3 + TS + Office.js + Yeoman-blessed `manifest.xml`).
- D2.2 — `manifest.xml` with `DisplayName=Synamail`, Mailbox 1.8 + DialogApi 1.1 requirement sets, ribbon entry for ReadMode and ComposeMode, `Permissions=ReadWriteItem`.
- D2.3 — Working sideload script (`npm run sideload`) verified in Outlook on the Web + new Outlook for Windows.
- D2.4 — Vue views: `SignIn.vue`, `ReadMode.vue`, `ComposeMode.vue`, `Settings.vue` (configuration window), `RuleEditor.vue` (stub UI, wired in Sprint 3), `ContactKnowledgeBase.vue` (stub UI, wired in Sprint 3).
- D2.5 — **Settings / configuration window** complete: shows signed-in email + base URL, lets the user change the Synaplan base URL **before** the dialog opens, has a "Sign out" button.
- D2.6 — `useAuth.ts` + `dialog/auth-relay.html` — full Office Dialog API flow with `state` nonce validation and `messageParent` handling.
- D2.7 — **Cross-repo:** new Vue route `/addin/connect` in `synaplan/frontend/` (+ `AddinConnectView.vue`) that calls `POST /api/v1/apikeys` for the signed-in user and posts the issued key back to the taskpane via `Office.context.ui.messageParent`. Built into the Docker image; rolled out by `synaplan-platform` (`docker compose pull && re-startweb{1,2,3}.sh`). Passes Synaplan's pre-commit gate before merge.
- D2.8 — Mock Synaplan client used for read/compose actions in this sprint (Sprint 3 replaces it). The mock ships behind an env flag so the sideloaded build still demos visually.
- D2.9 — CI pipeline (`.github/workflows/ci.yml`): lint, types, manifest validate, unit + component tests, build artefact.

**Acceptance criteria (Sprint 2 exits when all are true)**

1. The user, on their own Outlook, can install the add-in via `npm run sideload` and see a **Synamail** ribbon button in both ReadMode and ComposeMode.
2. Clicking "Sign in" in the taskpane opens the live `https://web.synaplan.com/addin/connect` page, completes a real OAuth round-trip with the user's existing Synaplan provider (Google / GitHub / password / Keycloak / OIDC), and stores a real `apiKey` + `keyId` in `Office.context.roamingSettings`.
3. The new key appears in the user's Synaplan API-keys page as `Outlook Add-in (<host>)`.
4. Settings view shows the right email + base URL; "Sign out" calls `DELETE /api/v1/apikeys/{keyId}` and returns the taskpane to SignIn.
5. "Use a self-hosted instance" link in Settings successfully retargets the next sign-in dialog at a custom Synaplan URL.
6. Reload Outlook → state survives (roaming settings work).
7. CI is green on the branch; manifest passes `office-addin-validator` with zero warnings.

**Tests this sprint produces**

- Unit: `useAuth.ts`, `useRoamingSettings.ts`, `useOutlookItem.ts`, `synaplan-client.ts` (mock + real wrapper).
- Component: `SignIn.vue`, `Settings.vue` (incl. base-URL override and sign-out path).
- E2E (OWA): the full sign-in dance against live `web.synaplan.com` with a dedicated test user. State-nonce rejection scenario included.
- Manifest: `office-addin-validator` on every push.

**Open from the user before Sprint 2 starts**

- Confirmation that the `/addin/connect` Vue route in `synaplan/frontend/` is the preferred home (no `synaplan-website` repo exists).
- A dedicated test Synaplan account on `web.synaplan.com` with at least one RAG group seeded.
- CORS allow-list: confirm the dev sideload origin (`https://localhost:3000`) is permitted by the Synaplan API.

---

## Sprint 3 — Functionality & Live API Calls

**Goal:** Replace every mock with live calls to `web.synaplan.com`. All AI actions in the user's brief work end-to-end with real authentication and real data: summarise, translate, draft reply, classify, save-to-RAG, ask follow-ups, RULE integration, search-by-sender/recipient.

**Duration:** 2–3 weeks.

**Deliverables**

- D3.1 — Real `synaplan-client.ts` (fetch, `X-API-Key`, retry with exponential backoff, 401 → clear key → bounce to SignIn).
- D3.2 — Zod schemas generated from Synaplan's OpenAPI spec at `https://web.synaplan.com/api/doc.json`, checked in (mirror Synaplan's `make -C frontend generate-schemas` workflow).
- D3.3 — **Read-mode features**, wired to live endpoints:
  - **Summarise** → `POST /api/v1/messages/send` with summarise system prompt.
  - **Translate** → `POST /api/v1/messages/send` with translation system prompt + target language picker.
  - **Draft reply** → `messages/send` with reply prompt → `mailbox.displayReplyForm({ htmlBody })`.
  - **Classify** → `messages/send` with classification prompt + JSON output schema.
  - **Ask follow-ups** → `POST /api/v1/chats` (one chat per `mailbox.item.conversationId`) and `GET /api/v1/chats/{id}/messages`.
- D3.4 — **Save to knowledge base (RAG ingestion)** wired to live endpoints:
  - `getAsFileAsync` (Mailbox 1.8) → `POST /api/v1/files/upload` → `POST /api/v1/files/{id}/process`.
  - Group picker driven by `GET /api/v1/files/groups`, **with a default suggestion of `contact:<sender-email>`** to seed the contact knowledge base (see D3.6).
- D3.5 — **Compose-mode features** wired to live endpoints:
  - **Draft from prompt** → `messages/send` → `body.setAsync('html')`.
  - **Improve / shorten / translate selection** → `body.getSelectedDataAsync` → `messages/send` → `body.setSelectedDataAsync`.
  - **Insert from RAG** → `POST /api/v1/rag/search` → click result → `setSelectedDataAsync` with snippet + citation.
- D3.6 — **Contact knowledge base** (search by sender/recipient):
  - When viewing an email, ReadMode shows a "Contact" pill linking to `ContactKnowledgeBase.vue` for the sender (and a recipient picker for multi-recipient threads).
  - The view queries `POST /api/v1/rag/search` filtered to the per-contact group `contact:<email>` and lists snippets + source emails.
  - "Save this email to <contact>'s knowledge base" is a one-click action that calls upload + process with the `contact:<email>` group, creating the group on first use.
  - "Ask about <contact>" opens a chat seeded with the contact's group as RAG scope.
- D3.7 — **RULE integration** (Synapse Routing rules):
  - `RuleEditor.vue` lists existing Synapse Routing rules for the user's topics (read via `GET /api/v1/prompts` or the Synapse admin endpoint — confirm during Sprint 2 wire-up).
  - Lets the user add/remove keyword matchers (`BSELECTION_RULES`) for the topic that should fire when an email matches (e.g. invoice-keyword routes to billing topic).
  - Optional "Apply rule to current email" button previews which topic the current email would route to.
- D3.8 — i18n (en + de) for every UI string — mirrors Synaplan's house rule.

**Acceptance criteria (Sprint 3 exits when all are true)**

1. Summarise, translate, draft reply, classify, ask follow-ups, and save-to-RAG all complete end-to-end against the live test user, with the Playwright E2E suite green.
2. Saving an email to the suggested `contact:<sender>` group creates the group on first use; subsequent uploads accumulate under it.
3. `ContactKnowledgeBase.vue` returns sender-scoped or recipient-scoped RAG results within ≤ 3 s against a seeded test mailbox.
4. `RuleEditor.vue` round-trips rule edits to Synaplan and the rule actually fires (Tier-0 hit visible in routing logs / Synaplan admin UI).
5. Manually revoking the test key from Synaplan's API-keys page causes the **next** action in the add-in to clear roaming settings and return to SignIn.
6. Token-limit / 5xx / 401 / network-down paths all surface human-readable errors (no raw stack traces).
7. CI green; new component + E2E tests added in this sprint pass.

**Tests this sprint produces**

- Unit: prompt templates render deterministically (`summarise`, `translate(lang)`, `reply(tone)`, `classify`, `contact_search`, `rule_match_preview`).
- Component: ReadMode action flows, ComposeMode action flows, ContactKnowledgeBase, RuleEditor.
- E2E (Playwright + OWA, live API):
  - Full Summarise → Ask → Draft reply → Save-to-RAG flow on a seed email.
  - Save to contact group → reopen another email from same sender → contact KB shows the previous email.
  - RULE editor adds a keyword rule → forward an email with that keyword → verify the routed topic.
  - Translate selection in Compose → verify body content updated.
  - 401 path: revoke key, trigger any action, assert SignIn.

**Open from the user before Sprint 3 starts**

- Confirmation that the test Synaplan account has Synapse Routing rules enabled (or guidance if it's an admin-only feature on the live tenant).
- Decision on the exact API key scopes Synamail should request: proposed `messages:*`, `chats:*`, `files:*`, `rag:*`, plus a `synapse:rules:*` if RULE integration needs a dedicated scope.

---

## Sprint 4 — Release: AppSource Publishing

**Goal:** Publish the plugin in the **Microsoft AppSource store** with proper graphics, descriptions, privacy/terms URLs, and a successful certification pass.

**Duration:** 1–2 weeks of work + 4–8 weeks of Microsoft turnaround.

**Deliverables**

- D4.1 — Production hosting of the taskpane SPA at a stable HTTPS origin (proposed `https://addin.synaplan.com`; fallback `https://web.synaplan.com/addin/`).
- D4.2 — Strict CSP on the production host; CORS allow-list on the Synaplan API includes the production origin.
- D4.3 — `manifest.xml` updated to the production `SourceLocation` and `AppDomains`. `manifest.unified.json` generated via `office-addin-manifest-converter` and validated in CI.
- D4.4 — Graphical assets in `assets/store/`:
  - Icons: 16, 32, 64, 80, 128 PNG.
  - Hero: 256 + 512 PNG.
  - Screenshots: 1366×768 PNG of SignIn, ReadMode, ComposeMode, Settings, ContactKnowledgeBase, RuleEditor.
  - Screencast video (≤ 60 s) showing summarise → save-to-RAG → ask follow-up flow.
- D4.5 — Store copy in `assets/store/copy.md`: title, short description, long description, search keywords, support info — en + de.
- D4.6 — Privacy policy + Terms of Service + Support URLs **publicly hosted** as routes in `synaplan/frontend/` (or as static pages in the `synaplan-docs` site) and linked from the manifest.
- D4.7 — Microsoft Partner Center submission (Microsoft 365 and Copilot program).
- D4.8 — Demo Microsoft 365 tenant prepared for certification testers, pre-seeded with a Synaplan account so reviewers can complete the sign-in flow end-to-end.

**Acceptance criteria (Sprint 4 exits when all are true)**

1. `office-addin-validator` returns zero warnings on both `manifest.xml` and `manifest.unified.json`.
2. AppSource certification policy 1100 checklist is fully ticked in `planning/APPSOURCE_CHECKLIST.md` (added in this sprint).
3. The add-in works end-to-end on: Outlook on the Web (Edge + Chrome), new Outlook for Windows, classic Outlook 2024, Outlook on Mac. Smoke tests recorded.
4. The Partner Center submission is accepted **or** Microsoft's review comments are tracked as follow-up tasks and a second submission is in flight.

**Tests this sprint produces**

- `office-addin-validator` runs on `manifest.xml` and `manifest.unified.json` in CI on every push.
- Manual sign-off matrix for the four desktop/web Outlook clients above.
- Production deployment smoke test (Playwright against `https://addin.synaplan.com`).

**Open from the user before Sprint 4 starts**

- Partner Center account access (or onboarding kickoff — verification takes ~1 week).
- Decision on production hostname (`addin.synaplan.com` vs `web.synaplan.com/addin/`).
- Brand name pre-validation: confirm "Synamail" passes Microsoft's name-collision check.

---

## Sprint dependency graph

```
Sprint 1 (planning + specs)
    │
    ▼
Sprint 2 (GUI + sideload + live auth)        ←── cross-repo PR to synaplan/frontend/ + image pull on synaplan-platform
    │
    ▼
Sprint 3 (live AI features + RAG + RULE + contact KB)
    │
    ▼
Sprint 4 (production host + AppSource submission)
```

No sprint depends on a feature from a later sprint. Sprint 2 requires the `synaplan/frontend/` bridge-page PR to be merged **and** the new Docker image rolled onto `synaplan-platform` (web1/web2/web3) before its acceptance criteria pass — schedule that PR in the first half of Sprint 2.
