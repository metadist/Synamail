# synaMail — Outlook Add-in for Synaplan

> Plan and design document. The repository currently contains only a README + LICENSE; this document is the blueprint we will execute against.

## 1. Goal

Build an Outlook add-in (taskpane + ribbon commands) that connects to a user's
Synaplan instance (cloud or self-hosted) so they can, from inside Outlook:

1. Summarise / explain / classify the currently selected email.
2. Generate a draft reply (with tone / length / language controls) and insert it into a compose window.
3. Index ("Save to RAG") an email and/or its attachments into a Synaplan RAG group of the user's choice.
4. Ask follow-up questions in a chat panel anchored to the current message thread.
5. Answer with retrieval grounded in the user's own RAG groups.
6. Optional: trigger Synaplan's existing `smart+keyword@` email-webhook flow by forwarding programmatically.

## 2. Which Outlook release is the "actual" one?

There is no single "current Outlook version" — the supported surface is a matrix of clients × API requirement sets. The relevant facts (verified Q1 2026):

### 2.1 Latest Outlook JavaScript API requirement set

| Requirement set | Release vehicle | Notes |
|---|---|---|
| **Mailbox 1.15** | classic Outlook on Windows v2412 (build 18324.20172) and current Outlook on the web / new Outlook on Windows / Mac (some APIs flagged) | Current GA |
| 1.14 | classic Outlook 2024 LTSC, Mac new UI | |
| 1.5 | Outlook on iOS / Android (subscription) | Hard ceiling for mobile |
| 1.6 | Exchange on-prem 2019 | Hard ceiling for on-prem |

We will **target Mailbox 1.8 as the minimum** (gives us `getAttachmentsAsync`, async body access, `AttachmentsChanged` event, attachment content APIs) and **feature-detect 1.10–1.15** for nice-to-haves. Below 1.8 the add-in just doesn't load; that is acceptable.

### 2.2 Outlook clients to support (web add-in, single codebase)

Web add-ins (Office.js) run on:

* **Outlook on the Web** — Microsoft 365 + Outlook.com (modern UI, full requirement-set support).
* **New Outlook for Windows** — same web-add-in runtime as OWA (this is Microsoft's strategic client; COM/VSTO add-ins are NOT supported here).
* **Classic Outlook for Windows** — Microsoft 365 subscription or perpetual 2021 / 2024.
* **Outlook for Mac** — new UI (since v16.38) and classic UI.
* **Outlook for iOS / Android** — subscription accounts only, capped at 1.5.

We will explicitly target the first four. iOS/Android = best-effort / phase 2.

### 2.3 We do NOT build a COM/VSTO add-in

* COM/VSTO is **deprecated for the new Outlook on Windows** and Microsoft tells partners to migrate. Sources: [Develop add-ins for new Outlook on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook), [Migrate from COM add-ins to web add-ins](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/migrate-com-to-web-addins).
* A web add-in covers all current and future Outlook clients with one codebase.

## 3. Manifest format: XML or unified (JSON)?

Microsoft is migrating from the XML "add-in only" manifest to the **Unified Manifest for Microsoft 365** (JSON, shared with Teams). Status today:

* XML manifest is **still supported** and still required for AppSource on some surfaces.
* Unified manifest is **preferred for the new Outlook on Windows / Microsoft Marketplace** and is what new partner submissions are steered toward.
* Microsoft ships a converter (`office-addin-manifest-converter`) and lets you maintain both side-by-side.

**Decision:** start with **XML manifest** (broadest tooling, easier sideloading, Mac/iOS/Android-ready, every requirement-set version documented for it), and **add a unified-manifest variant** before AppSource submission. That is the path Microsoft itself documents in [Manage both manifests](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/duplicate-legacy-metaos-add-ins).

## 4. Architecture overview

```
┌─────────────────────────── User's Outlook ─────────────────────────────┐
│                                                                          │
│   ┌─ Ribbon button (synaMail) ──┐    ┌─ Event-based handler (1.10+) ─┐  │
│   │ Opens taskpane              │    │ OnNewMessageCompose, Smart-   │  │
│   └────────────┬─────────────────┘    │ Alerts (optional, phase 2)    │  │
│                │                       └───────────────────────────────┘  │
│                ▼                                                          │
│   ┌──────────────────────────────────────────────────────────────────┐   │
│   │  Taskpane SPA (Vue 3 + TS, Vite, Office.js + Fluent UI React v9) │   │
│   │   - Read item: subject, body, from/to, attachments                │   │
│   │   - Compose item: setSelectedDataAsync, setAsync(body)            │   │
│   │   - Stores per-mailbox config in Office.context.roamingSettings   │   │
│   │   - Optional EWS / Identity-API token for delegated SSO           │   │
│   └────────────────────────┬──────────────────────────────────────────┘   │
└────────────────────────────│──────────────────────────────────────────────┘
                             │  HTTPS, JSON, X-API-Key header
                             ▼
                ┌──────────────────────────────┐
                │  Synaplan (cloud or self-    │
                │  hosted) — already exposes   │
                │  the API surface we need.    │
                │                              │
                │  /api/v1/messages/send       │
                │  /api/v1/files/upload        │
                │  /api/v1/files/{id}/process  │
                │  /api/v1/rag/search          │
                │  /api/v1/chats   (CRUD)      │
                │  /api/v1/webhooks/email      │
                └──────────────┬───────────────┘
                               │
                               ▼
              ┌────── optional Synaplan plugin ──────┐
              │ plugins/synamail/                    │
              │   manifest.json                      │
              │   backend/Controller/                │
              │     SynaMailController.php           │
              │     - /reply-suggest                 │
              │     - /classify-mail                 │
              │     - /index-mail                    │
              │   migrations/001_setup.sql           │
              │     (BCONFIG group "P_synamail")     │
              └──────────────────────────────────────┘
```

Two pieces of code, two repos:

| Repo | What lives there |
|---|---|
| **`synaMail/`** (this repo) | Outlook web add-in: manifest, taskpane SPA, build, AppSource assets, sideloading scripts, Playwright tests |
| **`synaplan/plugins/synamail/`** *(optional, phase 1.5)* | Server-side companion plugin that adds purpose-built endpoints (`/reply-suggest`, `/index-mail`) on top of the generic Synaplan API. Same non-invasive pattern already used by `castingdata` and `hello_world`. |

Phase 1 only needs the add-in; it can talk directly to existing Synaplan endpoints. The companion plugin is a phase-1.5 optimisation.

## 5. Tech stack

| Layer | Choice | Why |
|---|---|---|
| Add-in scaffolding | Yeoman generator `yo office` (Outlook taskpane, TypeScript) | Microsoft-blessed starting point, gives manifest + dev cert + sideload scripts |
| Build | **Vite** + Office.js | We already use Vite across the Synaplan stack; faster than the default webpack template |
| UI | **Vue 3 + TypeScript** (consistent with synaplan frontend) **or** Fluent UI React v9 (Microsoft's official) | Pick Vue 3 to match house style; wrap Fluent design tokens via CSS vars. Final decision in spike (§13) |
| HTTP | `fetch` + Zod schemas — *re-use generated schemas* from `synaplan/frontend/src/generated/api-schemas.ts` so the contract stays in sync | Same pattern documented in `docs/API_PATTERNS.md` |
| Auth | API-key (Synaplan style: `X-API-Key`). Stored encrypted in `Office.context.roamingSettings` (per-mailbox) | Matches existing Synaplan auth; no OAuth dance in v1 |
| Hosting | Static SPA on Cloudflare Pages or `synaplan-platform` (HTTPS mandatory). Service URL hard-coded in manifest | Manifest needs an SSL URL |
| Tests | Vitest (unit), Playwright (E2E against Outlook on the Web), `office-addin-validator` (manifest), `office-addin-debugging` (sideload) | All Microsoft-recommended |
| CI | GitHub Actions: lint, manifest validate, unit, build, deploy preview | |

### 5.1 Why not COM / VSTO / Outlook MAPI plugin

Already discussed in §2.3. Adding here for emphasis: a single Office.js add-in covers all five client surfaces; a COM add-in covers exactly one (classic Windows) and is being phased out.

## 6. Repository layout (target)

```
synaMail/
├── README.md
├── PLAN.md                       (this file)
├── LICENSE
├── package.json
├── tsconfig.json
├── vite.config.ts
├── manifest.xml                  Outlook XML manifest (primary)
├── manifest.unified.json         Unified Microsoft 365 manifest (parallel)
├── assets/                       icons (16/32/64/80/128), screenshots, store copy
│   ├── icon-*.png
│   └── store/                    AppSource hero / video / privacy URLs
├── src/
│   ├── taskpane/
│   │   ├── index.html
│   │   ├── main.ts               Office.onReady, root mount
│   │   ├── App.vue
│   │   ├── views/
│   │   │   ├── ReadMode.vue      summary, draft-reply, save-to-RAG, ask
│   │   │   ├── ComposeMode.vue   draft helper, tone shifter, attach-from-RAG
│   │   │   └── Settings.vue      API URL, API key, default RAG group, language
│   │   ├── components/
│   │   ├── composables/
│   │   │   ├── useOutlookItem.ts current item, attachments, body
│   │   │   ├── useSynaplan.ts    typed client over httpClient
│   │   │   └── useRoamingSettings.ts
│   │   └── styles/
│   ├── commands/
│   │   ├── commands.html         Office function-file shell
│   │   └── commands.ts           ribbon button handlers (no UI), Smart Alerts
│   ├── shared/
│   │   ├── synaplan-client.ts    fetch wrapper, X-API-Key, retries
│   │   ├── schemas.ts            Zod (re-export from synaplan-generated)
│   │   ├── prompts.ts            reply / summarise / classify prompt templates
│   │   └── types.ts
│   └── dialog/
│       └── consent.html          first-run "paste your API key" dialog
├── tests/
│   ├── unit/                     Vitest
│   └── e2e/                      Playwright + Outlook on the Web
├── scripts/
│   ├── sideload.sh               wraps office-addin-debugging
│   ├── validate-manifest.sh
│   └── generate-icons.sh
└── .github/workflows/
    ├── ci.yml
    └── deploy-preview.yml
```

## 7. Synaplan API surface we rely on (existing, no core changes)

All endpoints accept `X-API-Key`. Reference: `synaplan/backend/src/Controller/`.

| Use case | Endpoint | Notes |
|---|---|---|
| Send chat / get AI answer | `POST /api/v1/messages/send` | Used for summarise, draft reply, classify, ask. We pass the email body + system prompt. |
| Multi-turn chat for a thread | `POST /api/v1/chats`, `GET /api/v1/chats/{id}/messages` | Pin a chat to an Outlook conversation by storing `chatId` in roamingSettings keyed by `conversationId`. |
| Upload attachment / mail body to RAG | `POST /api/v1/files/upload` then `POST /api/v1/files/{id}/process` | "Save to knowledge base" button. |
| Retrieval | `POST /api/v1/rag/search` | Used in compose-mode "find relevant prior emails". |
| Existing email-bridge (optional fallback) | `POST /api/v1/webhooks/email` | We could forward instead of calling per-action endpoints. |

If any add-in scenario benefits from a tailored endpoint (e.g. "given an email and a thread, return a drafted reply with citations"), we add it as **`plugins/synamail/`** in the Synaplan repo following the existing plugin pattern (manifest.json + Controller + BCONFIG `P_synamail`). We do this only when the generic endpoints aren't enough.

## 8. Feature design

### 8.1 First-run / settings

* On first activation, taskpane checks `Office.context.roamingSettings` for `synaplan.apiUrl` + `synaplan.apiKey`.
* If missing → renders `Settings` view: API URL (default `https://web.synaplan.com`), API key, "Test connection" button (calls `GET /api/v1/health` or `/api/v1/profile`).
* Roaming settings persist per mailbox across all Outlook clients for that user.

### 8.2 Read mode (selected message in reading pane)

Buttons (Fluent):

1. **Summarise** — `messageread.body.getAsync('text')` → `messages/send` with system prompt = "Summarise in N bullets, German/English/auto".
2. **Draft reply** — open compose window with `mailbox.displayReplyForm({ htmlBody, attachments })` after Synaplan returns an answer. Provide tone selector (formal/concise/friendly) + language selector.
3. **Classify** — sends body + headers to Synaplan, displays category + confidence. (We can re-use the SortX-style category model.)
4. **Save to knowledge base** — uploads `getAsFileAsync` (full .eml) and/or each attachment to `/files/upload`, then `/files/{id}/process` with chosen RAG group.
5. **Ask…** — text box, anchored to a chat that is keyed by `mailbox.item.conversationId`. So replies in the same thread continue the same Synaplan chat.

### 8.3 Compose mode (writing or replying)

* Detect via `Office.context.mailbox.item.itemType` and the Compose interfaces.
* Buttons:
  1. **Draft from prompt** — user types intent, Synaplan generates body, `body.setAsync(html, { coercionType: 'html' })`.
  2. **Improve / shorten / translate** — selection-based: `body.setSelectedDataAsync`.
  3. **Insert from RAG** — search bar → `/rag/search` → click result → inline citation + content insert.
* Optional **Smart Alerts (`OnMessageSend`)** in phase 2: warn before sending if Synaplan policy flags PII / off-topic / missing attachment. Requires Mailbox 1.12+.

### 8.4 Event-based activation (phase 2, requires 1.10+)

* `OnNewMessageCompose` event auto-opens the assistant when the user starts a reply, pre-populating a "Draft reply" suggestion based on the original email — gated behind a per-user toggle in Settings.
* `OnMessageSend` Smart Alerts integration for outbound governance (optional).

### 8.5 Mobile (phase 3)

iOS/Android cap at requirement set 1.5 → reduced feature set: summarise + ask only (no compose helpers). Same code, conditional UI.

## 9. Security, privacy, compliance

This is the section that must pass AppSource validation, so we design for it from day 1.

* **Transport** — every URL in the manifest and every API call is HTTPS only.
* **Secrets** — API key never logged; stored in `roamingSettings` (encrypted at rest by Exchange). Optional: support **Nested App Auth (NAA) / Identity API** in phase 2 to remove the API-key paste step for Synaplan-cloud users who sign in with Microsoft.
* **Data minimisation** — body and attachments leave Outlook only on explicit user action (no autoload). Default to "Extract Only" RAG processing level; Vectorise / Full Analysis only on explicit click.
* **Telemetry** — opt-in only. No PII leaves the client without the user's action.
* **Permissions** — XML manifest `Permissions` element set to `ReadWriteItem` (we need to write reply bodies). Justify in store listing.
* **Privacy policy + terms URLs** — required in manifest and Partner Center, hosted under `synaplan-website` repo.
* **Accessibility** — Fluent components + WCAG 2.1 AA; keyboard nav; high-contrast theme; tested against AppSource accessibility checklist.
* **CSP** — strict CSP on the taskpane host; no third-party scripts inside the taskpane.
* **Tenant admin install** — manifest must work for centrally-deployed installs (Microsoft 365 admin centre); tested via Integrated Apps page.

## 10. Roadmap (phased)

| Phase | Outcome | Time-box |
|---|---|---|
| **0 — Spike** | `yo office` scaffold, sideload to Outlook on the Web + new Outlook for Windows, hello-world taskpane reads the selected message, calls `messages/send` against a local synaplan dev container, prints reply. Decide Vue vs React. | 1–2 days |
| **1 — Read-mode MVP** | Settings flow, Summarise + Ask + Draft-reply + Save-to-RAG. Manifest XML only. Self-hosted load via shared link. | 2–3 weeks |
| **1.5 — Synaplan plugin** | Add `plugins/synamail/` to the Synaplan repo with `/reply-suggest`, `/classify-mail`, `/index-mail` for cleaner contracts. Use idempotent BCONFIG seed for default prompts. | 3–5 days |
| **2 — Compose mode + Smart Alerts** | Draft helpers, RAG insert, optional `OnMessageSend` policy guard, optional `OnNewMessageCompose` autoload. | 2 weeks |
| **3 — Distribution** | Unified manifest variant, store assets, Partner Center submission, validation iterations (4–6 weeks Microsoft turnaround). | 4–8 weeks calendar |
| **4 — Mobile + on-prem Exchange** | Verify against Mailbox 1.5 / Exchange SE, restricted feature set on iOS/Android. | 1 week dev + verification |

## 11. Test plan

Tests are mandatory before each merge. We aim for parity with the synaplan house rule: "lint + static + tests must pass before commit".

### 11.1 Static / build gates (CI on every PR)

1. `npm run lint` — ESLint + Prettier (Vue + TS).
2. `npm run check:types` — `vue-tsc -b` (catches things ESLint misses, just like the Synaplan rule says).
3. `npx office-addin-manifest validate manifest.xml`
4. `npx office-addin-manifest validate manifest.unified.json`
5. `npm run build` — Vite production build, fail if bundle > size budget.
6. License + dependency scan.

### 11.2 Unit (Vitest)

* `synaplan-client.ts` — header injection, retry, error mapping, Zod parse failures.
* `prompts.ts` — every prompt template renders deterministically.
* `useOutlookItem.ts` — given a mocked `Office.context.mailbox.item`, returns expected shape (read + compose modes).
* `useRoamingSettings.ts` — get/set/remove round-trip with stubbed Office.

### 11.3 Component (Vitest + @vue/test-utils)

* `ReadMode.vue` — buttons disabled until config valid; shows loading + error states; renders summary; "Save to RAG" calls upload + process in order.
* `ComposeMode.vue` — `setAsync` / `setSelectedDataAsync` invoked with the right coercion type.
* `Settings.vue` — "Test connection" handles 200/401/network.

### 11.4 E2E (Playwright + Outlook on the Web)

We'll use the [office-addin-test-helpers](https://github.com/OfficeDev/office-js-test-helpers) approach + Playwright headless.

| Scenario | Steps | Pass criteria |
|---|---|---|
| Sideload | Run `npm run sideload`; assert ribbon button appears in OWA | Manifest accepted, button visible |
| First-run | Open taskpane, no key configured | Settings view shown |
| Test connection — happy path | Enter key, click Test | "Connected as <email>" |
| Test connection — bad key | Enter junk, click Test | Error toast |
| Summarise | Open a known seed email, click Summarise | Summary appears within N seconds, matches snapshot prefix |
| Draft reply | Click Draft reply | Reply compose window opens with synaplan-generated HTML body |
| Compose body insert | Open new message, type intent, click Draft from prompt | `body` value contains generated text |
| Save to RAG | Click Save to RAG, pick group | `/files/upload` + `/files/{id}/process` invoked, success toast |
| Ask thread | Send three follow-ups | All hit the same `chatId`, history visible |
| Settings persistence | Reload Outlook | Key still present, no re-prompt |
| 1.8 fallback | Run on Outlook 1.8 simulator | Compose helpers gated, read-mode features work |

Run the suite against the matrix:

* Outlook on the Web (Edge, Chrome, Firefox, Safari)
* New Outlook on Windows (manual smoke once per release)
* Classic Outlook 2024 on Windows VM (manual smoke)
* Outlook on Mac (manual smoke)
* Outlook iOS / Android (manual smoke, phase 3)

### 11.5 Manifest / store validation

* Run **`office-addin-validator`** before every push and as part of CI — it uses the same engine Microsoft uses during certification.
* Run **Office Add-in Validation** sample via the [Validator API](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest) before submission.
* Manual checklist against [Microsoft Marketplace certification policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies) (esp. policies 1100 — Office Add-ins).

### 11.6 Synaplan-side regression

When we add `plugins/synamail/` to the Synaplan repo, we **must** keep its workflow:

```bash
make -C backend lint && make -C backend phpstan && make -C backend test
```

This is the rule from `synaplan/AGENTS.md` — non-negotiable before any plugin commit.

## 12. Publishing — Microsoft Marketplace (AppSource)

Reference: [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource), [Publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).

### 12.1 Pre-requisites

1. **Partner Center account** in the "Microsoft 365 and Copilot" program (same account works for AppSource and Teams Store). One-off setup, ~1 week verification.
2. **Privacy policy + Terms of Service** URLs hosted publicly without auth — go in `synaplan-website` repo.
3. **Support URL** (likewise public).
4. Stable HTTPS host for the taskpane — `https://addin.synaplan.com` (or under `web.synaplan.com/addin/`).
5. **AppSource icons** in all required sizes (16, 32, 64, 80, 128 PNG) plus a 256/512 hero image.
6. Demo / screencast video covering the add-in in action.
7. Demo Microsoft 365 tenant for the certification testers.

### 12.2 Manifest requirements (XML for legacy store, unified for new store)

* `Id` — generate a stable GUID (one per environment: dev / staging / prod).
* `Version` — `a.b.c.d` four-part version, must match the form submitted in Partner Center.
* `SupportUrl`, `IconUrl`, `HighResolutionIconUrl`, `SourceLocation` — all HTTPS.
* `Permissions` — least privilege (`ReadWriteItem`).
* `Hosts` — only `Mailbox`.
* `Requirements` — `Mailbox` minVersion 1.8.
* `AppDomains` — list every external domain the taskpane will load (synaplan API host).

### 12.3 Submission checklist (we'll keep this in the repo as `docs/APPSOURCE_CHECKLIST.md` later)

* [ ] Manifest passes `office-addin-validator` with zero warnings.
* [ ] Add-in works without errors in Outlook on the Web, new Outlook on Windows, classic Outlook 2024, Outlook on Mac (new + classic).
* [ ] No telemetry without consent.
* [ ] Privacy policy + Terms + Support URLs all publicly reachable.
* [ ] All UI strings localised (en + de minimum, matching Synaplan).
* [ ] Functions advertised in the listing actually work.
* [ ] Unique add-in ID + matching version number in form and manifest.
* [ ] Demo account credentials provided to validators.
* [ ] AppSource certification policy 1100 (Office Add-ins) — every line ticked.
* [ ] Microsoft 365 App Compliance Program (recommended, helps installer flow).

### 12.4 Timeline expectation

* First validation pass typically returns issues — Microsoft documents 4–6 weeks until full approval, sometimes more for first-time submitters.
* We plan **two submission iterations** in the schedule.

### 12.5 Distribution alternatives during dev/beta

While AppSource is in flight we use:

* **Microsoft 365 admin centre — Integrated Apps** for tenant-wide deployment to friendly customers.
* **Manifest sideloading** for individual testers (we ship a one-line `npm run sideload` script and a docs page).
* **Direct manifest URL** sharable to anyone with sideload permission.

## 13. Open questions / decisions to make in spike

1. **UI library**: Vue 3 (matches synaplan house style) vs Fluent UI React v9 (matches Outlook). Recommendation: Vue 3 + Fluent design tokens via CSS, but worth a 1-day spike to compare bundle size and ergonomics inside Office.js.
2. **Auth**: API key only in v1, vs introducing Nested App Auth (NAA) so Microsoft-signed-in users get SSO into Synaplan. NAA is great UX but only works for Synaplan accounts that already federate with Microsoft. Recommendation: API key in v1, NAA in phase 2.
3. **Companion plugin in Synaplan**: do we actually need `plugins/synamail/` or are the existing endpoints enough? Defer to end of phase 1 — only add if endpoint contracts get awkward.
4. **Brand / store name**: AppSource may push back on the name "synaMail" (case mixing). Need to pre-validate.
5. **Self-hosted Synaplan support**: many users run their own. Add-in must let them paste any base URL — already in the design — but we need a CORS check (Synaplan FrankenPHP must allow the add-in origin). Confirm during spike.

## 14. Concrete next actions (in order)

1. `cd synaMail && npx --package yo --package generator-office -- yo office` → choose **Outlook Add-in**, **TypeScript**, **Custom** (we strip the default React layer and bring our own).
2. Replace the default webpack config with **Vite** + Vue 3 (or commit to Fluent UI React after spike).
3. Add `manifest.xml` with `Mailbox 1.8` + a temporary `https://localhost:3000` source location.
4. `npm run sideload` and verify the ribbon button shows in Outlook on the Web with a Microsoft 365 dev tenant.
5. Wire `useSynaplan` against a local synaplan stack (`docker compose up -d` in `synaplan/`).
6. Implement Settings → Summarise → Draft-reply → Save-to-RAG → Ask, in that order.
7. Set up GitHub Actions: lint, type-check, manifest validate, unit tests, build artefact.
8. Begin AppSource asset preparation in parallel (icons, screenshots, copy, privacy policy hosting).

## 15. References (verified Q1 2026)

* [Outlook JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) — current set 1.15, full client matrix.
* [Develop Outlook add-ins for the new Outlook on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook) — confirms web add-ins are the supported model and COM/VSTO is deprecated for new Outlook.
* [Migrate from COM add-ins to web add-ins](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/migrate-com-to-web-addins).
* [Office Add-ins with the unified app manifest for Microsoft 365](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview).
* [Compare the add-in only manifest with the unified manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/json-manifest-overview).
* [Get an Outlook item's attachments from Exchange](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item) — `getAttachmentsAsync`, `getAttachmentContentAsync`, requirement set 1.8.
* [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource).
* [Microsoft Marketplace publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).
* [Microsoft Marketplace certification policies (1100 — Office Add-ins)](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies).

Synaplan side:

* `synaplan/docs/EMAIL.md` — existing email-webhook flow (`smart+keyword@`).
* `synaplan/docs/RAG.md` — file upload + processing levels + groups.
* `synaplan/docs/API_PATTERNS.md` — Zod + OpenAPI contract pattern we will follow on the add-in's HTTP client.
* `synaplan/plugins/hello_world/` and `synaplan/plugins/castingdata/` — non-invasive plugin template if/when we add `plugins/synamail/`.
