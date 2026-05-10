# Synamail — Outlook Add-in for Synaplan

> Plan and design document. The repository currently contains only a README + LICENSE; this document is the blueprint we will execute against.
>
> The repo on disk is at `~/wwwroot/synaMail/` (filesystem path is what it is) — but the **product/brand name is "Synamail"** and that is what appears in the manifest, the AppSource listing, the ribbon button, the taskpane title, and every user-facing string.

## 1. Goal

Build an Outlook add-in (taskpane + ribbon commands) that connects to a user's
Synaplan instance (cloud or self-hosted) so they can, from inside Outlook:

1. Summarise / explain / classify the currently selected email.
2. Generate a draft reply (with tone / length / language controls) and insert it into a compose window.
3. Index ("Save to RAG") an email and/or its attachments into a Synaplan RAG group of the user's choice.
4. Ask follow-up questions in a chat panel anchored to the current message thread.
5. Answer with retrieval grounded in the user's own RAG groups.
6. Optional: trigger Synaplan's existing `smart+keyword@` email-webhook flow by forwarding programmatically.

**No copy-paste during onboarding.** First launch shows a single "Sign in" button that opens Synaplan's existing login (Google / GitHub / password / Keycloak / OIDC), and the add-in is connected automatically. See §7.5.

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

We additionally require **DialogApi 1.1** and (later, when we add NAA) **NestedAppAuth 1.1**.

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
│   ┌─ Ribbon button (Synamail) ──┐    ┌─ Event-based handler (1.10+) ─┐  │
│   │ Opens taskpane              │    │ OnNewMessageCompose, Smart-   │  │
│   └────────────┬─────────────────┘    │ Alerts (optional, phase 2)    │  │
│                │                       └───────────────────────────────┘  │
│                ▼                                                          │
│   ┌──────────────────────────────────────────────────────────────────┐   │
│   │  Taskpane SPA  (Vue 3 + TS, Vite, Office.js, Fluent design tokens)│  │
│   │   - Read item: subject, body, from/to, attachments                │   │
│   │   - Compose item: setSelectedDataAsync, setAsync(body)            │   │
│   │   - Stores apiKey + apiUrl in Office.context.roamingSettings      │   │
│   │     (encrypted at rest by Exchange, per-mailbox)                  │   │
│   │   - Sign-in via Office.context.ui.displayDialogAsync (§7.5)       │   │
│   └────────────────────────┬──────────────────────────────────────────┘   │
└────────────────────────────│──────────────────────────────────────────────┘
                             │  HTTPS, JSON, X-API-Key header
                             ▼
                ┌──────────────────────────────┐
                │  Synaplan (cloud or self-    │
                │  hosted) — uses ONLY the     │
                │  endpoints already shipped.  │
                │                              │
                │  Auth (existing):            │
                │   /api/v1/auth/google/login  │
                │   /api/v1/auth/github/...    │
                │   /api/v1/auth/keycloak/...  │
                │   /api/v1/auth/oidc/...      │
                │   /api/v1/apikeys (POST)     │
                │                              │
                │  Functional (existing):      │
                │   /api/v1/messages/send      │
                │   /api/v1/files/upload       │
                │   /api/v1/files/{id}/process │
                │   /api/v1/rag/search         │
                │   /api/v1/chats              │
                └──────────────────────────────┘
```

One repo, one codebase, **zero Synaplan-side code changes required for v1** (we add a tiny `/addin/connect` static page on the Synaplan frontend later in phase 1 to make the OAuth dialog flow seamless — see §7.5).

## 5. Tech stack

| Layer | Choice | Why |
|---|---|---|
| Add-in scaffolding | Yeoman generator `yo office` (Outlook taskpane, TypeScript) | Microsoft-blessed starting point; gives manifest + dev cert + sideload scripts |
| Build | **Vite** + Office.js | Matches Synaplan's existing stack; faster than the default webpack template |
| UI framework | **Vue 3 + TypeScript** (locked) | House style across the Synaplan platform; consistent with `synaplan/frontend/`; easier to share components long-term |
| Design tokens | Fluent UI design tokens via CSS variables (no React dependency) | Native-looking inside Outlook without pulling in the Fluent React tree |
| HTTP | `fetch` + Zod schemas — *re-use generated schemas* from `synaplan/frontend/src/generated/api-schemas.ts` so the contract stays in sync with the backend OpenAPI spec | Same pattern documented in `docs/API_PATTERNS.md` |
| Auth | OAuth via Office Dialog API → Synaplan's existing login (Google / GitHub / Keycloak / OIDC / password) → API key issued automatically; stored encrypted in `Office.context.roamingSettings` | See §7.5. Phase 2: optional Nested App Auth (NAA) for silent SSO |
| Hosting | Static SPA on Cloudflare Pages or `synaplan-platform` (HTTPS mandatory). Service URL hard-coded in manifest | Manifest needs an SSL URL |
| Tests | Vitest (unit), Playwright (E2E against Outlook on the Web), `office-addin-validator` (manifest), `office-addin-debugging` (sideload) | All Microsoft-recommended |
| CI | GitHub Actions: lint, manifest validate, unit, build, deploy preview | |

### 5.1 Why not COM / VSTO / Outlook MAPI plugin

Already discussed in §2.3. Adding here for emphasis: a single Office.js add-in covers all five client surfaces; a COM add-in covers exactly one (classic Windows) and is being phased out.

## 6. Repository layout (target)

```
synaMail/                         (filesystem; product is "Synamail")
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
│   │   │   ├── SignIn.vue        first-run; calls openSignInDialog()
│   │   │   ├── ReadMode.vue      summary, draft-reply, save-to-RAG, ask
│   │   │   ├── ComposeMode.vue   draft helper, tone shifter, attach-from-RAG
│   │   │   └── Settings.vue      switch tenant, sign out, default RAG group, language
│   │   ├── components/
│   │   ├── composables/
│   │   │   ├── useOutlookItem.ts current item, attachments, body
│   │   │   ├── useSynaplan.ts    typed client over httpClient
│   │   │   ├── useAuth.ts        sign-in / sign-out / token refresh
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
│       ├── auth-relay.html       loaded by displayDialogAsync; redirects to
│       │                         Synaplan and listens for the bridge postMessage
│       └── auth-relay.ts
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
| **Auth — Google / GitHub / Keycloak / OIDC / password** | `GET /api/v1/auth/{provider}/login` (existing) | Sets auth cookies on the frontend. We open this in the Office dialog. |
| **Auth — issue API key** | `POST /api/v1/apikeys` (existing) | Authenticated user creates a named, scoped API key; returned **once**. We use this from the bridge page below. |
| **Auth — revoke API key** | `DELETE /api/v1/apikeys/{id}` (existing) | "Sign out from this Outlook" calls this with the stored key id. |

> **No companion `plugins/synamail/` is planned.** All add-in features run on existing endpoints. If a single endpoint contract ever becomes awkward, we revisit; until then, no Synaplan-side code changes.

## 7.5 Authentication & onboarding (no copy-paste)

This is the section the user asked us to expand. Goal: zero copy-paste, the user signs into Synaplan once (using whichever provider Synaplan already supports — Google included) and Outlook is connected for good.

### 7.5.1 What was decided

* **Phase 1 (ship-it):** OAuth via the **Office Dialog API**. The user clicks "Sign in" once; a popup hosts Synaplan's existing login page (Google / GitHub / password / Keycloak / OIDC, whatever the Synaplan instance has enabled); after sign-in a tiny Synaplan-hosted bridge page (`/addin/connect`) issues an API key and posts it back to the taskpane. **No password, no key, nothing for the user to type or paste.**
* **Phase 2 (silent SSO):** **Nested App Auth (NAA)** for users whose Outlook is signed in to a Microsoft 365 work/school account whose email matches a Synaplan account. NAA gets a Microsoft Entra token silently; Synaplan validates it (we add Microsoft Entra to the existing OIDC controller stack) and issues an API key without any popup at all. NAA does not work for Google identities, so the dialog flow remains the universal fallback.

### 7.5.2 Why NOT plain NAA in v1

The user wrote: *"can NAA be done with a simple login to Synaplan (maybe via Google login) and the accounts are all connected?"* — that conflates two different things, so to be precise:

* **NAA** issues a token from **Microsoft Entra ID** (the user's Microsoft account). It tells Synaplan "this is the M365 user `alice@acme.com`". Silent — zero clicks. But:
  * It can never tell Synaplan "this is the user signed in with **Google**". Google identity isn't part of NAA.
  * It requires Synaplan to trust Microsoft Entra (new auth provider work, new app registration in Microsoft Entra, new validation logic).
  * It only matches an existing Synaplan account if the Microsoft email = the Synaplan email. If the user signed up to Synaplan with `alice@gmail.com` but their Outlook is `alice@acme.com`, NAA can't link them on its own.

* **Office Dialog API + existing Synaplan login** (what we ship in v1) lets the user sign in with **any** provider Synaplan already supports — including the Google login they specifically mentioned. One click. No copy-paste. Works on every Outlook client that supports Mailbox 1.1 + DialogApi 1.1 (which is all of them in our target matrix).

So we ship the Dialog flow first (covers Google, GitHub, password, Keycloak, OIDC — every Synaplan auth provider for free), and add NAA on top later as a "silent upgrade" for the M365-mail-matches-Synaplan-account case.

### 7.5.3 Phase-1 sign-in flow, step by step

```
Outlook (taskpane)                Synaplan web                            Synaplan API
─────────────────                 ──────────────                          ───────────────
1. User clicks "Sign in"
2. Office.context.ui
   .displayDialogAsync(
      'https://web.synaplan
       .com/addin/connect?
       state=<nonce>&
       label=Outlook+Add-in')
                          ────────►  /addin/connect
                                     - if not logged in →
                                       redirect to /login
                                       (Google / GitHub /
                                        password / Keycloak /
                                        OIDC — existing UI)
                                     - user signs in (existing
                                       cookie session)
                                     - /addin/connect renders:
                                       "Connect this Outlook to
                                        <user@email>?"
                                       [Connect]
                                     - on Connect:
                                            POST /api/v1/apikeys
                                              { name: 'Outlook
                                                  on Windows
                                                  (<host>)',
                                                scopes:
                                                  ['addin:*'] }
                                                                            ────────►
                                                                            201
                                                                            { id, key, ... }
                                                                            ◄────────
                                     - tiny inline JS:
                                       Office.context.ui
                                        .messageParent(
                                          JSON.stringify({
                                            apiKey, keyId,
                                            email, baseUrl,
                                            state }))
3. Dialog 'message'
   handler validates
   `state`, stores
   { apiKey, keyId,
     email, baseUrl }
   in roamingSettings
4. Dialog auto-closes;
   taskpane renders
   ReadMode/ComposeMode
```

Key UX points:

* The user sees their **own Synaplan login screen** — the one they already know — branded like Synaplan, not like our add-in. That builds trust (no "third-party app asking for my password" anti-pattern).
* If the user is already signed into Synaplan in their default browser, Synaplan won't even prompt them; the dialog flashes through and `messageParent` fires within a second.
* "Connect" gives them an explicit confirmation step — required by AppSource certification policy 1100.
* The bridge page is plain HTML + a few lines of JS — **the only Synaplan-side change for v1**. It's a static page that calls the existing `/api/v1/apikeys` endpoint. No new API.

### 7.5.4 Account linking

Account linking is implicit in this flow:

* In phase 1: **the user picks** which Synaplan account this Outlook is linked to, by signing in. It can be a Google-backed Synaplan account, a password-backed one, an enterprise Keycloak one — Synamail doesn't care, it only stores the API key.
* "Outlook email" and "Synaplan email" do **not** have to match. That is a feature, not a bug, and matches how Synaplan already handles channels (an email account can drive a Synaplan account regardless of mail address — see `docs/EMAIL.md`).
* In phase 2 (NAA): if the M365 mailbox email *does* match a Synaplan account (any provider), NAA links them silently. If the user wants a different mapping, they sign out and use the dialog flow.

### 7.5.5 Sign-out & key revocation

* Settings view shows "Signed in as `<email>` on `<baseUrl>`. **Sign out**".
* Sign-out calls `DELETE /api/v1/apikeys/{keyId}` with the stored `keyId` and clears `roamingSettings`.
* Synaplan's existing API-keys page lists the key as `Outlook Add-in (<host>)` so power users can revoke from any browser.

### 7.5.6 Phase 2: NAA upgrade path

When we add NAA later:

* Add-in calls `Office.auth.getAccessToken({ allowSignInPrompt: false, forMSGraphAccess: false })` (NAA).
* On success, POST `/api/v1/auth/microsoft/nested` (new endpoint, Synaplan-side) with the bearer token.
* Synaplan validates the Microsoft Entra ID token, looks up a user by email; if found, issues an API key and returns it; if not found, falls back to the dialog flow.
* This becomes the silent first-launch path for the matching-email case. Everyone else (Google-only Synaplan users, mismatched emails) continues to use the dialog.

### 7.5.7 Token lifecycle

* API key has no expiry by default (matches Synaplan's existing behaviour). We can opt into a 1-year scoped key once Synaplan supports `expiresAt` on `/apikeys` (not blocking).
* On any 401 from Synaplan, the add-in clears its stored key and returns the user to the SignIn view. They click "Sign in" again, the dialog flashes through if they're still in their browser session.

## 8. Feature design

### 8.1 First run

1. Taskpane mounts, reads `Office.context.roamingSettings` for `synamail.apiKey` + `synamail.baseUrl`.
2. Missing → render `SignIn.vue`: single button "Sign in to Synaplan". Below it a small "Use a self-hosted instance →" link that lets the user override the base URL **before** opening the dialog (still no copy-paste of secrets).
3. Click → `Office.context.ui.displayDialogAsync(...)` (see §7.5.3).
4. On success → store, render the appropriate mode view.

### 8.2 Read mode (selected message in reading pane)

Buttons (Fluent-styled):

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
* **Secrets** — API key never logged; stored in `Office.context.roamingSettings` (encrypted at rest by Exchange). Cleared on sign-out and on persistent 401.
* **Auth** — OAuth via the Office Dialog API (§7.5). The user never types or pastes a credential into the taskpane; they only see Synaplan's own login UI inside the dialog. Bridge page validates a `state` nonce supplied by the add-in, exactly like a normal OAuth client.
* **Origin allow-list** — `auth-relay.html` only accepts `messageParent` events from the configured Synaplan origin (default `https://web.synaplan.com`, override via "use a self-hosted instance"). The dialog also passes `origin` in `displayDialogAsync` options where supported.
* **Data minimisation** — body and attachments leave Outlook only on explicit user action (no autoload). Default to "Extract Only" RAG processing level; Vectorise / Full Analysis only on explicit click.
* **Telemetry** — opt-in only. No PII leaves the client without the user's action.
* **Permissions** — XML manifest `Permissions` element set to `ReadWriteItem` (we need to write reply bodies). Justify in store listing.
* **Privacy policy + terms URLs** — required in manifest and Partner Center, hosted under `synaplan-website` repo.
* **Accessibility** — Vue 3 + Fluent design tokens, WCAG 2.1 AA; keyboard nav; high-contrast theme; tested against AppSource accessibility checklist.
* **CSP** — strict CSP on the taskpane host; no third-party scripts inside the taskpane.
* **Tenant admin install** — manifest must work for centrally-deployed installs (Microsoft 365 admin centre); tested via Integrated Apps page.
* **CORS** — confirmed acceptable on the Synaplan side; the taskpane origin (`https://addin.synaplan.com` or wherever we host) is added to the API's `Access-Control-Allow-Origin` config alongside the existing frontend origins.

## 10. Roadmap (phased)

| Phase | Outcome | Time-box |
|---|---|---|
| **0 — Spike** | `yo office` scaffold, sideload to Outlook on the Web + new Outlook for Windows, hello-world taskpane reads the selected message, calls `messages/send` against a local synaplan dev container, prints reply. | 1–2 days |
| **1 — Read-mode MVP** | `SignIn.vue` + dialog OAuth (§7.5) + tiny `/addin/connect` bridge page in `synaplan-website` (or `synaplan/frontend`); Summarise + Ask + Draft-reply + Save-to-RAG. Manifest XML only. Self-hosted load via shared link. | 2–3 weeks |
| **2 — Compose mode + Smart Alerts** | Draft helpers, RAG insert, optional `OnMessageSend` policy guard, optional `OnNewMessageCompose` autoload. | 2 weeks |
| **2.5 — NAA silent SSO** | Add Microsoft Entra as an OIDC provider in Synaplan; add `/api/v1/auth/microsoft/nested` to exchange the NAA token for an API key; add-in tries NAA on first launch, falls back to dialog. | 1 week |
| **3 — Distribution** | Unified manifest variant, store assets, Partner Center submission, validation iterations (4–6 weeks Microsoft turnaround). | 4–8 weeks calendar |
| **4 — Mobile + on-prem Exchange** | Verify against Mailbox 1.5 / Exchange SE, restricted feature set on iOS/Android. | 1 week dev + verification |

(Companion `plugins/synamail/` was considered and **dropped** — the existing Synaplan endpoints are sufficient.)

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
* `useAuth.ts` — dialog open, message-handler dispatch, state-nonce validation, 401-clears-key, sign-out calls revoke.

### 11.3 Component (Vitest + @vue/test-utils)

* `SignIn.vue` — single button, click triggers `displayDialogAsync` with the right URL + `state`; resolves on `messageParent`; rejects on dialog close without message.
* `ReadMode.vue` — buttons disabled until apiKey present; shows loading + error states; renders summary; "Save to RAG" calls upload + process in order.
* `ComposeMode.vue` — `setAsync` / `setSelectedDataAsync` invoked with the right coercion type.
* `Settings.vue` — "Sign out" calls revoke endpoint + clears roaming settings.

### 11.4 E2E (Playwright + Outlook on the Web)

We'll use the [office-addin-test-helpers](https://github.com/OfficeDev/office-js-test-helpers) approach + Playwright headless.

| Scenario | Steps | Pass criteria |
|---|---|---|
| Sideload | Run `npm run sideload`; assert ribbon button appears in OWA | Manifest accepted, button visible |
| First-run | Open taskpane, no key configured | SignIn view shown |
| Sign-in — Google fast path | Click "Sign in", browser already has Synaplan session, click "Connect" in bridge page | Dialog auto-closes within ~2s, ReadMode renders |
| Sign-in — full Google login | Click "Sign in", complete Google OAuth in dialog, click "Connect" | Same as above; `roamingSettings.apiKey` populated |
| Sign-in — wrong state nonce | Inject a bad `state` in the bridge response | Add-in rejects, surfaces error toast, no key stored |
| Sign-in — user cancels dialog | Close dialog without connecting | No key stored, SignIn view still shown |
| Summarise | Open a known seed email, click Summarise | Summary appears within N seconds, matches snapshot prefix |
| Draft reply | Click Draft reply | Reply compose window opens with synaplan-generated HTML body |
| Compose body insert | Open new message, type intent, click Draft from prompt | `body` value contains generated text |
| Save to RAG | Click Save to RAG, pick group | `/files/upload` + `/files/{id}/process` invoked, success toast |
| Ask thread | Send three follow-ups | All hit the same `chatId`, history visible |
| Settings persistence | Reload Outlook | Key still present, no re-prompt |
| Sign-out | Settings → Sign out | `/api/v1/apikeys/{id}` DELETE called, roaming cleared, SignIn view shown |
| 401 from API | Manually revoke key from Synaplan UI, trigger any action | Key cleared, SignIn view shown |
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

### 11.6 Synaplan-side regression (when we touch the bridge page)

The only Synaplan-repo change in v1 is the `/addin/connect` bridge page. Whatever repo it lives in (`synaplan-website` for the marketing site, or `synaplan/frontend/` for the in-app version), the Synaplan house rule from `synaplan/AGENTS.md` still applies:

```bash
make lint && make -C backend phpstan && make test \
  && docker compose exec -T frontend npm run check:types
```

Non-negotiable before any commit.

## 12. Publishing — Microsoft Marketplace (AppSource)

Reference: [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource), [Publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).

### 12.1 Pre-requisites

1. **Partner Center account** in the "Microsoft 365 and Copilot" program (same account works for AppSource and Teams Store). One-off setup, ~1 week verification.
2. **Privacy policy + Terms of Service** URLs hosted publicly without auth — go in `synaplan-website` repo.
3. **Support URL** (likewise public).
4. Stable HTTPS host for the taskpane — `https://addin.synaplan.com` (or under `web.synaplan.com/addin/`).
5. **AppSource icons** in all required sizes (16, 32, 64, 80, 128 PNG) plus a 256/512 hero image.
6. Demo / screencast video covering the add-in in action.
7. Demo Microsoft 365 tenant for the certification testers, pre-loaded with a Synaplan account so reviewers can complete the sign-in flow end-to-end.

### 12.2 Manifest requirements (XML for legacy store, unified for new store)

* `Id` — generate a stable GUID (one per environment: dev / staging / prod).
* `Version` — `a.b.c.d` four-part version, must match the form submitted in Partner Center.
* `DisplayName` — `Synamail` (exact casing).
* `SupportUrl`, `IconUrl`, `HighResolutionIconUrl`, `SourceLocation` — all HTTPS.
* `Permissions` — least privilege (`ReadWriteItem`).
* `Hosts` — only `Mailbox`.
* `Requirements` — `Mailbox` minVersion 1.8.
* `AppDomains` — list every external domain the taskpane will load, including the Synaplan API host **and** the Synaplan login origin (the dialog redirects there).

### 12.3 Submission checklist

* [ ] Manifest passes `office-addin-validator` with zero warnings.
* [ ] Add-in works without errors in Outlook on the Web, new Outlook on Windows, classic Outlook 2024, Outlook on Mac (new + classic).
* [ ] Sign-in flow works end-to-end with the demo tenant supplied to reviewers.
* [ ] No telemetry without consent.
* [ ] Privacy policy + Terms + Support URLs all publicly reachable.
* [ ] All UI strings localised (en + de minimum, matching Synaplan).
* [ ] Functions advertised in the listing actually work.
* [ ] Unique add-in ID + matching version number in form and manifest.
* [ ] Brand name "Synamail" pre-validated against Microsoft's name-collision check.
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

## 13. Decisions made & remaining questions

Closed after this round of feedback:

| # | Decision | Result |
|---|---|---|
| 1 | UI library | **Vue 3 + TypeScript** (locked). Fluent design tokens via CSS, no Fluent React dependency. |
| 2 | Auth | **OAuth via Office Dialog API → existing Synaplan login → `/api/v1/apikeys`** in v1. **NAA silent SSO** in phase 2.5 for Microsoft-account users with matching email. Detailed flow in §7.5. |
| 3 | Companion plugin | **Dropped.** Existing endpoints are sufficient. |
| 4 | Brand / store name | **Synamail** (exact casing). Will pre-validate name during AppSource setup. |
| 5 | CORS | **Confirmed** acceptable on the Synaplan API. Taskpane origin will be added to the API's allow-list alongside the existing frontend origins. |
| 6 | Self-hosted Synaplan support | Settings has a "Use a self-hosted instance" link; the user picks the base URL **before** the dialog opens, so the dialog goes to their instance's login. |

Open / to confirm during the spike (none of these block phase 1):

| # | Question |
|---|---|
| A | Where does `/addin/connect` live — `synaplan/frontend/` (in-app) or `synaplan-website` (marketing site)? Probably the latter, since it's mostly a static page that posts back to the API. Check during phase 1. |
| B | Should the issued API key carry a real `expiresAt` (e.g. 1 year) once Synaplan supports it? |
| C | What are the exact scopes we want on the issued key? Likely `messages:*`, `chats:*`, `files:*`, `rag:*`. Match against existing Synaplan scope strings during phase 1 wire-up. |
| D | Phase 2.5: Microsoft Entra "tenant" model — single multi-tenant app registration owned by Synaplan, or per-customer registrations? Almost certainly multi-tenant; confirm with the team before phase 2.5. |

## 14. Concrete next actions (in order)

1. `cd synaMail && npx --package yo --package generator-office -- yo office` → choose **Outlook Add-in**, **TypeScript**, **Custom** (we strip the default React layer and bring our own).
2. Replace the default webpack config with **Vite + Vue 3 + TS**.
3. Add `manifest.xml` with `DisplayName: Synamail`, `Mailbox 1.8`, `DialogApi 1.1`, and a temporary `https://localhost:3000` source location.
4. `npm run sideload` and verify the ribbon button shows in Outlook on the Web with a Microsoft 365 dev tenant.
5. Build `SignIn.vue` + `useAuth.ts` + `dialog/auth-relay.html`.
6. In parallel, add `/addin/connect` bridge page on the Synaplan side (one HTML + a small JS that calls `/api/v1/apikeys` and `Office.context.ui.messageParent`). Follow the Synaplan repo's pre-commit gate (`make lint && phpstan && test`).
7. Wire `useSynaplan` against a local synaplan stack (`docker compose up -d` in `synaplan/`).
8. Implement Summarise → Draft-reply → Save-to-RAG → Ask, in that order.
9. Set up GitHub Actions: lint, type-check, manifest validate, unit tests, build artefact.
10. Begin AppSource asset preparation in parallel (icons, screenshots, copy, privacy policy hosting).

## 15. References (verified Q1 2026)

Outlook add-in foundation:

* [Outlook JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) — current set 1.15, full client matrix.
* [Develop Outlook add-ins for the new Outlook on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook) — confirms web add-ins are the supported model and COM/VSTO is deprecated for new Outlook.
* [Migrate from COM add-ins to web add-ins](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/migrate-com-to-web-addins).
* [Office Add-ins with the unified app manifest for Microsoft 365](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview).
* [Compare the add-in only manifest with the unified manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/json-manifest-overview).
* [Get an Outlook item's attachments from Exchange](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item) — `getAttachmentsAsync`, `getAttachmentContentAsync`, requirement set 1.8.

Auth in add-ins:

* [Authenticate a user with the Office Dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api) — the pattern we use in v1.
* [`Office.context.ui.displayDialogAsync`](https://learn.microsoft.com/en-us/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)).
* [`Office.context.ui.messageParent`](https://learn.microsoft.com/en-us/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)).
* [Nested App Authentication (NAA) for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in) — phase 2.5.

Distribution:

* [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource).
* [Microsoft Marketplace publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).
* [Microsoft Marketplace certification policies (1100 — Office Add-ins)](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies).

Synaplan side:

* `synaplan/docs/EMAIL.md` — existing email-webhook flow (`smart+keyword@`).
* `synaplan/docs/RAG.md` — file upload + processing levels + groups.
* `synaplan/docs/API_PATTERNS.md` — Zod + OpenAPI contract pattern we will follow on the add-in's HTTP client.
* `synaplan/backend/src/Controller/ApiKeyController.php` — `POST /api/v1/apikeys` (issue) and `DELETE /api/v1/apikeys/{id}` (revoke), the only endpoints the bridge page needs.
* `synaplan/backend/src/Controller/GoogleAuthController.php`, `GitHubAuthController.php`, `KeycloakAuthController.php`, `OidcController.php`, `AuthController.php` — the existing login providers the dialog reuses.
