# Synamail — execution steps

> Companion to [`PLAN.md`](./PLAN.md). Each step has a clear "Done when…" criterion and an integration test. Each milestone is independently shippable and reviewable. Keep this file up to date — tick boxes as we go.

## Milestone A — Skeleton (no Synaplan integration yet)

These steps prove the add-in loads and our build/test/CI loop works. **No credentials needed from you.**

### Step 1 — Bootstrap project (½ day)

**Build**

* `cd synaMail && npx --package yo --package generator-office -- yo office` → **Outlook Add-in**, **TypeScript**, name **Synamail**.
* Commit the generated scaffold to a feature branch.
* Replace the default webpack toolchain with **Vite + Vue 3 + TS** (we keep Yeoman's `manifest.xml`, `commands/`, and dev-cert scripts; everything else is rewritten).
* Add `package.json` scripts: `dev`, `build`, `lint`, `check:types`, `test`, `validate`, `sideload`.
* Add `.editorconfig`, `.prettierrc`, `eslint.config.js` matching `synaplan/frontend/` conventions (no semicolons, single quotes, no `any`).
* Pin Office.js types (`@types/office-js`) and `office-addin-validator`, `office-addin-debugging`, `office-addin-dev-certs`.

**Test**

```bash
npm install
npm run build       # Vite production build succeeds
npm run lint        # zero errors
npm run check:types # vue-tsc -b clean
npm run validate    # office-addin-manifest validate manifest.xml → pass
```

**Done when** all four scripts above exit 0 in CI. No functionality yet.

---

### Step 2 — Hello-Synamail taskpane sideloaded into Outlook on the Web (½ day)

**Build**

* `manifest.xml`: `DisplayName="Synamail"`, `Mailbox 1.8` minimum, `DialogApi 1.1`, source location `https://localhost:3000/taskpane.html`, ribbon button (read mode + compose mode), permission `ReadWriteItem`, app domains list with `web.synaplan.com` placeholder.
* `src/taskpane/App.vue` renders **"Synamail"** title + the current message subject (read mode) or "Compose mode" placeholder. Uses `Office.onReady`.
* Generate dev cert (`office-addin-dev-certs install`).
* `npm run sideload` → wraps `office-addin-debugging start manifest.xml`.

**Test (manual, against your own Outlook on the Web)**

1. Run `npm run dev` (Vite serves on `https://localhost:3000`).
2. Run `npm run sideload`.
3. Outlook on the Web opens, ribbon shows the **Synamail** button.
4. Click it on a selected email → taskpane opens, shows the subject.
5. Open a new message (Compose) → ribbon button works → taskpane shows "Compose mode".

**Test (automated)**

* Vitest unit test on `App.vue` with a mocked `Office.context.mailbox.item`.
* CI: lint + types + manifest validate + build.

**Done when** the human steps above pass in OWA Edge + OWA Chrome, and CI is green on the branch.

---

### Step 3 — CI pipeline + protected `main` (½ day)

**Build**

* `.github/workflows/ci.yml`: matrix Node 20 + 22; runs lint, types, validate, unit tests, build, uploads `dist/` artefact.
* Branch protection on `main`: require CI green + 1 review.
* `README.md` with quickstart (clone → `npm install` → `npm run dev` → `npm run sideload`).

**Test**

* Open a PR with an intentional lint error → CI fails. Fix it → CI passes.

**Done when** PRs cannot merge with red CI.

---

## Milestone B — Outlook integration (still no Synaplan, all calls stubbed)

We finish the Office.js side with a fake API client so we can iterate on UX without depending on Synaplan. **Still no credentials needed.**

### Step 4 — Read-mode UI shell + mock client (1 day)

**Build**

* `src/shared/synaplan-client.ts` with a real fetch wrapper *and* a `MockSynaplanClient` selectable via env var.
* `src/taskpane/views/ReadMode.vue` with five disabled buttons: Summarise, Draft reply, Classify, Save to knowledge base, Ask.
* `src/taskpane/composables/useOutlookItem.ts` exposes `subject`, `from`, `to`, `bodyText` (lazy-loaded via `body.getAsync('text')`), `attachments`.
* Mock client returns canned strings for each action with a 500ms simulated delay.

**Test**

* Vitest covers `useOutlookItem` against a mocked Office context.
* Manual: Click Summarise → fake summary appears. Click Save to RAG → fake "saved" toast.

**Done when** all five read-mode buttons round-trip through the mock client and the spinner/error states render correctly.

---

### Step 5 — Compose-mode UI shell + mock client (1 day)

**Build**

* `src/taskpane/views/ComposeMode.vue` with: Draft from prompt, Improve selection, Translate selection.
* `Draft from prompt` → calls mock client, then `body.setAsync(html, { coercionType: 'html' })`.
* `Improve/Translate` → reads via `body.getSelectedDataAsync`, replaces via `body.setSelectedDataAsync`.

**Test**

* Vitest unit test on the composable that wraps `setAsync` / `setSelectedDataAsync`.
* Manual: open new mail, click Draft from prompt → body is filled. Select text, click Translate → selection replaced.

**Done when** all three compose buttons round-trip through the mock client.

---

### Step 6 — SignIn view with mock dialog (½ day)

**Build**

* `src/taskpane/views/SignIn.vue`: single "Sign in to Synaplan" button + small "Use a self-hosted instance →" link.
* `src/taskpane/composables/useAuth.ts`: opens `Office.context.ui.displayDialogAsync(authRelayUrl + '?state=...')`, handles `DialogMessageReceived` and `DialogEventReceived`.
* `src/dialog/auth-relay.html` (mock): static page that immediately calls `Office.context.ui.messageParent(JSON.stringify({ apiKey: 'fake-test-key', keyId: 1, email: 'demo@synaplan.test', baseUrl: 'https://web.synaplan.com', state }))` so the loop closes without leaving localhost.
* `useRoamingSettings.ts`: persists `{ apiKey, keyId, email, baseUrl }`.
* App-level routing: no key → SignIn; has key + read item → ReadMode; has key + compose item → ComposeMode.

**Test (manual)**

1. Open taskpane fresh → SignIn view.
2. Click "Sign in" → mock dialog opens, immediately closes.
3. Taskpane re-renders to ReadMode/ComposeMode.
4. Reload Outlook → state survives (roaming settings work).

**Test (automated)**

* Vitest mocks `Office.context.ui.displayDialogAsync` and asserts the state nonce, message dispatch, and storage write.

**Done when** the full sign-in dance works against the local fake bridge, including state-nonce rejection.

---

### Step 7 — Settings + Sign-out (½ day)

**Build**

* `Settings.vue`: shows email + base URL, "Sign out" button.
* Sign-out clears roaming settings (and in step 12 will call `DELETE /api/v1/apikeys/{keyId}`).
* In Settings, "Use a self-hosted instance" lets the user override `baseUrl` *before* signing in (so the next dialog opens against their server).

**Test (manual)**

* Sign in → see settings → sign out → SignIn view again.
* Override base URL to `https://example.invalid` → next sign-in dialog opens that URL (will fail, but proves wiring).

**Done when** signing out clears state and the base URL override is respected by the next dialog open.

---

🟢 **Checkpoint A** — at this point we have a fully working Outlook add-in with mocked Synaplan responses. Reviewable, demoable. **Now we connect to the real platform.**

---

## Milestone C — Live integration with web.synaplan.com

**Here is where I need credentials from you — see "Access I need" at the bottom.**

### Step 8 — Add `/addin/connect` bridge page on the Synaplan side (1 day)

This is the **only** Synaplan-repo change in v1.

**Build (in `synaplan-website` or `synaplan/frontend`, to be decided in this step)**

* New static page at `/addin/connect`:
  * Loads Office.js (`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`).
  * If user is not authenticated → redirect to existing `/login` with `?redirect=/addin/connect&state=...`.
  * If authenticated → render "Connect this Outlook to `<email>`?" with a single **Connect** button.
  * On Connect: `POST /api/v1/apikeys` with `{ name: "Outlook Add-in (<UA-derived host>)", scopes: ["messages:*","chats:*","files:*","rag:*"] }`.
  * On 201: `Office.context.ui.messageParent(JSON.stringify({ apiKey, keyId, email, baseUrl, state }))`.
  * On error: render the error and a "Try again" button.
* Whitelist the add-in's dev + prod origins in CORS config (you confirmed CORS is fine — this just makes it explicit in code review).

**Test**

* Sideloaded add-in → click Sign in → bridge page opens at `https://web.synaplan.com/addin/connect` → if you're already logged in to Synaplan in your default browser, it flashes through and the taskpane gets a real key.
* Synaplan's API-keys page (in your browser) lists the new "Outlook Add-in (...)" entry.
* Run the synaplan repo pre-commit gate before merging the bridge page:

  ```bash
  make lint && make -C backend phpstan && make test \
    && docker compose exec -T frontend npm run check:types
  ```

**Done when** a fresh sideload of the add-in, signing in via the live web.synaplan.com login (Google OAuth path included), produces a real API key in `roamingSettings` with no copy-paste.

---

### Step 9 — Replace mock client with real Synaplan client (½ day)

**Build**

* `synaplan-client.ts`: real fetch implementation, `X-API-Key` header, retries with exponential backoff for 5xx, surfaces 401 by clearing roaming settings and bouncing to SignIn.
* Generate Zod schemas in CI from the OpenAPI spec at `https://web.synaplan.com/api/doc.json` and check them in. (Mirror the `make -C frontend generate-schemas` workflow used by Synaplan.)
* Smoke test on app boot: `GET /api/v1/profile` (or equivalent), if 401 → clear key.

**Test (manual)**

* Sign in (real flow). Observe Network tab: `GET` to `/api/v1/profile` returns 200 with the test user.
* Manually revoke the key from Synaplan's API-keys page → next action in the add-in returns 401 → add-in clears state and shows SignIn.

**Done when** mock client is removed and every taskpane action goes through real HTTP to web.synaplan.com.

---

### Step 10 — Read-mode features wired up (1.5 days)

Implement against real endpoints, in this order:

1. **Summarise** → `POST /api/v1/messages/send` with the email body + summarise system prompt.
2. **Ask follow-ups** → `POST /api/v1/chats` (creating one chat per Outlook conversation, key by `mailbox.item.conversationId` in roamingSettings) and `GET /api/v1/chats/{id}/messages`.
3. **Save to knowledge base** → `getAsFileAsync` (1.8) → `POST /api/v1/files/upload` → `POST /api/v1/files/{id}/process` with chosen RAG group (group picker calls `GET /api/v1/files/groups`).
4. **Draft reply** → `messages/send` with reply system prompt → `mailbox.displayReplyForm({ htmlBody, attachments: [] })`.
5. **Classify** (optional in this step) → `messages/send` with classification prompt + JSON output schema.

**Test (manual against test user on web.synaplan.com)**

* Open a known seed email → Summarise → real Synaplan summary appears.
* Send three Ask messages → all attach to the same `chatId`; verify in Synaplan's chat history that they're under one conversation.
* Save to RAG → file appears in the user's RAG group; `GET /api/v1/files` lists it.
* Draft reply → reply window opens with generated body.

**Test (automated)**

* Playwright E2E: full read-mode flow against a dedicated test mailbox (Microsoft 365 dev tenant) and a dedicated test Synaplan user.

**Done when** all four flows work against the live platform with a test user, and the Playwright suite passes.

---

### Step 11 — Compose-mode features wired up (1 day)

* **Draft from prompt** → `messages/send` → `body.setAsync` (HTML).
* **Improve / shorten / translate** → `messages/send` with the relevant system prompt → `body.setSelectedDataAsync`.
* **Insert from RAG** → `POST /api/v1/rag/search` → result list → on click, `setSelectedDataAsync` with the formatted snippet + citation.

**Test (manual)**

* Open a new mail, click Draft from prompt → body filled with a real Synaplan response.
* Select 3–4 lines, click Improve → those lines are replaced.
* Insert from RAG: query "quarterly report" → real RAG hits from the test user's group; click → snippet inserted.

**Done when** all three compose flows work end-to-end against the live platform.

---

🟢 **Checkpoint B** — fully functional Outlook add-in talking to web.synaplan.com with a real account. Demoable to customers.

---

## Milestone D — Polish & ship

### Step 12 — Sign-out revokes the key (½ day)

* Wire `Settings.vue` Sign out to `DELETE /api/v1/apikeys/{keyId}`.
* Manual + automated test that the key disappears from Synaplan's API-keys page after sign-out.

### Step 13 — i18n (en + de) (½ day)

* `vue-i18n`, `en.json`, `de.json` mirroring Synaplan's house rule ("Always update BOTH").
* Locale auto-detected from `Office.context.displayLanguage` with manual override in Settings.
* Tested by switching Outlook language to German and reloading.

### Step 14 — Production hosting + manifest (½ day)

* Deploy the taskpane to `https://addin.synaplan.com` (Cloudflare Pages or `synaplan-platform`).
* Update `manifest.xml` source location and `AppDomains`.
* Strict CSP on the host.
* Confirm CORS allow-list on the API includes the production origin.

### Step 15 — Smart Alerts + event-based activation (1 week, optional, phase 2 of the original plan)

* `OnNewMessageCompose` event auto-suggests reply (gated by Settings toggle).
* `OnMessageSend` Smart Alerts integration for outbound governance.

### Step 16 — Unified manifest variant (½ day)

* Run `office-addin-manifest-converter` on `manifest.xml` → `manifest.unified.json`.
* CI validates both.
* Test sideloading the unified manifest into the new Outlook on Windows.

### Step 17 — AppSource submission prep (4–8 weeks calendar)

* Partner Center account.
* Privacy policy + Terms + Support URLs hosted on `synaplan-website`.
* Icons (16/32/64/80/128) + 256/512 hero + screencast video.
* Demo Microsoft 365 tenant credentials prepared for reviewers.
* Run `office-addin-validator` with zero warnings.
* Submit; iterate (typically 2 rounds).

---

## Cross-cutting test plan (runs continuously from Step 1)

Every step adds tests; nothing replaces what came before. The full pyramid:

| Layer | Tooling | Runs on |
|---|---|---|
| Static | ESLint + Prettier + `vue-tsc -b` + `office-addin-manifest validate` | every commit |
| Unit | Vitest | every commit |
| Component | Vitest + `@vue/test-utils` with mocked Office.context | every commit |
| E2E (OWA) | Playwright headless + sideload | nightly + before release |
| Manual smoke | new Outlook for Windows, classic Outlook 2024, Outlook on Mac | before release |
| Mobile smoke | Outlook iOS / Android (1.5 cap) | before release in phase 3 |

---

## What I need from you, and when

| When | What | Why |
|---|---|---|
| **Step 1–7** (skeleton + mocks) | nothing | I can build and test entirely locally |
| **Step 8** (bridge page) | A test Synaplan user on `web.synaplan.com` (email + temporary password, or pointed at a Google account I can sign into). Plus confirmation of where the bridge page should live: `synaplan-website` (marketing) or `synaplan/frontend` (in-app). | To verify the OAuth dialog flow end-to-end against the live login + `/api/v1/apikeys` |
| **Step 8** | Confirmation that `https://localhost:3000` (sideload dev origin) is in the API CORS allow-list. If not, please add it; I'll send you the exact origin string when I get there. | Without this the add-in will get blocked CORS preflights when calling the live API from the dev manifest |
| **Step 9–11** (real integration) | Same test user as above. Ideally with at least one RAG group seeded with a few documents so I can test "Save to RAG" listing and "Insert from RAG" search. | To exercise the full feature surface |
| **Step 14** (prod hosting) | Decision on production host: `addin.synaplan.com` (new) vs `web.synaplan.com/addin/`. DNS + cert configuration if it's a new hostname. | The manifest's `SourceLocation` must match the production HTTPS origin |
| **Step 17** (AppSource) | Partner Center account access, plus a Microsoft 365 dev tenant for reviewers. I can self-provision a Microsoft 365 dev tenant via the free Microsoft 365 Developer Program for *my own* testing through Step 16. | Required by Microsoft for store submission |

**Practical request right now:** for Step 8 onwards, please create a dedicated Synaplan test user (something like `synamail-dev@<yourdomain>`) and DM me the credentials. A regular email/password account is enough — I don't need it linked to Google specifically, since the dialog flow will exercise whichever providers Synaplan offers. If you'd like me to specifically test the **Google** sign-in path end-to-end, I'll also need a Google account that's permitted to sign into that Synaplan user.

I will not need any of this for Steps 1 through 7.
