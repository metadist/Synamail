# Outlook E2E Automation

How we drive the **real** Synamail taskpane under test, and how that maps onto
the real Outlook clients — including the Windows desktop client from WSL.

A Synamail taskpane only ever runs inside a web runtime: an **iframe** (Outlook
on the Web) or **Edge WebView2** (desktop — classic Outlook 2024 and new Outlook
for Windows). "Automating Outlook" therefore means getting a Chrome DevTools
Protocol (CDP) handle on that runtime and driving it with Playwright. There is no
COM/AppleScript shortcut for the add-in surface.

We use a layered strategy. Build and rely on the lower layers first — they are
faster, deterministic, and catch the vast majority of regressions.

| Layer                                               | What it proves                                      | Outlook needed? | Where it runs       | Status                 |
| --------------------------------------------------- | --------------------------------------------------- | --------------- | ------------------- | ---------------------- |
| **L1** — Vitest + Office stub (`tests/setup.ts`)    | UI/logic units                                      | No              | CI + local          | shipped                |
| **L2** — Playwright + Office shim → mocked Synaplan | The real Vue taskpane + real client code, all flows | No              | CI + local          | **shipped (this doc)** |
| **L2-live** — same specs → real Synaplan            | Real API round-trips end-to-end                     | No              | local / opt-in      | shipped (env toggle)   |
| **L3a** — Playwright → Outlook on the Web           | Real Office iframe runtime                          | OWA only        | local / nightly     | runbook below          |
| **L3b** — Playwright CDP → desktop Windows Outlook  | Real WebView2 desktop runtime                       | Yes, Windows    | local / pre-release | runbook below          |

## L2 — the everyday harness (`tests/e2e/`)

L2 loads the **real built taskpane** in Playwright Chromium and feeds it a
faithful, deterministic `Office` global via `page.addInitScript` (installed before
any app script, exactly as a host would). It pre-seeds roaming settings with an
API key so the app boots straight into the signed-in Home view, then drives the
real UI. The Synaplan API is stubbed at the network boundary with the exact wire
shapes verified against a live local stack (see `docs/PROJECT_PLAN.md`), so the
app's real client code path runs unchanged — only `fetch` responses are canned.

Files:

- `tests/e2e/support/office-shim.ts` — the injectable `Office` shim + fake mail
  items (`READ_ITEM`, `COMPOSE_ITEM`). Records every Office mutation the add-in
  makes (`displayReplyForm`, `body.setAsync`, `body.setSelectedDataAsync`,
  `displayNewMessageFormAsync`, …) on `window.__officeCalls` for assertions.
- `tests/e2e/support/synaplan-mock.ts` — `page.route('**/api/v1/**', …)` answering
  every endpoint the client calls (`/auth/me`, `/messages/send`, `/chats`,
  `/files/groups`, `/files/upload`, `/rag/search`, `/config/models*`). The AI
  reply for `/messages/send` is chosen by the system prompt, so each action gets
  a sensible, assertable answer.
- `tests/e2e/support/harness.ts` — `bootTaskpane(page, item)` ties it together
  (it also blocks the Office.js CDN so the shim isn't overwritten), plus
  `openReadView` / `openComposeView` / `officeCalls` helpers and the `LIVE` flag.
- `tests/e2e/{read-mode,compose-mode,contact-profile}.spec.ts` — the flow specs.

All specs are tagged `@ci`, so they run in GitHub CI's E2E job (which runs
`--grep '@ci'`) as well as locally. They are deterministic — no real backend, no
ambient env.

Run them:

```bash
make test-e2e                 # whole Playwright suite (mocked)
npx playwright test tests/e2e/read-mode.spec.ts   # one file
npx playwright test --grep '@ci'                  # exactly what CI runs
```

### L2-live — same specs, real Synaplan

Set the env and point at a real instance; the harness skips the network mock and
the specs relax their content assertions to "a result appeared" (exact strings
are mock-only):

```bash
export SYNAPLAN_E2E_LIVE=1
export SYNAPLAN_BASE_URL=https://localhost:5174   # local HTTPS bridge, or web.synaplan.com
export SYNAPLAN_API_KEY=sk_...                     # mint via the dashboard / API
make test-e2e-live
```

## L3a — Outlook on the Web (real Office runtime)

Playwright's own Linux Chromium (in WSL) reaches `outlook.office.com` directly —
no Windows needed. Prerequisites:

1. A dedicated **M365 dev-tenant** test mailbox (never a personal account).
2. The add-in **admin-deployed** into that tenant once (Integrated Apps), so the
   taskpane is present without per-run sideloading.
3. A saved Playwright **`storageState`** (logged-in cookies) captured once
   interactively, to skip the login/MFA dance on every run.

Then a spec navigates OWA, opens an email, opens the Synamail panel, and asserts.
This is the highest-fidelity automated layer that needs no Windows client.

## L3b — the desktop Windows Outlook client, from WSL

This drives the **actual Windows Outlook** taskpane (WebView2) over CDP. It is the
most brittle layer — reserve it for a small pre-release desktop smoke, not the
everyday net.

Environment on this machine (probed 2026-06-04): classic Outlook at
`C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE`, new Outlook
(`Microsoft.OutlookForWindows`) installed, **WebView2 runtime 148.x** present,
WSL in **NAT mode** with the Windows host reachable at the default gateway
(`ip route | awk '/default/{print $3}'`, e.g. `172.28.32.1`).

1. **Sideload once** — `make sync` copies the manifest into `C:\addin-catalog`
   (the trusted catalog classic Outlook reads). No tenant admin required.
2. **Launch Outlook with the WebView2 debug port** from WSL via `powershell.exe`:

   ```bash
   powershell.exe -NoProfile -Command \
     '$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=9222"; \
      Start-Process "C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE"'
   ```

3. **Bridge the loopback port** — WebView2 binds the debug port to `127.0.0.1`,
   which NAT-mode WSL can't reach. Add a one-time port proxy (elevated PowerShell
   on Windows):

   ```powershell
   netsh interface portproxy add v4tov4 listenport=9222 listenaddress=0.0.0.0 connectport=9222 connectaddress=127.0.0.1
   ```

4. **Attach Playwright from WSL** to the host IP and find the taskpane target:

   ```ts
   import { chromium } from '@playwright/test'
   const host = '172.28.32.1' // ip route default gateway
   const browser = await chromium.connectOverCDP(`http://${host}:9222`)
   const ctx = browser.contexts()[0]
   const page = ctx.pages().find((p) => p.url().includes('taskpane.html'))
   ```

Caveats: WebView2 attach timing is racy (open the panel, then poll for the
target); the port proxy must exist; `office-addin-debugging` scripted sideload
needs Node **on Windows** (not installed here) — that's why we use the trusted
catalog instead. Use L3b only to certify the desktop runtime; L2 + L3a cover the
behaviour.

## Why block Office.js in L2?

`src/taskpane/taskpane.html` loads the real Office.js from the Microsoft CDN.
Outside a real host it installs an **empty, item-less** context and would
overwrite our injected shim. The harness routes that CDN request to an empty
script so the shim remains the Office runtime. This only affects the L2 harness;
the shipped HTML is untouched.
