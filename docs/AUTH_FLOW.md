# Synamail Authentication / Sign-in Flow — Authoritative Reference

This document describes the **complete, working** sign-in / sign-out flow for the
Synamail Outlook add-in, across **local dev** and **live** Synaplan servers. It
exists because this flow is subtle, spans two repositories, and has regressed
more than once. **Read this before touching anything in the auth path.**

> If you change any file listed under "Files in the flow", re-verify every
> numbered step below and update this document in the same change.

---

## TL;DR (the mental model)

- **There is no mock mode.** Sign-in is _always_ a real round-trip to whatever
  **Synaplan server URL** the user picks on the SignIn screen.
- The signed-in **email** identifies the environment (local `admin@synaplan.com`
  vs your live account). Switching environments = **Settings → Log out**
  (full reset), then sign in to a different server URL.
- The dialog hands the API key back to the taskpane through a **same-origin
  relay**, reached via a **`redirect` query param**. If that param is lost
  anywhere in the chain, the dialog reports "connected" but **never closes**.

---

## The two servers

| Environment | Server URL (set on SignIn) | What it is                                                                     |
| ----------- | -------------------------- | ------------------------------------------------------------------------------ |
| Local dev   | `https://localhost:5174`   | HTTPS bridge (`local-ssl-proxy`) → Synaplan frontend `:5173` → backend `:8000` |
| Live        | `https://web.synaplan.com` | Production Synaplan                                                            |
| Self-hosted | `https://<your-host>`      | Any HTTPS Synaplan instance                                                    |

The base URL **must be HTTPS** — `Office.context.ui.displayDialogAsync` hard-rejects
non-HTTPS. That is why local dev uses the `:5174` bridge, not `http://localhost:8000`.

---

## Files in the flow

**Synamail (this repo):**

- `src/taskpane/composables/useAuth.ts` — builds the dialog URL, opens the
  Office dialog, validates the state nonce, persists the payload.
- `src/dialog/auth-relay.html` + `src/dialog/auth-relay.ts` — the **same-origin
  relay**. Loads Office.js and calls `messageParent` with the payload.
- `src/taskpane/composables/useRoamingSettings.ts` — stores `{ apiKey, keyId,
email, baseUrl }`; `clearAllSettings()` is the "Log out" (full reset) wipe.
- `src/taskpane/views/SignIn.vue` — server URL field + Sign in button.
- `src/taskpane/views/Settings.vue` — Sign out + **Log out** (full reset).

**Synaplan (`/wwwroot/synaplan/frontend/`):**

- `src/views/PlatformConnectView.vue` — the `/connect/platform` **bridge page**
  served by Synaplan. Outlook uses `client=outlook` (not flag-gated). Calls
  `POST /api/v1/addin/connect` and follows the relay redirect the server
  returns.
- `backend/src/Service/PlatformLink/OutlookConnectService.php` — mints the
  add-in key and builds the relay URL **on the server**; the relay host must
  prefix-match the `outlook-builtin` row in `BPLATFORMINSTANCES`.
- `/addin/connect` is a **redirect** to `/connect/platform?client=outlook&…`
  so Synamail's `buildDialogUrl` stays unchanged. Every query param survives.
- `src/views/LoginView.vue` + `src/utils/pendingAuthRedirect.ts` — the login
  round-trip that **must preserve the `redirect` param**.

---

## The handshake, step by step

1. **SignIn (taskpane).** User picks a server URL and clicks "Sign in to
   Synaplan". `useAuth.signIn({ baseUrl })` → `openSignInDialog` →
   `buildDialogUrl(baseUrl, state)` produces:

   ```
   <baseUrl>/addin/connect?state=<nonce>&label=Outlook+Add-in&redirect=<relayUrl>
   ```

   - `state` is a random nonce, validated on return (anti-CSRF).
   - `redirect` = `relayUrl()` = `<taskpane-origin>/src/dialog/auth-relay.html`
     (e.g. `https://localhost:3000/src/dialog/auth-relay.html`).

2. **Dialog opens** the bridge via `displayDialogAsync(url, { displayInIframe:
false })`. The URL is still `/addin/connect?…`; Synaplan **redirects** it to
   `/connect/platform?client=outlook&…` (same query, plus `client=outlook`).
   The taskpane listens on **two** channels: Office
   `DialogMessageReceived` _and_ a `window` `message` event. The bridge only
   ever delivers through the relay (step 5) or Office `messageParent`; it no
   longer posts to `window.opener`, so the `window` listener is a
   compatibility no-op kept for older bridges.

3. **Bridge bootstrap (`PlatformConnectView.vue`).** Waits for `authReady`.
   - **If NOT authenticated:** it must redirect to
     `/login?redirect=<full original query>` — **including the relay `redirect`
     param**. It uses `route.fullPath` for exactly this reason (see "The #1
     regression" below). The user logs in (e.g. `admin@synaplan.com`).
   - **If authenticated:** shows "Connect Outlook at `<host>` as `<email>`?" +
     button.

4. **Login round-trip (`LoginView.vue`).** After login it navigates to
   `route.query.redirect` (validated by `isSafeRedirectPath`) — restoring
   `/connect/platform?client=outlook&…&redirect=<relayUrl>` intact. The relay
   param **survives** only because step 3 used `route.fullPath`.

5. **Connect (`handleConnect`).** The bridge calls
   `POST /api/v1/addin/connect { state, redirect_uri, base_url }`. Synaplan
   mints the scoped add-in key, checks `redirect_uri` against the server-side
   allow-list (`outlook-builtin` in `BPLATFORMINSTANCES`: `https://localhost`
   and `https://127.0.0.1` on any port, `https://addin.synaplan.com`,
   `https://*.synaplan.com`; HTTPS only) and answers
   `{ redirect, payload }`. When the relay is allow-listed, `redirect` is

   ```
   <redirect_uri>#payload=<base64(JSON payload)>
   ```

   and the bridge does `window.location.assign(redirect)`. When it is not,
   `redirect` is `null` and the bridge falls back to Office `messageParent`
   with `payload` (the rejection is audited as `platform_link.redirect_rejected`).

   The dialog now navigates to the relay, **same-origin as the taskpane**.

6. **Relay (`auth-relay.ts`).** On `Office.onReady`, reads `#payload`, base64-decodes
   it, and calls `Office.context.ui.messageParent(json)`. This works because the
   relay is same-origin as the taskpane **and loads Office.js** (`<script
src=".../office.js">` in `auth-relay.html`).

7. **Taskpane receives** the payload, validates `state` matches the nonce from
   step 1, requires `apiKey` + `email` + `baseUrl`, then `saveSettings(...)`,
   `hydrateAuthState()`, closes the dialog, and routes to Home.

8. **Sign out (`Settings.vue` → `signOut`).** Revokes the key
   (`DELETE /api/v1/apikeys/{id}`, best-effort) and `clearSettings()`.
   **Log out** (full reset, `clearAllSettings`) additionally wipes the preferred
   base URL and any `synamail.*` localStorage keys — use this to switch
   environments cleanly (local → live).

---

## The #1 regression — "it says connected but the window stays open"

**Symptom:** after clicking **Connect**, the bridge shows success but the Office
dialog never closes and the taskpane never signs in.

**Root cause:** the `redirect` (relay) param was lost somewhere, so step 5 fell
back to a **cross-origin `messageParent`** from the bridge — which **Outlook
desktop silently drops** after the login navigations. The fallback only works
reliably in Outlook on the Web.

**The specific bug we keep reintroducing:** `PlatformConnectView.bootstrap()`
rebuilds the login-redirect URL from a _subset_ of params (e.g. only `state` +
`baseUrl`), dropping `redirect`. After login the bridge has no relay target.

**The fix (and invariant):** the login round-trip MUST preserve the **entire**
original query. Use `route.fullPath`:

```ts
// PlatformConnectView.bootstrap(), unauthenticated branch
const redirect = route.fullPath // keeps state + label + baseUrl + redirect
setPendingRedirect(redirect)
void router.push({ path: '/login', query: { redirect } })
```

This is regression-proof against future param additions. Do **not** reconstruct
the redirect URL field-by-field.

---

## Invariants — do not break these

1. **Every manifest/dialog URL is HTTPS.** Local dev therefore uses the `:5174`
   bridge, never `http://localhost:8000`.
2. **The relay (`auth-relay.html`) must (a) be same-origin as the taskpane and
   (b) load Office.js.** Without Office.js, `messageParent` is undefined and the
   dialog can't close.
3. **`buildDialogUrl` must send `redirect=relayUrl()`.** Desktop Outlook needs
   the relay path; the postMessage fallback alone is not enough.
4. **The login round-trip must preserve the `redirect` param** (see above).
5. **The `state` nonce must round-trip unchanged** and is re-validated in the
   taskpane. A mismatch is a _correct_ rejection, not a bug to "fix" by removing
   the check.
6. **The relay allow-list lives on the server**, in the `outlook-builtin` row
   of `BPLATFORMINSTANCES` (seeded by Synaplan migration
   `Version20260908120000`, matched by `RedirectUriPolicy`). The bridge never
   decides where the key goes and never posts it to `window.opener`. Adding a
   relay host is a Synaplan data change, not a frontend edit — keep the list
   tight, it guards against leaking the API key to an attacker-controlled
   origin.
7. **No mock mode.** Don't reintroduce a mock relay, `mock-key-` auto-selection,
   or a `MockSynaplanClient`. Sign-in is always real.

---

## Local dev: bringing the flow up

```bash
cd /wwwroot/Synamail
./start-dev.sh        # Synaplan Docker + taskpane :3000 + HTTPS bridge :5174
```

This is idempotent. It must leave **three** things listening:

| Port | Process           | Check                                                                             |
| ---- | ----------------- | --------------------------------------------------------------------------------- |
| 3000 | Vite taskpane     | `curl -ksI https://localhost:3000/src/taskpane/taskpane.html` → 200               |
| 5174 | `local-ssl-proxy` | `curl -kso/dev/null -w '%{http_code}' https://localhost:5174/addin/connect` → 200 |
| 5173 | Synaplan frontend | (Docker) `curl -s -o/dev/null -w '%{http_code}' http://localhost:5173/` → 200     |

The bridge `:5174` is a **separate process that does NOT survive a reboot**. If
`https://localhost:5174` errors, it's almost always that the bridge isn't
running — start it with `./scripts/dev-bridge-proxy.sh` (or `./start-dev.sh`).

The bridge reuses the `office-addin-dev-certs` certificate, so a single trusted
root in Windows/macOS covers both `:3000` and `:5174`.

---

## Switching local ↔ live (the supported workflow)

1. Sign in locally: SignIn → server URL `https://localhost:5174` → sign in as
   `admin@synaplan.com`.
2. To switch: **Settings → Log out** (full reset — wipes key + email + baseUrl).
3. Sign in live: SignIn → server URL `https://web.synaplan.com` → sign in with
   your real account.

Both directions must work. The environment is identified by the signed-in email,
never by a build flag.

> Note: a fix to `PlatformConnectView.vue` only takes effect on a given server once
> that server serves the new code. **Local** picks it up immediately (the
> `synaplan-frontend` container hot-reloads `./frontend`). **Live** requires the
> Synaplan image to be rebuilt and `synaplan-platform` to pull it — see
> `docs/SYNAPLAN_INTEGRATION.md`.

---

## Quick triage

| Symptom                                                      | Most likely cause                                                             | Fix                                                                         |
| ------------------------------------------------------------ | ----------------------------------------------------------------------------- | --------------------------------------------------------------------------- |
| `https://localhost:5174` errors / refused                    | Bridge process not running (doesn't survive reboot)                           | `./start-dev.sh` (or `./scripts/dev-bridge-proxy.sh`)                       |
| "Connected" but dialog stays open                            | `redirect` relay param lost → cross-origin `messageParent` dropped by desktop | Ensure `PlatformConnectView.bootstrap` uses `route.fullPath` (above)        |
| Dialog lands on `/connect/platform` without `client=outlook` | `/addin/connect` redirect dropped `client` or a bookmark skipped it           | The legacy path must redirect with `client=outlook` plus the original query |
| Dialog opens then immediately closes with an error           | `state` nonce mismatch (this is correct for a malformed payload)              | Retry; capture the payload if reproducible                                  |
| Cert warning in the dialog                                   | Dev CA not trusted on the OS                                                  | Trust `~/.office-addin-dev-certs/ca.crt`; re-run dev-certs if expired       |
| Stuck signed in as the wrong user/server                     | Stale roaming settings                                                        | **Settings → Log out** (full reset), then sign in again                     |

---

## Related

- `docs/SYNAPLAN_INTEGRATION.md` — cross-repo map (bridge page, deploy path).
- `INSTALL.md` — full local dev setup, sideload, dev certs.
- `AGENTS.md` — links here under "Authentication / Sign-in flow".
