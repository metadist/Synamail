# Synaplan Integration

This is the canonical reference for **what changes on the Synaplan side to support Synamail**, split by which repo, who runs it, and when.

The answer to "what do I install on `synaplan-platform`?" is the short version: **almost nothing.** All real code changes land in `synaplan/` and are picked up by `synaplan-platform` via the standard `docker compose pull` cycle.

## 1. The three Synaplan repos and what each one does

| Repo                                  | What it is                                                                                                                                                | Who deploys                                                                    |
| ------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------ |
| **`/wwwroot/synaplan`**               | The actual product source — Symfony PHP backend + Vue 3 frontend. CI builds the Docker image `ghcr.io/metadist/synaplan:latest`.                          | GitHub Actions CI on push to `main` (image build + push to GHCR).              |
| **`/wwwroot/synaplan-platform`**      | Production deployment config for the live cluster. `docker-compose.yml`, `.env`, plugin bind-mounts, per-node helper scripts (`re-startweb{1,2,3}.sh`).   | A human on each of `web1` / `web2` / `web3`: pulls the new image and restarts. |
| **`/wwwroot/Synamail`** _(this repo)_ | The Outlook add-in. Hosted separately at `https://addin.synaplan.com` (Sprint 4) and sideloaded into Outlook via a manifest. Calls the live Synaplan API. | Separately deployed; nothing about it ships into the Synaplan Docker image.    |

```
                ┌──────────────────────────────────────────┐
                │   /wwwroot/synaplan        (source)       │
                │   PHP + Vue + CI → builds Docker image    │
                └────────────────────┬──────────────────────┘
                                     │ ghcr.io/metadist/synaplan:latest
                                     ▼
                ┌──────────────────────────────────────────┐
                │   /wwwroot/synaplan-platform (deploy)     │
                │   docker compose on web1 / web2 / web3    │
                │   pulls image, mounts .env + plugins/     │
                └────────────────────┬──────────────────────┘
                                     │ https://web.synaplan.com (API + SPA)
                                     ▼
                ┌──────────────────────────────────────────┐
                │   /wwwroot/Synamail            (add-in)   │
                │   Vue 3 + Office.js, sideloaded           │
                │   calls https://web.synaplan.com/api/v1   │
                └──────────────────────────────────────────┘
```

## 2. What the Synamail-supporting Synaplan changes actually are

### 2.1 Sprint 2 — the bridge page (in `synaplan/`, not in `synaplan-platform/`)

Add **one Vue route** to the SPA at `synaplan/frontend/`:

```
synaplan/frontend/src/
├── router/index.ts                  ← add a public route
│      path: '/addin/connect'
│      name: 'addin-connect'
│      component: AddinConnectView
│      meta: { requiresAuth: false, public: true }
│
└── views/AddinConnectView.vue       ← new file
        1. Load Office.js from CDN inside the page.
        2. Read ?state=<nonce>&baseUrl=... from query string.
        3. If not authenticated (auth store), redirect to
              /login?redirect=/addin/connect&state=<nonce>
        4. If authenticated, render "Connect this Outlook to <email>?" + button.
        5. On click: call apiKeysApi.create({
              name: 'Outlook Add-in (' + ua-derived-host + ')',
              scopes: ['messages:*', 'chats:*', 'files:*', 'rag:*'],
           })
        6. On 201, call Office.context.ui.messageParent(JSON.stringify({
              state, apiKey: response.key, keyId: response.id,
              email: currentUser.email, baseUrl: window.location.origin
           }))
```

**This is the only Synaplan-source change for Sprint 2.** Reuses the existing auth store, the existing `apiKeysApi.create()` client, the existing i18n, the existing route guards. No backend code changes. No new endpoint.

Pre-commit gate to pass before merging the PR in `synaplan/`:

```bash
make lint && make -C backend phpstan && make test \
  && docker compose exec -T frontend npm run check:types
```

CI builds the new image automatically. `ghcr.io/metadist/synaplan:latest` then carries the bridge page.

### 2.2 Sprint 3 — feature wire-up (in `synaplan/`, optional)

No structural changes required. Everything Synamail needs (Summarise, Translate, Draft reply, Classify, Ask, Save-to-RAG, Insert-from-RAG, Contact KB via the `contact:<email>` group convention) goes through the **existing** endpoints listed in `docs/FEATURES.md` §6.

The two open questions for Sprint 3 wire-up that **may** turn into a tiny Synaplan PR if the answer is "missing":

1. **RULE integration write path** — does `PromptController` expose a user-scoped writer for `BSELECTION_RULES`? Audit lands in `STEPS.md` Step 3.7. If absent, file a Synaplan-side ticket; Synamail ships `RuleEditor.vue` as read-only in v1 and a write endpoint arrives in v1.1.
2. **Synapse dry-run for non-admins** — `POST /api/v1/admin/synapse/dry-run` is admin-only. The Synamail "Test against current email" button is hidden for non-admins. If you want it for all users, that needs a non-admin alias on Synaplan; otherwise it's a feature gate, not a blocker.

### 2.3 What goes on `synaplan-platform` — the actual deploy steps

On each web node (`web1`, `web2`, `web3`):

```bash
cd /wwwroot/synaplan-platform

# 1. Pull the new image that carries the bridge page (and any later features).
docker compose pull backend worker

# 2. Restart this node's containers. The helper scripts handle per-node env.
bash re-startweb1.sh    # on web1
bash re-startweb2.sh    # on web2
bash re-startweb3.sh    # on web3

# 3. Verify health on this node.
docker compose ps
curl -sf http://localhost/api/health
```

**There is no schema migration to run from `synaplan-platform/`** for the Synamail bridge — the backend image runs its own migrations on cold start (`start_period: 300s` in the compose healthcheck already accounts for that). The bridge page is pure frontend.

**Plugin directory** (`synaplan-platform/plugins/` bind-mounted into the container) is **untouched** for Sprint 2 and Sprint 3. It only comes into play for the optional Sprint 3.5+ work in §2.5.

### 2.4 Acceptance verification on the platform side

After the rolling restart finishes on all three web nodes:

```bash
# Bridge page is publicly reachable (no auth required to GET; the page
# itself bounces to /login when the user is unauthenticated).
curl -sfL -o /dev/null -w '%{http_code}\n' https://web.synaplan.com/addin/connect
# expect: 200

# API key issuance still works end-to-end with a test account.
# (Use the test Synaplan user once it exists.)
curl -sf -H "X-API-Key: <test-key>" https://web.synaplan.com/api/v1/profile
# expect: 200 with the user's email
```

From the Synamail side, on a developer machine:

```bash
cd /wwwroot/Synamail
make dev &        # https://localhost:3000
make sideload     # opens Outlook on the Web with the manifest
# Click "Sign in to Synaplan" → the live https://web.synaplan.com/addin/connect
# page should now load inside the dialog.
```

### 2.5 SHIPPED (2026-06-10) — the `synamail` Synaplan plugin (Contact AI Profiling)

The plugin anticipated here now exists and carries **Contact AI Profiling**:
rolling per-contact profiles (summary / tone / facts / open loops) whose AI
prompt and storage live server-side, so profiling can evolve without an add-in
release. Same dev→release flow as Synaform:

- **Developed in this repo** under `synamail-plugin/`:

```
synamail-plugin/
├── manifest.json              (id synamail, namespace Plugin\Synamail)
├── backend/
│   └── Controller/
│       └── SynamailController.php   (profile endpoints + the rolling-summary prompt)
├── frontend/
│   ├── index.js               (Synaplan web panel: list + delete profiles)
│   └── i18n/{en,de,es,tr}.json
└── migrations/
    └── 001_setup.sql          (per-user BCONFIG group P_synamail)
```

- **Released to the main repo** with `make sync-plugin` (copies to
  `synaplan/plugins/synamail/`; `make sync-plugin-and-clear` also clears the
  Symfony cache). For production, copy the same directory to
  `synaplan-platform/plugins/synamail/`.

Routes namespace: `/api/v1/user/{userId}/plugins/synamail/profiles...`
(API-key authenticated; the add-in resolves the numeric user id via
`GET /api/v1/auth/me`). Profiles are stored in Synaplan's generic
`plugin_data` table — **no core schema change**.

After dropping the plugin in:

```bash
docker compose restart backend worker
# Then per-user install via:
docker compose exec backend php bin/console app:plugin:install <userId> synamail
```

**Graceful degradation:** on instances without the plugin, the add-in's
profile card shows an install hint (`PROFILING_UNAVAILABLE`); every other
feature keeps working. Remember to install the plugin for the AppSource
reviewer demo account.

## 3. CORS — already permissive

`synaplan/backend/config/packages/nelmio_cors.yaml`:

```yaml
nelmio_cors:
  defaults:
    allow_origin: ['*']
    allow_methods: ['GET', 'OPTIONS', 'POST', 'PUT', 'PATCH', 'DELETE']
    allow_headers: [..., 'X-API-Key']
```

So **any origin** can call `web.synaplan.com/api/v1/*` with the `X-API-Key` header. No platform-side CORS change is required for:

- Sideload dev origin: `https://localhost:3000`
- Future prod origin: `https://addin.synaplan.com`
- Self-hosted instances pointed at a customer's own Synaplan

If at some point Synaplan tightens CORS to an explicit allow-list, **then** the Synamail origins need to be added. Until then, no action.

## 4. What CHANGES vs CONFIGURATION you actually own on `synaplan-platform`

Almost nothing in v1. For completeness:

| `synaplan-platform/` artefact | Touched for Synamail?                                                                    |
| ----------------------------- | ---------------------------------------------------------------------------------------- |
| `docker-compose.yml`          | No.                                                                                      |
| `.env`                        | No — Synamail introduces no new env vars on the Synaplan side in v1.                     |
| `plugins/`                    | **Yes** — copy `synamail-plugin/` → `plugins/synamail/` for Contact AI Profiling (§2.5). |
| `up/` (NFS uploads)           | No — file-upload path is unchanged.                                                      |
| `re-startweb{1,2,3}.sh`       | No — just **run** them after `docker compose pull`.                                      |
| `synaplan.sql`                | No — no schema changes for Synamail.                                                     |
| Galera DB                     | No — no schema changes.                                                                  |
| Reverse proxy / TLS           | No — `web.synaplan.com` already terminates TLS and routes to the image.                  |

## 5. Cheat sheet — your operator actions per sprint

| Sprint | What you do on `synaplan-platform`                                                                                                                                   | What you do on `synaplan` (source)                                                                |
| ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| 1      | Nothing.                                                                                                                                                             | Nothing — planning only.                                                                          |
| 2      | After the bridge-page PR merges and CI republishes the image: `docker compose pull && re-startweb{1,2,3}.sh`. Verify `https://web.synaplan.com/addin/connect` → 200. | Merge the bridge-page PR (`frontend/src/router/index.ts` + `views/AddinConnectView.vue`).         |
| 3      | Same loop if any Sprint-3 Synaplan-side change ships (none expected for v1 unless the RULE audit forces it).                                                         | Optional: ship the `BSELECTION_RULES` user-scoped write endpoint if the audit shows it's missing. |
| 3.5    | Optional: drop `plugins/synamail/` into `synaplan-platform/plugins/`, restart backend + worker, run `app:plugin:install`.                                            | Optional plugin source.                                                                           |
| 4      | Same image-pull loop if any final polish ships.                                                                                                                      | Bump version if the API surface changes during certification iteration.                           |

## 6. References

- Image: `ghcr.io/metadist/synaplan:latest` (built by CI in `synaplan/`).
- Helper scripts: `synaplan-platform/re-startweb{1,2,3}.sh`, `synaplan-platform/startweb{1,2,3}.sh`.
- Plugin precedent: `synaplan-sortx/sortx-plugin/` (the source-of-truth pattern for synaplan plugins).
- Synaplan pre-commit gate (mandatory for the bridge-page PR): `make lint && make -C backend phpstan && make test && docker compose exec -T frontend npm run check:types`.

## 7. Local sign-in loop against a dev Synaplan (no Office.com detour)

This is the **fast local loop**: develop Synamail in WSL, run Synaplan in Docker, and test in your Windows Outlook — no Microsoft 365 dev tenant, no AppSource, no `web.synaplan.com` round-trip.

The Synamail sign-in dialog goes through `Office.context.ui.displayDialogAsync`, which Office hard-rejects unless the URL is HTTPS. The taskpane is itself an HTTPS origin (`https://localhost:3000`), so its `fetch()` calls to Synaplan are also blocked as mixed content unless Synaplan answers over HTTPS. Your local Synaplan exposes the frontend on plain `http://localhost:5173` (which proxies `/api` → backend `:8000`). `make bridge` puts a one-process HTTPS terminator in front of it on `:5174` — no source changes to `synaplan/`, no Docker changes.

The bridge **reuses the office-addin-dev-certs certificate** that already fronts Synamail on `:3000`, instead of a throwaway self-signed one. That matters: a **single** trusted root in Windows then covers both origins, so **desktop Outlook's WebView2 accepts the dialog too** — not just Outlook on the Web.

### 7.1 One-time setup

1. **Dev cert** (provisions `~/.office-addin-dev-certs/`, shared by `:3000` and the `:5174` bridge):

   ```bash
   cd /wwwroot/Synamail && npx office-addin-dev-certs install
   ```

2. **Trust the CA in Windows** so desktop Outlook's WebView2 stops warning (skip if you only use Outlook on the Web and are happy clicking through the warning once):

   ```bash
   cp ~/.office-addin-dev-certs/ca.crt /mnt/c/Users/<you>/Downloads/
   ```

   In Windows: double-click `ca.crt` → **Install Certificate** → **Local Machine** → **Place all certificates in the following store** → **Trusted Root Certification Authorities** → Finish.

3. **Pick the server at sign-in.** There is no mock mode and no env toggle. On the SignIn screen set the instance URL to `https://localhost:5174` (the HTTPS bridge in front of your Docker Synaplan) and sign in — the dialog opens the real `/addin/connect` bridge and stores a real API key. To talk to production instead, set the URL to `https://web.synaplan.com`. Switch any time via **Settings → Log out**.

The taskpane's outbound connections are restricted by `<AppDomains>` in `manifest.xml`; `https://localhost:5174` is already listed. A different port needs a new `<AppDomain>` entry and a re-`make sideload`.

### 7.2 The loop

```bash
# Terminal 1 — Synaplan Docker (frontend :5173, backend :8000)
cd /wwwroot/synaplan && docker compose up -d

# Terminal 2 — HTTPS bridge in front of the Synaplan frontend (:5174 → :5173)
cd /wwwroot/Synamail && make bridge

# Terminal 3 — Synamail taskpane
cd /wwwroot/Synamail && make dev        # https://localhost:3000
```

Then sideload into Outlook. Two no-tenant routes:

- **Outlook on the Web / new Outlook for Windows** — `make sideload` (registers `manifest.xml`). Use a personal `outlook.live.com` account, which allows user sideloading without admin policy.
- **Classic Outlook for Windows desktop** — `make sync` copies the manifest into `C:\addin-catalog\`; add that folder as a Trusted Add-in Catalog once (see `INSTALL.md` Route C). Bypasses tenant policy entirely.

In the Synamail panel: open a message → **"Use a self-hosted instance"** → enter `https://localhost:5174` → **Sign in to Synaplan**. The dialog loads `/addin/connect` from your local `synaplan/frontend/`, which (after you log into local Synaplan) calls `POST /api/v1/apikeys` and posts `{ state, apiKey, keyId, email, baseUrl }` back via both `window.opener.postMessage` and `Office.context.ui.messageParent`. Every subsequent feature call then hits `https://localhost:5174/api/v1/*` → Vite → backend.

### 7.3 Switching back to live before deploying

Synamail keeps the override only for the next sign-in (it isn't persisted). To verify against `web.synaplan.com` again:

1. Sign out from the Synamail taskpane (resets `Office.context.roamingSettings`).
2. On the next sign-in, leave the "Self-hosted instance" field at its default (`https://web.synaplan.com`) — that's what `defaultBaseUrl()` (`src/taskpane/composables/useAuth.ts`) returns when no settings are present.

Nothing on the production / `synaplan-platform` side changes.

### 7.4 Caveats

- **Desktop Outlook needs the trusted CA** (step 7.1.2). With it imported, new Outlook for Windows and classic Outlook 2024 accept the `:5174` dialog. Without it, you're limited to Outlook on the Web (accept the cert warning once per browser profile at `https://localhost:5174/addin/connect`).
- **Cert lifetime.** `office-addin-dev-certs` certs are valid ~30 days; re-run the install command when the bridge or `:3000` starts failing TLS.
- **AppSource submissions.** Strip the `<AppDomain>https://localhost:5174</AppDomain>` line from `manifest.xml` before the certification build (`docs/PROJECT_PLAN.md` calls this out). It's harmless in dev sideloads, but Microsoft's validator flags localhost domains in store submissions.
