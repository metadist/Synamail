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

### 2.5 Optional — Sprint 3.5+: `plugins/synamail/` on the platform

If we later want server-side hygiene specific to the Outlook add-in (e.g. enforcing the `contact:<email>` RAG-group naming convention, exposing a non-admin RULE-preview endpoint, tagging files with `from`/`to` headers automatically on upload), we add a Synaplan plugin under:

```
synaplan-platform/plugins/synamail/
├── manifest.json
├── backend/
│   ├── Controller/
│   │   └── SynamailController.php
│   └── Service/
└── migrations/
    └── 001_setup.sql          (per-user BCONFIG keys)
```

Routes namespace: `/api/v1/user/{userId}/plugins/synamail/...` (matches the SortX precedent).

After dropping the plugin in:

```bash
docker compose restart backend worker
# Then per-user install via:
docker compose exec backend php bin/console app:plugin:install <userId> synamail
```

**This is opt-in and non-blocking.** Nothing in Sprints 2, 3, or 4 of the Synamail plan depends on this plugin existing.

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

| `synaplan-platform/` artefact | Touched for Synamail?                                                   |
| ----------------------------- | ----------------------------------------------------------------------- |
| `docker-compose.yml`          | No.                                                                     |
| `.env`                        | No — Synamail introduces no new env vars on the Synaplan side in v1.    |
| `plugins/`                    | No in v1. Optional `plugins/synamail/` from Sprint 3.5 onwards (§2.5).  |
| `up/` (NFS uploads)           | No — file-upload path is unchanged.                                     |
| `re-startweb{1,2,3}.sh`       | No — just **run** them after `docker compose pull`.                     |
| `synaplan.sql`                | No — no schema changes for Synamail.                                    |
| Galera DB                     | No — no schema changes.                                                 |
| Reverse proxy / TLS           | No — `web.synaplan.com` already terminates TLS and routes to the image. |

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
