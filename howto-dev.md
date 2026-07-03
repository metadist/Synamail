# How-to: run Synamail locally and sideload it into Outlook

A quick, field-tested guide to get the add-in running against a **local
Synaplan** — no Microsoft 365 admin center needed. For the full picture see
[`docs/USER_GUIDE.md`](docs/USER_GUIDE.md) and
[`docs/AUTH_FLOW.md`](docs/AUTH_FLOW.md).

## Prerequisites (one-time)

- Node ≥ 22, Docker, and the [`synaplan`](https://github.com/metadist/synaplan)
  repo cloned next to this one (or set `SYNAPLAN_DIR=/path/to/synaplan`).
- Install dependencies and the local HTTPS dev certificate:

```bash
npm install
npx office-addin-dev-certs install   # creates ~/.office-addin-dev-certs
```

## 1. Start everything

```bash
./start-dev.sh        # or: make up
```

This brings up three things (idempotent — re-running skips what's already up):

| Service               | URL                      | What it is                                |
| --------------------- | ------------------------ | ----------------------------------------- |
| Synaplan Docker stack | `http://localhost:8000`  | Backend API (frontend dev server on 5173) |
| Synamail taskpane     | `https://localhost:3000` | Vite dev server the manifest points to    |
| HTTPS sign-in bridge  | `https://localhost:5174` | HTTPS front for the Synaplan frontend     |

> **Note:** on a cold start the Synaplan backend can take longer than the
> script's 60 s wait and you'll see `✗ Synaplan backend (:8000) did not come
up`. That's usually harmless — give it a minute, then verify:

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:8000/            # 200
curl -sk -o /dev/null -w "%{http_code}\n" https://localhost:3000/src/taskpane/taskpane.html  # 200
curl -sk -o /dev/null -w "%{http_code}\n" https://localhost:5174/          # 200
```

Logs land in `.dev-logs/{dev,bridge}.log`.

## 2. (WSL only) Trust the dev certificate on Windows

Outlook loads the taskpane from `https://localhost:3000`. If Windows doesn't
trust the dev CA, Outlook shows a **blank pane** with no error. WSL2 forwards
`localhost` ports to Windows automatically, but the certificate trust does
not carry over:

1. Copy the CA out of WSL:

   ```bash
   cp ~/.office-addin-dev-certs/ca.crt /mnt/c/Users/<you>/Downloads/
   ```

2. On Windows: double-click `ca.crt` → **Install Certificate** →
   **Local Machine** → **Trusted Root Certification Authorities**.
3. Sanity check: open `https://localhost:3000/src/taskpane/taskpane.html` in a
   Windows browser — it must load **without** a certificate warning.

On macOS / native Linux, `npx office-addin-dev-certs install` already trusted
the CA for you.

## 3. Sideload the manifest (no admin needed)

The dev `manifest.xml` points at `https://localhost:3000`, so you install it
as a _custom add-in_ into your own mailbox:

1. Open <https://aka.ms/olksideload> — it opens Outlook on the Web with the
   add-ins dialog.
2. **My add-ins → Custom Addins → Add a custom add-in → Add from file** and
   pick `manifest.xml` from this repo. From Windows the repo is reachable at
   `\\wsl$\<distro>\wwwroot\Synamail\manifest.xml` (or copy the file to
   Downloads first).
3. Open any email → toolbar **Apps** icon → **Synamail**. It can take a
   minute to appear after install, and it syncs to new Outlook for Windows
   automatically (same mailbox).

> **No "Add a custom add-in" option?** Your work tenant blocks custom
> add-ins — there is no non-admin workaround. Use a personal Outlook.com or
> Microsoft 365 developer-tenant account for local testing.

<!-- markdownlint keeps the two callouts as separate blockquotes -->

> `make sideload` (`office-addin-debugging start`) exists but can't launch
> Outlook from inside WSL — the manual upload above is the reliable route.
> For classic desktop Outlook there's also `make sync`, which copies a
> version-bumped manifest to `C:\addin-catalog` for a shared-folder catalog.

## 4. Sign in to your local Synaplan

In the Synamail sign-in view, set the server URL to:

- **local:** `https://localhost:5174` (the dev bridge — default in dev)
- **live:** `https://web.synaplan.com` (or your self-hosted host)

Sign-in is always a real round-trip — there is no mock mode. To switch
servers later: **Settings → Reset saved settings**, then sign in again.

## 5. Day-to-day

- Vue/TS changes are served live by Vite — just reopen the taskpane. No
  re-sideload needed.
- Only re-upload the manifest when `manifest.xml` itself changes.
- Stop the dev servers (Synaplan Docker keeps running):

```bash
./start-dev.sh stop   # or: make down
```
