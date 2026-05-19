# Install Synamail

This document covers two audiences. **Skip to the section that matches you:**

- **[For end users](#for-end-users-released-version)** — how to install Synamail from the Microsoft Marketplace (AppSource) into your Outlook, after the v1.0 release.
- **[For developers](#for-developers-running-from-source)** — how to run, sideload, and iterate on Synamail from a local checkout. This is where the WSL / HTTPS / M365-tenant-policy findings live.

> Status (as of May 2026): The user-facing install section is **forward-looking** — Synamail is not yet on AppSource. It ships with the **Sprint 4** release. The developer section is current and tested.

---

## For end users (released version)

### Quick install (one click)

Once Synamail is in AppSource, the install will look like this:

1. Open your Outlook (any of: Outlook on the Web, new Outlook for Windows, classic Outlook 2024, Outlook on Mac).
2. **Get Add-ins** → search **Synamail** → **Add**.
3. Sign in to your Synaplan workspace once. Done.

The exact entry point varies by client:

| Client                      | Where to find "Get Add-ins"                                      |
| --------------------------- | ---------------------------------------------------------------- |
| Outlook on the Web          | Open an email → toolbar **Apps** icon → **Add apps**             |
| New Outlook for Windows     | Ribbon **Home → Get Add-ins** (or **More → Get Add-ins**)        |
| Classic Outlook for Windows | Ribbon **Home → Get Add-ins**                                    |
| Outlook on Mac (new UI)     | Toolbar **Apps** icon                                            |
| Outlook for iOS / Android   | (Phase 3 — limited feature set, see `docs/ARCHITECTURE.md` §8.5) |

### What gets installed

A taskpane add-in. Synamail does **not**:

- read your mail in the background,
- send your mail anywhere on its own,
- show up as a sender or recipient.

You explicitly click each action ("Summarise", "Save to knowledge base", etc.); only then does the relevant email content reach your Synaplan workspace. The API key Synamail stores is per-mailbox, encrypted at rest by Exchange, and revocable from the Synaplan web UI.

### First-run setup

1. Click **Open Synamail** on any email or new draft.
2. The taskpane opens with a **Sign in to Synaplan** button. Click it.
3. A small Synaplan login window opens — pick **Google**, **GitHub**, **Microsoft**, **password**, or whichever provider your Synaplan instance offers.
4. After signing in, click **Connect this Outlook** in the confirmation page.
5. The window closes and Synamail is ready. You don't need to type or paste any API key.

If your organisation runs a self-hosted Synaplan, click **"Use a self-hosted instance"** under the Sign-in button **before** signing in, and enter your instance URL.

### What if my organisation blocks custom add-ins?

Many Microsoft 365 work/school tenants disable user-uploaded add-ins by policy. Synamail is a **Microsoft-certified AppSource add-in**, which is a different category and is normally allowed by default. If your admin has fully locked down the AppSource catalog too, ask them to:

1. Open the Microsoft 365 admin center → **Settings → Integrated apps**.
2. Search for **Synamail** → **Get it now** → assign to you or a security group.

### Uninstalling

1. **My add-ins → Synamail → Remove**.
2. (Optional) Sign in to your Synaplan workspace at `web.synaplan.com` and revoke the API key named **Outlook Add-in (...)** from the API Keys page.

### Privacy + permissions

| What Synamail can do       | When                                                                         |
| -------------------------- | ---------------------------------------------------------------------------- |
| Read the current email     | Only after **you** click an action in Synamail.                              |
| Read attachments           | Only after **you** click **Save to knowledge base** and tick the attachment. |
| Write a reply              | Only after **you** click **Draft reply**.                                    |
| Open an Office popup       | Only for **Sign in to Synaplan**.                                            |
| Talk to `web.synaplan.com` | On every action you trigger. No background activity.                         |

For the full privacy policy and the AppSource certification dossier, see `web.synaplan.com/privacy` and `web.synaplan.com/terms`.

---

## For developers (running from source)

This section is the **field-tested install guide** for running Synamail from a local checkout. It captures the install findings from May 2026 testing on WSL Ubuntu + Windows Outlook.

### Prerequisites

| Tool     | Required version                                                                                                                                                                                                                                                                      | How to install                                                                   |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| Node.js  | **≥ 22** (Active LTS is **24**, pinned in `.nvmrc`)                                                                                                                                                                                                                                   | [`fnm`](https://github.com/Schniz/fnm) recommended (see below). `nvm` works too. |
| npm      | ≥ 10 (ships with Node 22+)                                                                                                                                                                                                                                                            | Comes with Node                                                                  |
| Git      | any recent                                                                                                                                                                                                                                                                            | `apt install git` / Windows installer                                            |
| Outlook  | One of: Outlook on the Web (any browser), new Outlook for Windows, classic Outlook 2024, Outlook on Mac. **Must be a mailbox you can sideload custom add-ins to** — see [Tenant policy](#tenant-policy-why-work--school-accounts-block-custom-add-ins) below, this is the #1 blocker. | Use an existing account, or create a free [outlook.com](https://outlook.com) one |
| Python 3 | ≥ 3.10 (only for placeholder-icon regeneration; optional)                                                                                                                                                                                                                             | `apt install python3`                                                            |

#### Quick Node setup with `fnm` (recommended)

```bash
curl -fsSL https://fnm.vercel.app/install | bash
exec $SHELL                   # or: source ~/.bashrc
fnm install 24
fnm default 24
cd /path/to/Synamail
fnm use                       # reads .nvmrc → switches to 24
node --version                # → v24.x.x
```

### Step 1 — Clone and bootstrap

```bash
git clone https://github.com/<org>/Synamail.git
cd Synamail
make bootstrap     # enables .githooks/, runs npm install
make doctor        # verifies node/npm/hooks
```

If `make doctor` warns that Node is below 22, finish the `fnm` setup above first.

### Step 2 — Run the gate

```bash
make ci-local
```

You should see `ci-local: all gates passed` at the bottom — lint + type-check + 30 unit tests + manifest validate + production build. **Do this before you trust anything else in this guide.** If it fails on a fresh clone, the repo is broken; open an issue.

### Step 3 — Provision the HTTPS dev cert (one-time)

Office.js manifests require an HTTPS source location. The Vite dev server needs a trusted certificate for `https://localhost:3000`.

```bash
npx office-addin-dev-certs install
```

This writes a CA + leaf certificate to `~/.office-addin-dev-certs/`. On WSL it adds the CA to Linux's trust store (curl/wget will trust it), but **Windows browsers won't automatically trust it** — see step 5 below.

You can re-run this safely; it's idempotent.

### Step 4 — Start the Vite dev server

```bash
make dev
```

You should see:

```
  ➜  Local:   https://localhost:3000/
  ➜  Network: …
```

Leave this terminal running. The first hit to a route triggers a recompile so the first taskpane open after `make dev` will be ~1 s slow.

**Warning:** if you see `[vite] office-addin-dev-certs not provisioned — serving plain HTTP`, step 3 didn't complete. Outlook will refuse to sideload over HTTP.

### Step 5 — Make Windows / your OS trust the dev cert

#### WSL → Windows browser (most common path)

Open Edge or Chrome on Windows and visit:

```
https://localhost:3000/src/taskpane/taskpane.html
```

You'll see _"Your connection isn't private"_. Click **Advanced → Continue to localhost (unsafe)**. You should now see the **Synamail** taskpane render in the browser. That establishes a cert exception, and Outlook's iframe will reuse it.

This is the **30-second hack** and is enough for development.

#### Cleaner alternative — import the CA into Windows' Trusted Root store (do once, forever)

```bash
cp ~/.office-addin-dev-certs/ca.crt /mnt/c/Users/<you>/Downloads/
```

In Windows Explorer, double-click the file → **Install Certificate** → **Local Machine** (or Current User) → **Place all certificates in the following store** → **Trusted Root Certification Authorities** → confirm the warning → Finish. Reopen the browser. No more cert warnings ever.

#### Native macOS

```bash
sudo security add-trusted-cert -d -r trustRoot \
  -k /Library/Keychains/System.keychain \
  ~/.office-addin-dev-certs/ca.crt
```

#### Native Linux

`office-addin-dev-certs install` does it for you (it writes to `/usr/local/share/ca-certificates/`).

### Step 6 — Pick a sideload target (this is the bit that bites)

> The **#1 dev install failure** is here. Read [Tenant policy](#tenant-policy-why-work--school-accounts-block-custom-add-ins) before picking. Hard pill: most work/school M365 accounts **cannot** sideload Synamail without admin intervention.

You have four routes. Difficulty / capability listed:

| Route                                        | Effort | Works on                                  | Caveats                                                             |
| -------------------------------------------- | ------ | ----------------------------------------- | ------------------------------------------------------------------- |
| **A. Microsoft 365 Developer tenant**        | 5 min  | OWA + new Outlook for Windows + Mac       | Free, sandbox. **Recommended.**                                     |
| **B. Personal `outlook.live.com` account**   | 1 min  | OWA only                                  | Sometimes lags behind on requirement sets.                          |
| **C. Classic Outlook shared-folder catalog** | 10 min | Classic Outlook for Windows desktop only  | Doesn't need admin approval — bypasses tenant policy.               |
| **D. M365 admin centre deployment**          | 30 min | OWA + new Outlook + classic Outlook + Mac | You must be the tenant admin. Deployment takes 1–24 h to propagate. |

#### Route A — Microsoft 365 Developer Program (recommended)

The Developer Program gives you a free sandbox M365 tenant where **you are the admin** and can do anything.

1. Sign up at [developer.microsoft.com/microsoft-365/dev-program](https://developer.microsoft.com/en-us/microsoft-365/dev-program).
2. Choose **"Instant sandbox"** during setup (~5 minutes).
3. You'll get an alias like `you@<random>.onmicrosoft.com` and 25 pre-seeded fake users.
4. Sign in at `https://outlook.office.com` with that account.
5. Continue to [Step 7](#step-7--install-the-manifest-into-outlook).

This is also the tenant you'll use for Sprint 4 AppSource certification (Microsoft reviewers need a tenant they can sign into).

#### Route B — Personal `outlook.live.com`

If you already have a Hotmail / Outlook.com / Live personal account:

1. Open `https://outlook.live.com/mail/` (note: `.live.com`, **not** `.office.com`).
2. Sign in.
3. Continue to [Step 7](#step-7--install-the-manifest-into-outlook).

#### Route C — Classic Outlook desktop, shared folder catalog

Bypasses M365 tenant policy entirely on classic Outlook for Windows.

1. Create a folder on Windows: `C:\addin-catalog\`.
2. Copy the manifest into it:

   ```bash
   cp /wwwroot/Synamail/manifest.xml /mnt/c/addin-catalog/
   ```

3. Right-click the folder → **Properties → Sharing → Share**, share with yourself.
4. Note the UNC path (e.g. `\\YOUR-PC\addin-catalog`).
5. In classic Outlook: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
6. Paste the UNC path → tick **"Show in menu"** → **Add catalog** → **OK**.
7. **Restart classic Outlook**.
8. **Home → All Apps → Get Add-ins → My add-ins → Custom add-ins** — your shared folder's manifest is now listed. Install.

#### Route D — M365 admin centre

Only if you own / administer the tenant:

1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. **Office Add-in → Upload manifest file** → browse to `manifest.xml`.
3. Assign to yourself (or a test group).
4. **Wait 1–24 hours** for propagation. Yes, really.

### Step 7 — Install the manifest into Outlook

Once you've picked a route from Step 6 and are signed in to a tenant that allows sideloading:

#### Outlook on the Web (`outlook.office.com` or `outlook.live.com`)

1. Open any email (or click **New mail**).
2. Click the **Apps** icon in the message toolbar (four-squares icon; older UIs: **... → Get Add-ins**).
3. Find **My add-ins** in the left sidebar (sometimes called **Manage your apps**).
4. Scroll to **Custom add-ins** → **+ Add a custom add-in** → **Add from File…**.
5. Select the manifest. From Windows, the manifest is at:

   ```
   \\wsl$\Ubuntu\wwwroot\Synamail\manifest.xml
   ```

   Or copy it out first to make life simpler:

   ```bash
   cp /wwwroot/Synamail/manifest.xml /mnt/c/Users/<you>/Downloads/manifest.xml
   ```

6. Confirm the security warning → **Install**.
7. The **Synamail** ribbon group with **Open Synamail** appears in both Read view and Compose view.

#### New Outlook for Windows

Same flow as OWA — new Outlook uses the same web add-in runtime:

1. **Home → More → Get Add-ins** (or **Settings → General → Add-ins → Manage add-ins**).
2. Same **My add-ins → Add a custom add-in → Add from File** path.
3. Pick `\\wsl$\Ubuntu\wwwroot\Synamail\manifest.xml`.

#### Classic Outlook for Windows

If you went via Route C, the manifest is already in your trusted catalog — just **Get Add-ins → My add-ins → Custom add-ins → install**.

If you went via Route D, the admin-deployed app shows up in **My add-ins** within 24 h.

#### Outlook on Mac (new UI)

Same as OWA. Toolbar **Apps** icon → **Add custom add-in**.

### Step 8 — Verify the loop works

1. Open any email in your sideloaded Outlook.
2. Click **Open Synamail** in the ribbon.
3. The taskpane should show the **Sign in** view.
4. Click **Sign in to Synaplan** → a dialog flashes through (the local mock relay) → the view switches to **ReadMode**.
5. Click **Summarise** → after ~250 ms you'll see canned bullets containing "This is a mock summary — Sprint 3 swaps in the real Synaplan call."

If you see all that: **you're done. Sprint 2 acceptance criterion #1 is met.** Welcome to the dev loop.

To see Compose mode: click **New mail** → click **Open Synamail** in the compose ribbon.

### Step 9 — Switch from mock to live API (later)

Currently the auth flow uses a local mock relay (`src/dialog/auth-relay.html`) and the Synaplan client auto-selects `MockSynaplanClient` when the API key starts with `mock-key-`. Switching to the live `web.synaplan.com` flow requires:

1. The `/addin/connect` Vue route to land in `synaplan/frontend/` (Sprint 2.8 cross-repo PR). See `docs/SYNAPLAN_INTEGRATION.md`.
2. Once that ships and `synaplan-platform` rolls the new image, your sideloaded Synamail will automatically use the live flow on next sign-in — no code change in Synamail required.

---

## Tenant policy: why work / school accounts block custom add-ins

This is the friction point that consumes the most developer time. Microsoft's policy:

| Account type                           | Custom (sideloaded) add-ins | AppSource add-ins                      |
| -------------------------------------- | --------------------------- | -------------------------------------- |
| Personal Microsoft (outlook.live.com)  | ✅ allowed by user          | ✅ allowed by user                     |
| Microsoft 365 Developer Program tenant | ✅ allowed by you (admin)   | ✅ allowed by you (admin)              |
| Work / school M365 (default)           | ⛔ blocked by tenant policy | ⛔ usually blocked, ✅ if admin allows |
| Work / school M365 (admin-enabled)     | ⛔ still blocked            | ✅                                     |

The locked-down "Apps" page in a corporate M365 OWA looks like this (real screenshot from the May 2026 install test):

- A "Built for your org" section with apps the IT team has pushed (Viva Insights, Bing Maps, etc.).
- "Popular" / "Editor's choice" listings from AppSource.
- **No "Add a custom add-in" or "Upload manifest" button anywhere.**

If you see this UI, your tenant is locked. Switch to Route A or Route B from Step 6.

### Why this exists

Microsoft made user-uploaded add-ins a default-off tenant capability because a malicious manifest could:

- Read every email the user opens (`ReadWriteItem` permission is the standard).
- Exfiltrate to any HTTPS endpoint.
- Persist via roaming settings.

For genuine enterprise deployment of Synamail, the admin deploys via AppSource or via the Integrated Apps "Upload custom" route — both of which leave an audit trail. For development, you bypass the policy entirely with a dev tenant or personal account.

---

## Troubleshooting

| Symptom                                                             | Likely cause                                                                                      | Fix                                                                                                                                   |
| ------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| `make ci-local` fails on a fresh clone                              | Stale `node_modules` or wrong Node version                                                        | `rm -rf node_modules package-lock.json && make bootstrap && make ci-local`                                                            |
| Vite warns "office-addin-dev-certs not provisioned"                 | Step 3 didn't finish                                                                              | Re-run `npx office-addin-dev-certs install`                                                                                           |
| Browser visit to `https://localhost:3000` says "connection refused" | Vite isn't running, or wrong port                                                                 | Confirm `make dev` is running; the dev port is **3000**.                                                                              |
| "Couldn't install the add-in" with a URL / cert error               | Windows doesn't trust the dev cert                                                                | Visit `https://localhost:3000/src/taskpane/taskpane.html` in Edge/Chrome first; accept the warning.                                   |
| OWA sideload UI says "manifest.xml is not valid"                    | XML edit broke schema                                                                             | `make validate` should still pass; if not, `git diff manifest.xml` and revert the bad change.                                         |
| "Add a custom add-in" button is missing                             | Tenant policy blocks user sideload (Step 6!)                                                      | Switch to a dev tenant (Route A) or personal `outlook.live.com` (Route B).                                                            |
| Ribbon button doesn't appear after install                          | OWA cache or new-Outlook lag                                                                      | Hard-refresh OWA (Ctrl+F5); for new Outlook, restart the app. Can take 1–2 minutes.                                                   |
| Taskpane shows a white blank                                        | Vite died, port collision, or cert untrusted                                                      | Check the `make dev` terminal for errors. Confirm `https://localhost:3000/src/taskpane/taskpane.html` opens directly in your browser. |
| Sign-in dialog opens and immediately closes with an error           | State-nonce mismatch (this is the _correct_ behaviour for a malformed payload — security feature) | Try again; if it's reproducible, capture the `messageParent` payload from devtools and report.                                        |
| Clicking Summarise does nothing                                     | Mock client got mistaken for the real one                                                         | Check `localStorage` / roaming for an `apiKey` not starting with `mock-key-`. Sign out and back in.                                   |
| `npm install` is extremely slow on WSL                              | Filesystem on the Windows side                                                                    | Make sure the repo is under `/wwwroot/...` (Linux side), **not** `/mnt/c/...` (Windows side mounted into WSL — 10× slower).           |

---

## Architecture notes for the installer

- **Manifest source location:** `https://localhost:3000/src/taskpane/taskpane.html` during dev; `https://addin.synaplan.com/...` once Sprint 4 ships production hosting.
- **Storage:** API key + base URL live in `Office.context.roamingSettings` (encrypted per-mailbox by Exchange). Outside Office (tests, plain browser dev) they fall back to `localStorage`.
- **Base URL override:** Settings → "Synaplan instance" lets the user point at a self-hosted Synaplan **before** signing in. The next sign-in dialog opens against that URL.
- **CORS:** The Synaplan API at `web.synaplan.com` is configured `allow_origin: ['*']` (see `synaplan/backend/config/packages/nelmio_cors.yaml`). Your localhost dev origin works without any allow-list change.
- **Cert lifetime:** `office-addin-dev-certs install` provisions certs valid for ~30 days. Re-run the command when they expire.
- **Multiple developers / multiple machines:** every developer provisions their own cert via step 3 — certificates are not shared.

---

## Cross-platform notes

### WSL (Ubuntu / Debian on Windows)

- Repo should live under `/wwwroot/...` or `~/...` — the **Linux** filesystem. Putting the repo under `/mnt/c/...` (Windows filesystem mounted into WSL) makes everything 10× slower due to filesystem translation overhead.
- Windows Edge/Chrome can reach the WSL Vite server at `https://localhost:3000` automatically (WSL2 localhost forwarding).
- File access from Windows: `\\wsl$\Ubuntu\wwwroot\Synamail\manifest.xml`.

### Native Windows (PowerShell / WSL not used)

- Everything works the same way, but the `\\wsl$\` UNC path is replaced by a normal Windows path.
- `office-addin-dev-certs install` writes to `%USERPROFILE%\.office-addin-dev-certs\` and adds the CA to Windows' trust store directly — Step 5 isn't needed.

### Native macOS

- `office-addin-dev-certs install` may need the keychain password (it adds the CA to the System keychain).
- Outlook for Mac sideloads via the same UI as OWA.

### Native Linux (running Outlook in a browser only)

- `office-addin-dev-certs install` writes to `/usr/local/share/ca-certificates/` (may need sudo).
- Outlook desktop isn't available on Linux — you're restricted to Outlook on the Web. All Synamail features work in OWA, so this is fine for development.

---

## Related documents

- [`planning/STATUS.md`](planning/STATUS.md) — current build status, Sprint 2 acceptance state, next actions.
- [`docs/SYNAPLAN_INTEGRATION.md`](docs/SYNAPLAN_INTEGRATION.md) — what the Synaplan platform side needs (the bridge page).
- [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) — pre-commit gate, hooks, CI.
- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — full architecture brief.
- [`AGENTS.md`](AGENTS.md) — rules every contributor (and AI agent) follows.
- [`planning/STEPS.md`](planning/STEPS.md) — sprint-by-sprint execution plan.
- [`planning/APPSOURCE_CHECKLIST.md`](planning/APPSOURCE_CHECKLIST.md) — what Microsoft will check at certification.

---

## Changelog (install-process findings)

| Date       | Finding                                                                                                                                                                        |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 2026-05-16 | Initial dev-cert + Vite HTTPS plumbing verified end-to-end. Manifest validates clean.                                                                                          |
| 2026-05-19 | Tenant-policy block surfaced on a corporate M365 OWA — no "Add custom add-in" UI present. Documented the four routes (dev tenant / personal / classic catalog / admin upload). |
| 2026-05-19 | `fnm` + Node 24 confirmed as the dev runtime. `.nvmrc` pins 24; `engines.node: ">=22"`.                                                                                        |
