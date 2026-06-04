# Synamail — Release & Launch Plan (v1.0)

**Goal:** take Synamail from its current working beta to a **public v1.0** released
on **two channels at once**:

1. **Microsoft AppSource** (the Microsoft Marketplace / "App Store") — the one-click
   install path for everyday Outlook users.
2. **Open Source** (public GitHub repo, Apache-2.0) — the self-host / contribute /
   audit path for developers and privacy-conscious orgs.

This document is the **single source of truth for the launch**. It starts from the
_current_ state (2026-06-04), names every remaining workstream, and orders them into
a step-by-step plan. It does **not** repeat the feature contract or the sprint
mechanics — those live in [`FEATURES.md`](../docs/FEATURES.md),
[`PLAN.md`](PLAN.md), [`STEPS.md`](STEPS.md), and [`STATUS.md`](STATUS.md). This is
the layer **on top**: features freeze, branding, help, and the two release tracks.

> **How to read this:** Sections 1–2 are context (where we are, where it runs).
> Section 3 is the user-experience spine (the two audiences + one entry point).
> Sections A–F are the work, in order. Section 7 is the timeline. Section 8 is the
> short list of decisions only you can make.

---

## 1. Where we are today (2026-06-04)

| Area                                                              | State                                                                                           | Gap to v1.0                                        |
| ----------------------------------------------------------------- | ----------------------------------------------------------------------------------------------- | -------------------------------------------------- |
| **Auth / sign-in**                                                | ✅ Live OAuth bridge shipped & deployed on web1/web2/web3. Real API keys, no copy-paste.        | Hardening + final E2E only.                        |
| **Read-mode AI** (summarise / translate / reply / classify / ask) | ✅ Wired to the real client; chat-id persistence done.                                          | One live smoke per action; pin response shapes.    |
| **Save-to-RAG + group picker**                                    | ✅ Real `.eml` capture, group picker, process levels.                                           | E2E against a live test user.                      |
| **Contact knowledge base**                                        | 🟡 Save-side done; search/view side partially wired.                                            | Finish `ContactKnowledgeBase.vue` search.          |
| **Compose-mode AI**                                               | 🟡 Calls the real client; "insert from RAG" needs polish.                                       | E2E + polish.                                      |
| **Mail Routes** (per-mail triggers)                               | 🟡 New, in-flight (`MailRoutes.vue`, building blocks landed).                                   | **Scope decision for v1.0 — see §A.**              |
| **RULE integration** (Synapse routing)                            | ⏳ Pending `PromptController` write-path audit.                                                 | **Likely a v1.1 deferral — see §A.**               |
| **Branding / look & feel**                                        | 🟥 Only a source mascot (`assets/source/parrot.svg`). Icons are placeholders.                   | Full brand system + store art — **§B**.            |
| **Help / docs**                                                   | 🟡 Strong dev docs + `INSTALL.md`. No "simple people" guide; no single entry point.             | Dual-audience doc set — **§C**.                    |
| **AppSource readiness**                                           | 🟡 Manifest valid; checklist exists; not submitted.                                             | Production hosting + assets + submission — **§D**. |
| **Open-source readiness**                                         | 🟡 Apache-2.0 `LICENSE` present; public org is **`metadist`** (`github.com/metadist/Synamail`). | Public-repo hygiene + launch — **§E**.             |

**One-line summary:** the _engine_ is essentially done and live; what remains is
**finish two features, dress it up (brand), write the human-facing docs, and run the
two release gauntlets.**

The pre-commit gate (`make ci-local`) is green and **non-negotiable** for every
change below.

---

## 2. Where Synamail runs — the platform reality (incl. your Linux/PWA question)

Synamail is a **web add-in** (Office.js). It runs inside Outlook's web runtime, not
as a native binary. So "which Outlook" matters more than "which OS".

| Client                                                 | Supported | Notes                                                       |
| ------------------------------------------------------ | --------- | ----------------------------------------------------------- |
| **Outlook on the Web** (Edge, Chrome, Safari, Firefox) | ✅        | The universal path. Works on **Windows, macOS, and Linux**. |
| **New Outlook for Windows**                            | ✅        | Same web runtime as OWA.                                    |
| **Classic Outlook 2024 (Windows)**                     | ✅        | WebView2-based.                                             |
| **Outlook for Mac** (new + classic)                    | ✅        |                                                             |
| **Outlook mobile (iOS/Android)**                       | ⏳        | Limited feature set — out of scope for v1.0.                |

### Your Linux/PWA question — answered

**Yes, it works for Linux users — through Outlook on the Web, which can be installed
as a PWA.** There is no native Outlook desktop app for Linux, so Linux users open
Outlook in a browser (`outlook.office.com` or `outlook.live.com`) and, if they like,
click the browser's **"Install app"** to get the OWA PWA. Synamail runs **inside that
PWA** because it's the exact same web add-in runtime as OWA — no separate work needed.

**One hard requirement to make explicit in our docs:** add-ins only run when the
mailbox is **Microsoft-hosted (Microsoft 365 / Exchange Online / Outlook.com)**.
They do **not** load against a plain IMAP/Gmail account added to Outlook web. This is
a Microsoft platform constraint, not a Synamail limitation.

**Action:** add "Outlook on the Web (Linux, Chrome + Edge PWA)" as a row in the
Sprint-4 client test matrix (`APPSOURCE_CHECKLIST.md` → "Functionality across
clients") so we verify and can advertise it.

---

## 3. Two audiences, one entry point (the UX spine of the launch)

The brief asks for a **clear entry point for both user groups** with **short, easy
leading** to the right next step. We solve this with a **single front door that
forks immediately** into two lanes.

### 3.1 The front door

`README.md` (and the GitHub repo home, and the AppSource "Learn more" link) opens
with one question and two buttons:

```
                 Synamail — your Synaplan workspace, inside Outlook

        ┌───────────────────────────────┐   ┌───────────────────────────────┐
        │  👤  I just want to use it      │   │  🛠  I want to build / host it  │
        │  → Install from AppSource       │   │  → Run from source / self-host  │
        │  → 60-second first-run guide     │   │  → Architecture + contributing  │
        └───────────────────────────────┘   └───────────────────────────────┘
```

- **Left lane (everyday users):** never sees a terminal. Goes to AppSource → installs
  → one short first-run guide with screenshots.
- **Right lane (developers / admins):** goes to the existing dev docs (`INSTALL.md`,
  `ARCHITECTURE.md`, `CONTRIBUTING.md`).

### 3.2 The document map (target state)

| Audience     | Document                                            | Status                       | One-line job                                             |
| ------------ | --------------------------------------------------- | ---------------------------- | -------------------------------------------------------- |
| **Everyone** | `README.md`                                         | ✏️ rewrite as the front door | Fork to the right lane in <10 s.                         |
| **User**     | `docs/USER_GUIDE.md` _(new)_                        | 🆕 create in §C              | "Simple people" walkthrough, screenshot-led, task-based. |
| **User**     | `INSTALL.md` → "For end users"                      | ✅ exists, refine            | One-click AppSource install + first run.                 |
| **User**     | In-app help (`?` link → USER_GUIDE)                 | 🆕 small UI add              | Help is one tap away inside the taskpane.                |
| **User**     | FAQ / Troubleshooting (user-level)                  | 🆕 section in USER_GUIDE     | "It didn't work" answers in plain language.              |
| **Dev**      | `INSTALL.md` → "For developers"                     | ✅ excellent                 | Field-tested sideload guide.                             |
| **Dev**      | `docs/ARCHITECTURE.md`                              | ✅                           | How it's built.                                          |
| **Dev**      | `docs/CONTRIBUTING.md` + `docs/COMMIT_PROCESS.md`   | ✅                           | How to contribute.                                       |
| **Dev**      | `docs/SELF_HOSTING.md` _(new, optional)_            | 🆕 §E                        | Point your own build at your own Synaplan.               |
| **Both**     | `SECURITY.md`, `CODE_OF_CONDUCT.md`, `CHANGELOG.md` | 🆕 §E                        | OSS table stakes.                                        |

### 3.3 The "short, easy leading" principle

Every user-facing page ends with **exactly one obvious next step** ("Next: sign in →",
"Next: summarise your first email →"). No page dumps the whole feature list. The
USER_GUIDE is **task-first** ("How do I…"), not feature-first.

---

## A. Feature freeze for v1.0

**Decide what ships in v1.0 vs v1.1 — then stop adding.** A frozen scope is what makes
the AppSource + OSS launch finishable. Proposal:

### A.1 In v1.0 (must work end-to-end, with E2E green)

- Sign in / sign out (live OAuth bridge). ✅
- Read mode: Summarise, Translate, Draft reply, Classify, Ask follow-ups.
- Save to knowledge base (group picker + processing levels).
- Contact knowledge base: save + **search** (finish the search side).
- Compose mode: Draft from prompt, Improve/Shorten/Translate selection, Insert from RAG.
- Settings: account, instance/self-host override, language, reset.

### A.2 Mail Routes & RULE integration — DECIDED (2026-06-04)

- **Mail Routes** (`MailRoutes.vue`): **ships in v1.0 labelled "Preview/Beta".** The
  tab must carry a visible "Preview" badge, must never silently act on a user's mail,
  and must **not** be advertised as a headline feature in the AppSource listing (to
  avoid the "advertised feature doesn't fully work" certification failure). Treat it as
  opt-in, transparent, and undoable per `FEATURES.md` §7.
- **RULE integration** (Synapse routing): **deferred to v1.1.** `RuleEditor.vue` is not
  part of v1.0 scope. No `PromptController` write-path audit is needed for launch. Keep
  the documented fallback in `FEATURES.md` §5.4 as the v1.1 starting point.

> **Why freeze hard:** AppSource certification fails if an _advertised_ feature
> doesn't fully work. Preview-labelled, non-advertised surfaces are acceptable; cut
> everything else.

### A.3 Definition of "feature-done" (the gate for every A.1 item)

`make ci-local` green **and** a Playwright E2E against the live test user **and** a
manual smoke on OWA + new Outlook for Windows + Outlook on Mac.

**Done when:** the A.1 list is all green by this definition and the A.2 decisions are
recorded here.

---

## B. Branding & look-and-feel

Today there is one asset: `assets/source/parrot.svg` (a parrot mascot silhouette).
Everything else (ribbon icons, store art) is a placeholder. v1.0 needs a coherent,
on-brand system that also reads correctly in Outlook's light **and** dark themes.

### B.1 Brand foundations (decide first)

- **Name:** `Synamail` (locked in manifest). Pre-validate it against Microsoft's
  name-collision check before submission (§D).
- **Mascot/logo:** confirm the **parrot** is the brand mark. If yes, produce a clean
  logo lockup (mark + wordmark) from the existing SVG. A parrot = "repeats/echoes your
  words, multilingual" — a good fit for an AI mail assistant; worth keeping.
- **Colour & type:** inherit from the Synaplan brand so the add-in feels native to the
  product family. Tokens already flow through `tokens.css`/`app.css` — extend, don't
  fork. **Never** hardcode hex; add light+dark token pairs for any new need.

### B.2 Deliverables

| Asset             | Spec                                                       | Where                           |
| ----------------- | ---------------------------------------------------------- | ------------------------------- |
| Logo lockup (SVG) | mark + wordmark, light & dark variants                     | `assets/source/`                |
| Ribbon icons      | 16, 32, 64, 80, 128 PNG, on-brand, dark-safe               | `assets/icon-*.png`             |
| Store hero        | 256, 512 PNG                                               | `assets/store/hero-*.png`       |
| Screenshots       | 1366×768, real content (no Lorem Ipsum), 6 views           | `assets/store/screenshot-*.png` |
| Screencast        | ≤60 s, 1080p, captioned: summarise → save to contact → ask | `assets/store/screencast.mp4`   |
| Social/OG image   | for the GitHub repo + README header                        | `assets/store/`                 |

### B.3 In-product polish pass

Run the existing **design verification checklist** (`GUI_DEFINITIONS.md` §6) across
every view: light/dark parity, focus rings, no 320↔450px breakage, loading/empty/error
states everywhere, axe-core on high-contrast. This is the "look and feel" sign-off.

**Done when:** every asset in B.2 exists at final quality, the GUI checklist is fully
ticked, and a non-team person calls the add-in "polished" in a 2-minute look-test.

---

## C. Help & documentation (the "developer howto" + "simple people howto")

### C.1 `docs/USER_GUIDE.md` — the "simple people" guide (NEW, the big gap)

Plain language, screenshot-led, **task-first**. No jargon, no terminal. Proposed
outline (each section = one screenshot + ≤5 steps + one "Next →"):

1. **What is Synamail?** — one paragraph + one picture.
2. **Install it** (60 seconds) — AppSource, per client, with the "Apps" icon shown.
3. **Connect your Synaplan account** — the sign-in window, the Connect button.
4. **Summarise an email** — the killer first action (build the "aha" fast).
5. **Translate / Draft a reply / Classify** — one mini-section each.
6. **Save an email to your knowledge base** — and what "knowledge base" means in 1 line.
7. **Build a knowledge base per contact** — search by sender/recipient.
8. **Write better emails (Compose mode).**
9. **Settings you might want** — language, self-hosted instance.
10. **Privacy in plain words** — "nothing leaves Outlook until you click."
11. **It didn't work?** — top 6 user-level fixes (not the dev troubleshooting table).

### C.2 Developer howto — mostly done, needs stitching

The dev path already exists and is strong: `INSTALL.md` (For developers),
`ARCHITECTURE.md`, `CONTRIBUTING.md`, `COMMIT_PROCESS.md`, `SYNAPLAN_INTEGRATION.md`.
Remaining work:

- Add `docs/SELF_HOSTING.md` (optional, §E) for "run my own build against my own Synaplan."
- Ensure the README "build/host" lane links these in reading order.

### C.3 In-app help

Add a small **"? Help"** link (footer of SignIn + Settings) that opens the USER_GUIDE
(hosted page) via `Office.context.ui.openBrowserWindow`. One tap to help, from inside
Outlook. i18n both strings (en + de).

### C.4 Hosting the help

User-facing help (USER_GUIDE, privacy, terms, support) must be **publicly reachable
with no auth wall** (AppSource requirement). **Decided:** the public user reference
lives on **`docs.synaplan.com`** (the `synaplan-docs` site) as the new "Outlook Add-in
(Synamail)" page. Privacy/Terms reuse the existing Synaplan pages. The in-app **? Help**
link and the AppSource listing both point at the `docs.synaplan.com` page.

**Done when:** USER_GUIDE.md exists (en + de), is publicly hosted, linked from the
README user lane, from AppSource, and from the in-app "? Help" link.

---

## D. Release Track 1 — Microsoft AppSource

This is **Sprint 4** in the existing plan. Run it against
[`APPSOURCE_CHECKLIST.md`](APPSOURCE_CHECKLIST.md). Ordered steps:

1. **Production hosting** — host the taskpane SPA at **`https://web.synaplan.com/addin/`**
   (decided — no new subdomain; served by the existing image/cluster). Strict CSP, no
   third-party scripts. Point `manifest.xml` `SourceLocation` at `web.synaplan.com/addin/`,
   keep `web.synaplan.com` in `AppDomains`, and **strip both the `localhost:5174` dev
   AppDomain and the now-unused `addin.synaplan.com` AppDomain.** (`STEPS.md` 4.1)
2. **Unified manifest** — generate `manifest.unified.json`; CI validates both;
   sideload-test the unified one in new Outlook for Windows. (`STEPS.md` 4.2)
3. **Brand assets** — from §B (icons, hero, screenshots, screencast). (`STEPS.md` 4.3)
4. **Store copy + legal URLs** — `assets/store/copy.md` (title, short/long description,
   keywords, support) in **en + de**; publicly hosted Privacy + Terms + Support pages
   linked from the manifest. (`STEPS.md` 4.4)
5. **Reviewer enablement** — demo M365 tenant + pre-seeded Synaplan account + sign-in
   walkthrough in the reviewer notes.
6. **Cross-client verification** — the full client matrix (incl. **Linux OWA/PWA** row
   from §2) smoke-tested and recorded.
7. **Submit via Partner Center** (Microsoft 365 and Copilot program). Expect 1–2 review
   iterations; **4–8 weeks calendar** turnaround. (`STEPS.md` 4.5)

**Done when:** `office-addin-validator` is 0-warnings on both manifests, the checklist
is fully ticked, and the submission is **accepted** (or review comments tracked +
resubmitted).

---

## E. Release Track 2 — Open Source launch

Apache-2.0 `LICENSE` is already in place. This track makes the repo safe and inviting
to publish publicly. It can run **in parallel** with §D.

### E.1 Repo hygiene (must do before going public)

- [ ] **Secret scrub** — scan full history (`gitleaks` / `trufflehog`) for keys,
      tokens, internal hosts. The repo references live infra (web1/web2/web3,
      `web.synaplan.com`) — confirm nothing sensitive (real API keys, `.env`) is or
      ever was committed.
- [x] **Org decided: `metadist`** — clone URLs now point at
      `github.com/metadist/Synamail` (matches `ghcr.io/metadist/...`). Done in README,
      INSTALL, CONTRIBUTING, CHANGELOG.
- [ ] **`SECURITY.md`** — how to report a vulnerability (private channel), supported
      versions, disclosure window.
- [ ] **`CODE_OF_CONDUCT.md`** — Contributor Covenant.
- [ ] **`CHANGELOG.md`** — start at `1.0.0`, "Keep a Changelog" format.
- [ ] **Issue/PR templates + CODEOWNERS** — templates exist; confirm CODEOWNERS is set
      (`STEPS.md` 2.3) and branch protection on `main` is on.
- [ ] **License headers / NOTICE** — confirm Apache-2.0 `NOTICE` if needed; check 3rd-
      party license compatibility of npm deps (`license-checker`).
- [ ] **README front door** — the §3.1 two-lane rewrite.

### E.2 Self-host story (optional but high-value for OSS)

- [ ] `docs/SELF_HOSTING.md` — how to build the add-in and point it at a self-hosted
      Synaplan (the `<AppDomains>` + base-URL override mechanics already exist).
- [ ] Decide whether to publish **prebuilt manifests** (dev-tenant / self-host
      variants) as GitHub Release assets so non-builders can sideload too.

### E.3 Launch mechanics

- [ ] Tag `v1.0.0`, cut a **GitHub Release** with notes + the screencast + manifest
      asset(s).
- [ ] Flip the repo to **public**.
- [ ] Announce: link AppSource ↔ GitHub both ways (AppSource listing → "Open source on
      GitHub"; README → "Install from AppSource").

**Done when:** the repo is public, secret-clean, has SECURITY/COC/CHANGELOG, the README
forks cleanly into the two lanes, and `v1.0.0` is tagged + released.

### E.4 Sequencing note (AppSource ↔ OSS)

You can publish OSS **before, with, or after** AppSource. **Recommendation: go public
with OSS at or just before AppSource acceptance**, so the AppSource listing can link to
a live public repo and reviewers can see the source. Going public _too early_ (before
the brand/docs land) means the world's first impression is a half-dressed repo.

---

## F. Launch day & post-launch

- **Launch checklist:** AppSource live ✅, repo public ✅, help pages 200-OK from an
  anonymous network ✅, in-app "? Help" works ✅, support inbox/issue triage owner
  assigned ✅.
- **Day-2:** monitor first installs, AppSource ratings, GitHub issues. Keep a `v1.1`
  milestone for the deferred items (Mail Routes GA, RULE editing, mobile).
- **Telemetry:** opt-in only, no PII without user action (already a checklist item).

---

## 7. Consolidated milestone view

```
NOW ──► A. Feature freeze ─┐
        (finish ContactKB   │
         search, decide      ├─► B. Branding ──┐
         Mail Routes/RULE)   │                 │
                             │   C. Help docs ──┼─► D. AppSource submit ──► (4–8 wk MS review) ──► AppSource LIVE
                             │   (USER_GUIDE)   │                                                        ▲
                             └─────────────────┘   E. OSS hygiene + launch ──► repo PUBLIC ─────────────┘
                                                    (parallel)                         (link both ways)
```

Rough effort (excluding Microsoft's review calendar): **A** ~2–3 days, **B** ~2–4 days,
**C** ~2–3 days, **D** ~2–3 days of work, **E** ~1–2 days. The long pole is **Microsoft's
4–8 week review**, so start **D** as soon as **A/B/C** are done; run **E** alongside.

---

## 8. Decisions

### Locked (2026-06-04)

1. ✅ **Mail Routes** → ship in v1.0 as **"Preview/Beta"**, not advertised. (§A.2)
2. ✅ **RULE integration** → **deferred to v1.1.** (§A.2)
3. ✅ **Brand mark** → the **parrot** is confirmed; build the brand around it. (§B.1)
4. ✅ **OSS timing** → publish the public repo **at / just before AppSource
   acceptance**, so the listing and reviewers can link to live source. (§E.4)

5. ✅ **Production host** → **`web.synaplan.com/addin/`** (no new subdomain; served by
   the existing image/cluster). The `addin.synaplan.com` AppDomain in `manifest.xml`
   can be dropped at the Sprint-4 manifest bump. (§D.1)
6. ✅ **Public GitHub org** → **`metadist`** (`github.com/metadist/Synamail`). (§E.1)
7. ✅ **Help hosting** → developer docs stay in-repo; the **public, user-facing**
   reference lives on **`docs.synaplan.com`** (the `synaplan-docs` site, new
   "Outlook Add-in" page). The in-app **? Help** link and AppSource listing point
   there. Privacy/Terms reuse the existing Synaplan pages. (§C.4 / §D.4)

### All launch decisions are now locked. Remaining work is execution.

---

## References

- Feature contract: [`docs/FEATURES.md`](../docs/FEATURES.md)
- Sprint plan / steps / status: [`PLAN.md`](PLAN.md) · [`STEPS.md`](STEPS.md) · [`STATUS.md`](STATUS.md)
- UI spec + asset sizes: [`GUI_DEFINITIONS.md`](GUI_DEFINITIONS.md)
- AppSource gate: [`APPSOURCE_CHECKLIST.md`](APPSOURCE_CHECKLIST.md)
- Install (dev + end-user): [`INSTALL.md`](../INSTALL.md)
- Synaplan-side integration: [`docs/SYNAPLAN_INTEGRATION.md`](../docs/SYNAPLAN_INTEGRATION.md)
