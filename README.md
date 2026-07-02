# Synamail

<img width="1721" height="1318" alt="Outlook_Plugin" src="https://github.com/user-attachments/assets/12b629c0-4a04-4502-8423-0abcc09a009a" />

**Your Synaplan AI workspace, right inside Outlook.** Summarise, translate, draft
replies, classify, save emails to your knowledge base, and ask follow-up questions —
without leaving your inbox. Works in Outlook on the Web (Windows, Mac, **and Linux**),
new Outlook for Windows, classic Outlook 2024, and Outlook on Mac.

Synamail is the Outlook companion to **[Synaplan](https://www.synaplan.com)** — the
open-source AI knowledge platform. Every action you click runs in **your** Synaplan
workspace, so you need an account first:

- **Just want to use it?** Create a **free account at
  [web.synaplan.com](https://web.synaplan.com)**, then install the add-in (below).
- **Run your own?** Self-host the platform from
  **[github.com/metadist/synaplan](https://github.com/metadist/synaplan)** and point
  Synamail at your instance under _"Use a self-hosted instance"_ on the sign-in screen.

---

## Pick your path

| 👤 **I just want to use it**                                                | 🛠 **I want to build / host it**                                                    |
| --------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| Install Synamail and start in 60 seconds — no technical setup.              | Run from source, contribute, or self-host against your own Synaplan.                |
| → **[User Guide](docs/USER_GUIDE.md)** (start here)                         | → **[Developer install & sideload](INSTALL.md#for-developers-running-from-source)** |
| → [Install from the Marketplace](INSTALL.md#for-end-users-released-version) | → [Architecture](docs/ARCHITECTURE.md) · [Contributing](docs/CONTRIBUTING.md)       |
| → [Privacy in plain words](docs/USER_GUIDE.md#10-privacy-in-plain-words)    | → [Commit process](docs/COMMIT_PROCESS.md) · [Security policy](SECURITY.md)         |

> Not sure if you can use it at work? Many corporate mailboxes restrict add-ins — see
> [the User Guide](docs/USER_GUIDE.md#2-install-synamail) for what to ask your IT team.

---

> The rest of this README is for **builders**. Everyday users should head to the
> [User Guide](docs/USER_GUIDE.md).

## Project structure

```
Synamail/
├── README.md                       — you are here
├── INSTALL.md                      — end-user install (post-release) + developer sideload guide
├── AGENTS.md                       — workspace rules (AI + human contributors)
├── LICENSE
├── Makefile                        — make help, make ci-local, make sideload, make sync-plugin, …
├── manifest.xml                    — Outlook add-in manifest, DEV (validated by CI)
├── manifest_1-0-2_prod.xml         — PRODUCTION manifest for the AppSource submission (versioned filename)
├── synamail-plugin/                — Synaplan-side plugin (Contact AI Profiling) — released
│                                     to synaplan/plugins/synamail via `make sync-plugin`
├── package.json / tsconfig*.json / vite.config.ts / vitest.config.ts / playwright.config.ts
├── eslint.config.js / commitlint.config.cjs
├── .editorconfig / .gitattributes / .gitignore / .markdownlint.jsonc / .nvmrc / .prettierrc.json / .prettierignore
├── .githooks/
│   ├── pre-commit                  — runs `make ci-local` (or fast docs lint)
│   ├── commit-msg                  — Conventional Commits validator
│   └── pre-push                    — reminds to run E2E when auth/views change
├── .github/
│   ├── workflows/ci.yml            — staged CI (docs / commitlint / manifest / build / e2e)
│   ├── workflows/deploy.yml        — build + publish the add-in host image to GHCR
│   ├── dependabot.yml              — weekly grouped updates
│   ├── PULL_REQUEST_TEMPLATE.md
│   └── ISSUE_TEMPLATE/             — bug + feature templates
├── Dockerfile / .dockerignore      — tiny static host image (serves addin.synaplan.com)
├── deploy/                         — container deploy + Caddy ingress runbook (self-host the add-in host)
├── docs/
│   ├── USER_GUIDE.md               — everyday-user guide (no command line)
│   ├── PROJECT_PLAN.md             — plan, status, and the AppSource checklist
│   ├── RELEASE_SUMMARY.md          — v1 scope summary
│   ├── ARCHITECTURE.md             — technical architecture, auth, API surface
│   ├── AUTH_FLOW.md                — authoritative sign-in / sign-out flow
│   ├── FEATURES.md                 — feature contract
│   ├── CONTACT_PROFILING.md        — Contact AI Profiling design
│   ├── CONTRIBUTING.md             — contributor entry point
│   ├── COMMIT_PROCESS.md           — commits, branches, reviews, releases
│   ├── GLOSSARY.md                 — canonical terminology
│   └── SYNAPLAN_INTEGRATION.md     — what changes in synaplan / synaplan-platform
├── src/                            — Vue 3 + TypeScript add-in source
│   ├── i18n.ts / locales/
│   ├── shared/                     — client, prompts, types
│   ├── taskpane/                   — entry, views, components, composables
│   ├── commands/                   — Office function-file shell
│   └── dialog/                     — auth-relay (mock until cross-repo bridge ships)
├── tests/
│   ├── unit/                       — Vitest unit tests
│   ├── component/                  — Vue component tests
│   ├── e2e/                        — Playwright suite
│   └── setup.ts                    — Office.js stub for Vitest
├── scripts/                        — one-off dev utilities
└── assets/                         — icons + AppSource store assets
```

## Reading order for new contributors

1. [`AGENTS.md`](AGENTS.md) — the rules that apply to every commit (humans included).
2. [`docs/CONTRIBUTING.md`](docs/CONTRIBUTING.md) — the contributor flow in 90 seconds.
3. [`docs/GLOSSARY.md`](docs/GLOSSARY.md) — terminology mapping (read before the others).
4. [`docs/FEATURES.md`](docs/FEATURES.md) — what the add-in does.
5. [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — how it does it (auth, API surface, security).
6. [`docs/AUTH_FLOW.md`](docs/AUTH_FLOW.md) — the authoritative sign-in / sign-out flow.
7. [`docs/SYNAPLAN_INTEGRATION.md`](docs/SYNAPLAN_INTEGRATION.md) — what changes (and doesn't) in `synaplan` / `synaplan-platform`.
8. [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) — Conventional Commits, branches, PR, release.
9. [`docs/PROJECT_PLAN.md`](docs/PROJECT_PLAN.md) — plan, status, and the AppSource checklist.
10. [`INSTALL.md`](INSTALL.md) — end-user install (post-release) + developer sideload guide with the field-tested gotchas.

## Quick start

```bash
git clone https://github.com/metadist/Synamail.git
cd Synamail
make bootstrap       # enables git hooks, installs dependencies
make doctor          # verifies your local toolchain
make help            # discover the rest
```

The pre-commit gate that the hook enforces:

```bash
make ci-local        # lint + check-types + test + validate + build
```

This is exactly what CI runs on your PR. See [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) and [`docs/CONTRIBUTING.md`](docs/CONTRIBUTING.md) for the full process.

## The Synaplan ecosystem

Synamail is one client of the open-source **Synaplan** platform:

- **[www.synaplan.com](https://www.synaplan.com)** — product site and hosted
  workspaces; the place to **create a free account**.
- **[web.synaplan.com](https://web.synaplan.com)** — the default workspace Synamail
  connects to out of the box (you can point it at any instance at sign-in).
- **[github.com/metadist/synaplan](https://github.com/metadist/synaplan)** — the
  Synaplan platform (PHP/Symfony + Vue 3), which you can self-host.

Contact AI Profiling ships as a Synaplan plugin in
[`synamail-plugin/`](synamail-plugin/) and is released into
`synaplan/plugins/synamail` via `make sync-plugin` — details in
[`docs/SYNAPLAN_INTEGRATION.md`](docs/SYNAPLAN_INTEGRATION.md).

## Sprint summary

| Sprint | Goal                                                               | Duration                                          |
| ------ | ------------------------------------------------------------------ | ------------------------------------------------- |
| 1      | Plan, design, test definitions, GUI spec, **engineering scaffold** | 3–5 days                                          |
| 2      | GUI + sideload + live `web.synaplan.com` identification            | 2–3 weeks                                         |
| 3      | Live AI features, RAG, RULE, contact KB                            | 2–3 weeks                                         |
| 4      | AppSource publishing                                               | 1–2 weeks + Microsoft review (4–8 weeks calendar) |
