# Synamail

MS Outlook add-in for Synaplan. Work directly with your Outlook client and your own Synaplan account, locally or remotely.

## Project structure

```
Synamail/
├── README.md                       — you are here
├── INSTALL.md                      — end-user install (post-release) + developer sideload guide
├── AGENTS.md                       — workspace rules (AI + human contributors)
├── LICENSE
├── Makefile                        — make help, make ci-local, make sideload, …
├── manifest.xml                    — Outlook add-in manifest (validated by CI)
├── package.json / tsconfig*.json / vite.config.ts / vitest.config.ts / playwright.config.ts
├── eslint.config.js / commitlint.config.cjs
├── .editorconfig / .gitattributes / .gitignore / .markdownlint.jsonc / .nvmrc / .prettierrc.json / .prettierignore
├── .githooks/
│   ├── pre-commit                  — runs `make ci-local` (or fast docs lint)
│   ├── commit-msg                  — Conventional Commits validator
│   └── pre-push                    — reminds to run E2E when auth/views change
├── .github/
│   ├── workflows/ci.yml            — staged CI (docs / commitlint / manifest / build / e2e)
│   ├── dependabot.yml              — weekly grouped updates
│   ├── PULL_REQUEST_TEMPLATE.md
│   └── ISSUE_TEMPLATE/             — bug + feature templates
├── docs/
│   ├── ARCHITECTURE.md             — technical architecture, auth, API surface
│   ├── COMMIT_PROCESS.md           — commits, branches, reviews, releases
│   ├── CONTRIBUTING.md             — contributor entry point
│   ├── FEATURES.md                 — feature contract
│   ├── GLOSSARY.md                 — canonical terminology
│   └── SYNAPLAN_INTEGRATION.md     — what changes in synaplan / synaplan-platform
├── planning/
│   ├── PLAN.md                     — 4-sprint plan
│   ├── STATUS.md                   — current build / sprint status
│   ├── STEPS.md                    — step-by-step plan, 4 sprints
│   ├── GUI_DEFINITIONS.md          — views, components, asset list
│   └── APPSOURCE_CHECKLIST.md      — Microsoft submission gate
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

1. [`planning/PLAN.md`](planning/PLAN.md) — what the four sprints deliver and when.
2. [`AGENTS.md`](AGENTS.md) — the rules that apply to every commit (humans included).
3. [`docs/CONTRIBUTING.md`](docs/CONTRIBUTING.md) — the contributor flow in 90 seconds.
4. [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) — Conventional Commits, branches, PR, release.
5. [`docs/GLOSSARY.md`](docs/GLOSSARY.md) — terminology mapping (read before the others).
6. [`docs/FEATURES.md`](docs/FEATURES.md) — what the add-in does.
7. [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — how it does it.
8. [`planning/STEPS.md`](planning/STEPS.md) — how we ship it, step by step.
9. [`planning/GUI_DEFINITIONS.md`](planning/GUI_DEFINITIONS.md) — what it looks like.
10. [`planning/STATUS.md`](planning/STATUS.md) — current build / sprint status.
11. [`docs/SYNAPLAN_INTEGRATION.md`](docs/SYNAPLAN_INTEGRATION.md) — what changes (and doesn't) in `synaplan` / `synaplan-platform`.
12. [`INSTALL.md`](INSTALL.md) — end-user install (post-AppSource release) + developer sideload guide with the field-tested gotchas.

## Quick start

```bash
git clone https://github.com/<org>/Synamail.git
cd Synamail
make bootstrap       # enables git hooks, installs deps (once Sprint 2.1 has landed)
make doctor          # verifies your local toolchain
make help            # discover the rest
```

The pre-commit gate that the hook enforces:

```bash
make ci-local        # lint + check-types + test + validate + build
```

This is exactly what CI runs on your PR. See [`docs/COMMIT_PROCESS.md`](docs/COMMIT_PROCESS.md) and [`docs/CONTRIBUTING.md`](docs/CONTRIBUTING.md) for the full process.

## Sprint summary

| Sprint | Goal                                                               | Duration                                          |
| ------ | ------------------------------------------------------------------ | ------------------------------------------------- |
| 1      | Plan, design, test definitions, GUI spec, **engineering scaffold** | 3–5 days                                          |
| 2      | GUI + sideload + live `web.synaplan.com` identification            | 2–3 weeks                                         |
| 3      | Live AI features, RAG, RULE, contact KB                            | 2–3 weeks                                         |
| 4      | AppSource publishing                                               | 1–2 weeks + Microsoft review (4–8 weeks calendar) |
