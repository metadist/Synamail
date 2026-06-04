# Contributing to Synamail

Thanks for your interest in Synamail. This page is the contributor entry point — the canonical rules live in [`AGENTS.md`](../AGENTS.md), [`docs/COMMIT_PROCESS.md`](COMMIT_PROCESS.md), and the docs they reference.

## Before you start

1. Read [`docs/PROJECT_PLAN.md`](../docs/PROJECT_PLAN.md) — the 4-sprint plan tells you which features are landing when and which work is still planning vs. coding.
2. Read [`docs/FEATURES.md`](FEATURES.md) — the feature contract for the add-in.
3. Skim [`docs/ARCHITECTURE.md`](ARCHITECTURE.md) — auth flow + Synaplan endpoint usage.
4. Read [`docs/COMMIT_PROCESS.md`](COMMIT_PROCESS.md) — **mandatory** reading before opening a PR.

If you can answer "yes" to "do I understand which sprint this work belongs to?", you're ready.

## Setup (once)

```bash
# Clone
git clone https://github.com/metadist/Synamail.git
cd Synamail

# Make sure you're on a supported Node version (>=22; .nvmrc pins to 24)
nvm use            # or: fnm use, volta install, etc.

# Enable the local git hooks
git config core.hooksPath .githooks

# Install dependencies
npm install
```

**Node version policy:** minimum `>=22` (`engines.node` in `package.json`). Node 24 is the Active LTS as of May 2026 (`.nvmrc`); CI exercises both 22 and 24 on every PR.

If you're contributing during the planning phase (Sprint 1), only `git config core.hooksPath .githooks` is needed — `npm install` will work once Step 2.1 lands the project scaffold.

## Workflow

```bash
# 1. Start a topic branch
git checkout -b feat/your-thing

# 2. Make your changes
# (write code, run dev server with `make dev`, etc.)

# 3. Run the gate BEFORE every commit
make ci-local

# 4. Commit using Conventional Commits
git commit -m "feat(taskpane): add Translate button"

# 5. Push and open a PR
git push -u origin feat/your-thing
gh pr create --fill

# 6. Address review feedback with new commits (do not force-push)
# 7. Once approved + green, squash-merge from the PR UI
```

## Commit messages

We use [Conventional Commits](https://www.conventionalcommits.org/). See the [allowed types and examples in `docs/COMMIT_PROCESS.md` §1](COMMIT_PROCESS.md#1-conventional-commits).

Quick examples:

```
feat(rag): add contact-scoped knowledge-base search
fix(auth): clear roaming settings on 401
docs(plan): bump Sprint 3 duration to 3 weeks
ci: cache office-addin-validator binary
```

**Never** add attribution footers ("Generated with …", "Co-Authored-By: …") for AI tools.

## The pre-commit gate

```bash
make ci-local        # lint + check-types + test + validate + build
```

This is what CI will run on your PR. Run it locally first to keep the loop fast.

For E2E tests (Playwright):

```bash
make test-e2e
```

Required locally when you touch auth, the Synaplan client, or any view. Required in CI on all PRs.

## Reporting bugs / asking for features

Use the issue templates:

- [Bug report](../.github/ISSUE_TEMPLATE/bug.md)
- [Feature request](../.github/ISSUE_TEMPLATE/feature.md)

## Cross-repo work

If your change requires a backend modification or a frontend bridge-page update in `@/wwwroot/synaplan`, mention it in the PR description and link the matching PR. Synaplan's pre-commit gate is `make lint && make -C backend phpstan && make test && docker compose exec -T frontend npm run check:types` — honour it there. Production rollout to the 3-node `synaplan-platform` cluster happens after the Synaplan PR merges and CI republishes the Docker image (see `docs/SYNAPLAN_INTEGRATION.md`).

## Code of conduct

Be kind, be specific, be concise. Reviews are about the code, not the author. If you disagree, ask "why" before "no".

## Licence

By contributing, you agree your contributions are licensed under the project's [LICENSE](../LICENSE).
