## Summary

<!-- One sentence on what changed and why. -->

## Type of change

<!-- Tick one. -->

- [ ] `feat` — new user-visible feature
- [ ] `fix` — bug fix
- [ ] `refactor` — no functional change
- [ ] `perf` — performance improvement
- [ ] `test` — tests only
- [ ] `docs` — documentation only
- [ ] `build` / `ci` — build, deps, workflows
- [ ] `chore` / `style` — housekeeping

## Sprint / Step

<!-- Which sprint and step from planning/STEPS.md does this advance? -->

Sprint `__`, Step `__.__`

## Changes

-

## Verification

### Pre-commit gate (mandatory)

- [ ] `make ci-local` passes locally (lint + check-types + test + validate + build)

### Tests

- [ ] Unit / component tests added or updated (Vitest)
- [ ] E2E tests added or updated (Playwright + sideload) — required if you touched auth, the Synaplan client, or any view
- [ ] Manifest still validates (`make validate`)

### Manual smoke (where relevant)

- [ ] Sideloaded into Outlook on the Web (Edge)
- [ ] Sideloaded into new Outlook for Windows
- [ ] Sideloaded into classic Outlook 2024 / Mac (release PRs only)

### Cross-repo

- [ ] Touches `synaplan/` (frontend or backend) — linked PR: \_\_
- [ ] Requires a rolling restart on `synaplan-platform` (web1/web2/web3) — coordinate via `docs/SYNAPLAN_INTEGRATION.md`

### Acceptance criterion

<!-- Quote the "Done when…" line from STEPS.md for the step this PR closes. -->

## Notes

<!-- Related issues/links/threads. Use "Fixes #N" / "Refs #N". -->

## Screenshots / logs

<!-- Paste a screenshot of the taskpane change, a Playwright trace link, or a CI failure that this PR fixes. -->
