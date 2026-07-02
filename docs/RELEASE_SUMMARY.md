# Synamail v1 — Release Summary

_Executive summary, 2026-06-10. Details: [`FEATURES.md`](FEATURES.md) §6 (scope),
[`CONTACT_PROFILING.md`](CONTACT_PROFILING.md) (profiling),
[`SYNAPLAN_INTEGRATION.md`](SYNAPLAN_INTEGRATION.md) §2.5 (plugin release flow)._

## What v1 is

Synamail brings the user's Synaplan workspace into Outlook. The v1 release was
deliberately **condensed to what demonstrably works on every required Outlook
host** (OWA, new Outlook for Windows, classic 2024, Mac):

- **Home chat** with Synaplan, streaming replies.
- **Email actions**: summarize, translate, draft reply, classify, find meeting
  times, ask follow-ups (with image vision), save to knowledge base.
- **Contact AI Profiling** — the headline new feature (below).
- Sign-in via the Office dialog bridge, self-hosted instances, settings.

**Cut from v1** (designs retained in `docs/`): compose-mode assistance, Mail
Routes, the RULE/Synapse editor, and every EWS-dependent feature (mailbox
search, sender history, block sender) — EWS retires for Exchange Online in
Oct 2026 and is unavailable in new Outlook for Windows. Consequence: the
manifest permission dropped from `ReadWriteMailbox` to **`ReadWriteItem`**,
which also simplifies AppSource review.

## Contact AI Profiling (new)

A **rolling profile per mailing partner**: a growing AI-merged summary (who the
person is, relationship tone, stable facts, open loops like _"they owe you the
contract draft"_), updated one email at a time from Outlook.

- **The prompt and storage ship as a separate Synaplan plugin** —
  `synamail-plugin/` in this repo, released to `synaplan/plugins/synamail/`
  (same model as Synaform: develop in the product repo, release to `plugins/`
  for customers who want it). No Synaplan core changes.
- Deterministic facts (email count, first/last seen, org) are computed in code;
  the AI only writes the narrative. Profiles are user-owned, listed in a
  Synaplan web panel, and deletable from both sides (GDPR).
- Add-ins on instances **without** the plugin degrade gracefully (install hint).

## Release flow (two repos — keep them honest)

1. Develop in `Synamail/` (add-in + `synamail-plugin/`).
2. `make sync-plugin` copies the plugin to `synaplan/plugins/synamail/` and
   **prints the git drift in the synaplan repo** — those files are tracked
   there, so always commit them as a **dedicated `feat(plugins)`/`fix(plugins)`
   commit**, never mixed into unrelated synaplan work.
3. Install per user: `php bin/console app:plugin:install <userId> synamail`.
4. Add-in store packaging: `manifest_1-0-2_prod.xml` (production GUID, hosted on
   `addin.synaplan.com`) → `make build-manifest` for the AppSource submission.

## Verification status

| Check                                                                    | Result      |
| ------------------------------------------------------------------------ | ----------- |
| `make ci-local` (lint, types, 90 unit/component tests, manifests, build) | ✅ green    |
| Playwright E2E (12 specs incl. rolling-profile flow)                     | ✅ 12/12    |
| Plugin live API (get/update×2-merge/list/delete, 401/400 guards)         | ✅ verified |
| Plugin PSR-12 + route registration + per-user install                    | ✅ verified |

## Open before submission

Publish `dist/` + `assets/` to `addin.synaplan.com`; public privacy/ToS pages
(must mention profiling); reviewer demo tenant with the plugin pre-installed;
store listing assets (en/de). Checklist: [`PROJECT_PLAN.md`](PROJECT_PLAN.md).
