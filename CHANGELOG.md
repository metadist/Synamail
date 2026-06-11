# Changelog

All notable changes to Synamail are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Contact AI Profiling — rolling profiles (Phase 1).** A growing, AI-merged
  summary of every mailing partner: summary, tone, facts, and open loops,
  updated one email at a time. The prompt and profile storage ship as a
  separate **Synaplan plugin** (`synamail-plugin/`, released to
  `synaplan/plugins/synamail/` via `make sync-plugin`); the add-in renders the
  profile card in `ContactProfile.vue` and rolls emails in explicitly or on
  save-to-contact-group. Includes a Synaplan-side panel to list/delete all
  profiles (privacy/GDPR surface) and full per-contact delete from Outlook.
- Production manifest `manifest.prod.xml` (production GUID, hosted URLs on
  `addin.synaplan.com`) — `make validate` checks both manifests and
  `make build-manifest` converts the production one for the store submission.
- Master release & launch plan (`docs/PROJECT_PLAN.md`) covering the path from the
  current beta to a dual v1.0 launch on Microsoft AppSource and as open source.
- End-user guide (`docs/USER_GUIDE.md`) — a plain-language, screenshot-led walkthrough.
- Open-source project files: `SECURITY.md`, `CODE_OF_CONDUCT.md`, this changelog.
- Two-lane README front door for the "use it" vs. "build / host it" audiences.

### Changed

- **Condensed the app to the working v1 feature set for the AppSource release.**
  Everything that was partial, unreachable, or EWS-dependent was removed:
  compose mode (orphaned view, wrong prompts), Mail Routes (config UI without a
  runtime), the RULE/Synapse editor (never built), "More from this sender",
  block-sender, and the mailbox→knowledge-base filter (all EWS-based with mock
  fallbacks; EWS retires for Exchange Online in Oct 2026). Their designs remain
  in `docs/` for later releases.
- Manifest permission narrowed from `ReadWriteMailbox` to `ReadWriteItem` — v1
  uses no EWS, which also simplifies AppSource review.
- Dead i18n removed and the new profiling strings added across all six shipped
  locales (en, de, fr, es, it, pt).

## [1.0.0] - TBD

First public release.

- Live sign-in to Synaplan via the Office Dialog OAuth bridge.
- Home chat with the user's Synaplan workspace.
- Read mode: summarise, translate, draft reply, classify, find meeting times,
  ask follow-ups (with email-image vision).
- Save emails and attachments to a Synaplan knowledge-base group.
- Contact AI Profiling: per-contact knowledge group plus a rolling AI profile
  (summary / tone / facts / open loops) served by the `synamail` Synaplan plugin.
- Localization in en, de, fr, es, it, pt.

[Unreleased]: https://github.com/metadist/Synamail/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/metadist/Synamail/releases/tag/v1.0.0
