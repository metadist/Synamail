# Changelog

All notable changes to Synamail are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Master release & launch plan (`docs/PROJECT_PLAN.md`) covering the path from the
  current beta to a dual v1.0 launch on Microsoft AppSource and as open source.
- End-user guide (`docs/USER_GUIDE.md`) — a plain-language, screenshot-led walkthrough.
- Open-source project files: `SECURITY.md`, `CODE_OF_CONDUCT.md`, this changelog.
- Two-lane README front door for the "use it" vs. "build / host it" audiences.

### Changed

- Decisions locked for v1.0: Mail Routes ships as **Preview**, RULE integration is
  deferred to **v1.1**, the **parrot** is the confirmed brand mark, and the public repo
  goes live at/just before AppSource acceptance.

## [1.0.0] - TBD

First public release.

- Live sign-in to Synaplan via the Office Dialog OAuth bridge.
- Read mode: summarise, translate, draft reply, classify, ask follow-ups.
- Save emails and attachments to a Synaplan knowledge-base group.
- Contact AI Profiling: per-contact knowledge group — save and search by sender / recipient.
- Compose mode: draft from prompt, improve / shorten / translate selection, insert
  from knowledge base.
- Mail Routes (Preview).
- English and German localization.

[Unreleased]: https://github.com/metadist/Synamail/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/metadist/Synamail/releases/tag/v1.0.0
