# Synamail — Project Plan & Reference

This document serves as the consolidated planning reference for the Synamail Outlook add-in project.

## Development Status

Synamail development is organized into 4 sprints:

- **Sprint 1 — Planning, Design & Test Definition**: Done.
- **Sprint 2 — GUI, Local Sideload, Live Identification**: Done.
- **Sprint 3 — Functionality & Live API Calls**: Done. All read and compose actions live.
- **Sprint 4 — Release (AppSource & Open Source)**: Current focus.

## GUI Definitions & Assets

### Routing

- No API key: `SignIn.vue`
- API key + message selected: `ReadMode.vue`
- API key + composing: `ComposeMode.vue`
- Navigation: `Settings.vue`, `RuleEditor.vue`, `ContactProfile.vue`

### Visual Language

- **Theme**: Fluent UI design tokens. Light + Dark modes driven by Office UI.
- **Typography**: Segoe UI Variable / system-ui.
- **Iconography**: Outline-style 16/20 px icons (Fluent System).

### AppSource Asset Checklist

- Icons (PNG): 16, 32, 64, 80, 128
- Hero images: 256x256, 512x512
- Screenshots (1366x768): SignIn, ReadMode, ComposeMode, Settings, ContactProfile, RuleEditor
- Screencast (1080p, ≤60s)
- Store copy (`assets/store/copy.md` en + de)

## AppSource Submission Checklist

Mirrors Microsoft's certification policy 1100.

- [ ] Manifest `Id`, `Version`, `DisplayName` = `Synamail`.
- [ ] `office-addin-validator manifest.xml` / `manifest.unified.json` → 0 warnings.
- [ ] Hosted on stable HTTPS (`web.synaplan.com/addin/` or `addin.synaplan.com`).
- [ ] Privacy, Terms of Service, and Support URLs publicly reachable.
- [ ] Functional across OWA, new Outlook (Windows), classic Outlook 2024, Mac.
- [ ] Demo Microsoft 365 tenant credentials + Synaplan account pre-seeded for reviewers.
- [ ] Store listing + screenshots + copy in en/de uploaded.

## Open Source Launch

- Clean git history (no secrets).
- Ensure `SECURITY.md`, `CODE_OF_CONDUCT.md`, `CHANGELOG.md` are up to date.
- Publish `v1.0.0` release on GitHub (`metadist/Synamail`).
- Make the repository public at/near AppSource approval.
