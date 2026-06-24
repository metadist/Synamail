# Synamail — Project Plan & Reference

This document serves as the consolidated planning reference for the Synamail Outlook add-in project.

## Development Status

Synamail development is organized into 4 sprints:

- **Sprint 1 — Planning, Design & Test Definition**: Done.
- **Sprint 2 — GUI, Local Sideload, Live Identification**: Done.
- **Sprint 3 — Functionality & Live API Calls**: Done for the v1 scope (read-mode actions; compose-mode deferred).
- **Sprint 4 — Release (AppSource & Open Source)**: Current focus. The app was
  **condensed for release on 2026-06-10**: compose mode, Mail Routes, the rule
  editor, and all EWS-dependent features were cut (see `FEATURES.md` §6), and
  **Contact AI Profiling (rolling profiles)** shipped via the new `synamail`
  Synaplan plugin (`synamail-plugin/`, released with `make sync-plugin`).

## GUI Definitions & Assets

### Routing (v1)

- No API key: `SignIn.vue`
- Signed in: `Home.vue` (chat + Email actions accordion)
- Navigation: `Settings.vue`, `ContactProfile.vue`

### Visual Language

- **Theme**: Fluent UI design tokens. Light + Dark modes driven by Office UI.
- **Typography**: Segoe UI Variable / system-ui.
- **Iconography**: Outline-style 16/20 px icons (Fluent System).

### AppSource Asset Checklist

- Icons (PNG): 16, 32, 64, 80, 128
- Hero images: 256x256, 512x512
- Screenshots (1366x768): SignIn, Home (chat + Email actions), Settings, ContactProfile (profile card)
- Screencast (1080p, ≤60s)
- Store copy (`assets/store/copy.md` en + de)

## AppSource Submission Checklist

Mirrors Microsoft's certification policy 1100.

- [x] Manifest `Id`, `Version`, `DisplayName` = `Synamail` — `manifest.prod.xml`
      carries the production GUID (`342cee66-…`), v1.0.0.1, host
      `addin.synaplan.com`, and the narrowed `ReadWriteItem` permission.
- [x] `make validate` (office-addin-manifest) green for `manifest.xml` AND
      `manifest.prod.xml`; `make build-manifest` converts the prod manifest to
      `manifest.unified.json` for submission.
- [ ] Hosted on stable HTTPS: publish `dist/` + `assets/` to
      `https://addin.synaplan.com` (paths are baked into `manifest.prod.xml`).
- [ ] Privacy, Terms of Service, and Support URLs publicly reachable. The
      privacy policy MUST mention Contact AI Profiling (profiles of email
      contacts, stored only in the user's own Synaplan workspace, fully
      user-deletable) — reviewers look for this.
- [ ] Functional across OWA, new Outlook (Windows), classic Outlook 2024, Mac.
      (v1 deliberately uses no EWS, so no feature is host-dependent.)
- [ ] Demo Microsoft 365 tenant credentials + Synaplan account pre-seeded for
      reviewers — install the `synamail` plugin for the demo user so the
      profile card works (`php bin/console app:plugin:install <userId> synamail`).
- [ ] Store listing + screenshots + copy in en/de uploaded.

## Open Source Launch

- Clean git history (no secrets).
- Ensure `SECURITY.md`, `CODE_OF_CONDUCT.md`, `CHANGELOG.md` are up to date.
- Publish `v1.0.0` release on GitHub (`metadist/Synamail`).
- Make the repository public at/near AppSource approval.
