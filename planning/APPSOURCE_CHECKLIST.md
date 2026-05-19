# AppSource Submission Checklist

Driven by Step 4.5 in [`STEPS.md`](STEPS.md). Mirrors Microsoft's [certification policy 1100 — Office Add-ins](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies) and the [publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).

## Manifest

- [ ] `Id` — stable GUID, distinct per environment (dev / staging / prod).
- [ ] `Version` — four-part `a.b.c.d`, matches the value in Partner Center.
- [ ] `DisplayName` — exactly `Synamail`.
- [ ] Brand name pre-validated against Microsoft's name-collision check.
- [ ] `Permissions` — least privilege (`ReadWriteItem`).
- [ ] `Hosts` — `Mailbox` only.
- [ ] `Requirements` — `Mailbox` minVersion 1.8.
- [ ] `AppDomains` includes every external domain the taskpane loads (Synaplan API host **and** the Synaplan login origin).
- [ ] Every URL in the manifest is HTTPS.
- [ ] `office-addin-validator manifest.xml` → 0 warnings.
- [ ] `office-addin-validator manifest.unified.json` → 0 warnings.

## Hosting & security

- [ ] Taskpane SPA hosted on a stable HTTPS origin (`addin.synaplan.com` or `web.synaplan.com/addin/`).
- [ ] Strict CSP active on the host.
- [ ] CORS allow-list on the Synaplan API includes the production origin.
- [ ] No third-party scripts loaded inside the taskpane.
- [ ] API key never logged; stored only in `Office.context.roamingSettings`.
- [ ] 401 from API clears the key and bounces to SignIn.

## Functionality across clients

- [ ] Works in Outlook on the Web (Edge + Chrome verified).
- [ ] Works in new Outlook for Windows.
- [ ] Works in classic Outlook 2024 on Windows.
- [ ] Works in Outlook on Mac (new + classic).
- [ ] Sign-in dialog flow works end-to-end against the demo tenant supplied to reviewers.
- [ ] Every advertised feature in the store listing actually works.

## Legal / privacy

- [ ] Privacy policy URL publicly reachable (no auth wall).
- [ ] Terms of Service URL publicly reachable.
- [ ] Support URL publicly reachable.
- [ ] Telemetry is **opt-in only**; no PII leaves the client without user action.
- [ ] Body and attachments leave Outlook only on explicit user action.

## Localisation

- [ ] All UI strings localised to en + de minimum (matches Synaplan policy).
- [ ] AppSource listing copy localised to en + de.

## Graphics

- [ ] Icons at 16, 32, 64, 80, 128 PNG present and on-brand.
- [ ] Hero at 256 and 512 PNG present and on-brand.
- [ ] At least three screenshots at 1366×768 PNG, showing real content.
- [ ] Screencast video ≤ 60 s, 1080p, narrated or captioned.

## Reviewer enablement

- [ ] Demo Microsoft 365 tenant credentials prepared.
- [ ] Demo Synaplan account pre-seeded with a RAG group + a few documents.
- [ ] Sign-in flow walkthrough included in the reviewer notes.

## Submission

- [ ] Partner Center account in "Microsoft 365 and Copilot" program.
- [ ] All Partner Center fields filled (privacy, terms, support, descriptions).
- [ ] First submission filed.
- [ ] Review comments tracked as issues (if any).
- [ ] Resubmission filed (if needed).
- [ ] Listing live in AppSource.
