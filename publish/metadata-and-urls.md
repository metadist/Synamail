# Synamail — submission metadata & URLs

## Product / package

| Field            | Value                                                |
| ---------------- | ---------------------------------------------------- |
| Product name     | Synamail                                             |
| Publisher        | Synaplan (metadist data management GmbH)             |
| Manifest         | `../manifest.prod.xml` (upload this on **Packages**) |
| Add-in ID (GUID) | `342cee66-dd27-471d-b2ac-fa1bbb5db54f`               |
| Version          | `1.0.0.0`                                            |
| Manifest type    | Add-in only manifest (XML), schema 1.1               |
| Permission       | `ReadWriteItem`                                      |
| Hosting          | `https://addin.synaplan.com` (static, HTTPS, live)   |
| Min requirement  | Mailbox requirement set 1.8                          |

> The Partner Center **publisher display name must match** the manifest
> `<ProviderName>` value: **Synaplan**.

## Required URLs

| Field              | URL                                       | Status   |
| ------------------ | ----------------------------------------- | -------- |
| Support URL        | `https://web.synaplan.com/support`        | live 200 |
| Privacy policy URL | `https://www.synaplan.com/privacy-policy` | see note |
| Terms of use URL   | `https://web.synaplan.com/terms`          | live 200 |
| Help / Learn more  | `https://github.com/metadist/Synamail`    | live     |

> **Privacy policy note:** use the marketing-site page
> `https://www.synaplan.com/privacy-policy`, which now covers the platform data
> handling, the Synamail plugin connection, Contact AI Profiling, and the
> Apache-2.0 source availability. **Deploy that update before submitting** —
> the live page must contain those sections (reviewers verify it).
> (`https://web.synaplan.com/privacy` is the platform's own page and also works,
> but the marketing page is the one written for this listing.)

## Supported clients (for the listing)

- Outlook on the web (Windows, Mac, and Linux)
- New Outlook for Windows
- Classic Outlook 2024 (Windows)
- Outlook on Mac

## Assets

Already in this `publish/` folder:

- Logo (300×300): `300x300logo.jpg`
- Screenshots (1366×768): `synamail_screen1366x768_1.jpg`, `synamail_screen1366x768_2.jpg`
- Extra branding: `big_logo.png`, `bird.png`
- Add-in icons + hero: `../assets/icon-*.png`, `../assets/store/hero-*.png`

Optional / nice to have:

- Up to 5 screenshots total (add a couple more, e.g. Settings, Contact Profile)
- Video: YouTube/Vimeo link + 1280×720 PNG thumbnail

> Note: the `*.jpg:com.dropbox.attrs` files are Dropbox sync sidecars — harmless,
> ignore them (don't upload them).
