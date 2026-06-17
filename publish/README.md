# Synamail — AppSource publishing pack

Ready-to-paste copy and metadata for the Microsoft Partner Center
("Marketplace listings") submission of the **Synamail** Outlook add-in.

## Files

| File                                           | What it's for                                          |
| ---------------------------------------------- | ------------------------------------------------------ |
| [`listing-en.md`](listing-en.md)               | English listing: Name, Summary, Description, Keywords  |
| [`listing-de.md`](listing-de.md)               | German listing (add a `de` language to the offer)      |
| [`listing-es.md`](listing-es.md)               | Spanish listing (add an `es` language to the offer)    |
| [`listing-tr.md`](listing-tr.md)               | Turkish listing (add a `tr` language to the offer)     |
| [`metadata-and-urls.md`](metadata-and-urls.md) | IDs, version, publisher, and all required URLs         |
| [`reviewer-notes.md`](reviewer-notes.md)       | "Notes for certification" — test account + how to test |

## Partner Center field map (English listing page)

| Partner Center field | Source                                       |
| -------------------- | -------------------------------------------- |
| **Name**             | `listing-en.md` → Name                       |
| **Summary**          | `listing-en.md` → Summary (≤ 100 chars)      |
| **Description**      | `listing-en.md` → Description (HTML allowed) |
| **Search keywords**  | `listing-en.md` → Search keywords (max 3)    |
| Privacy policy URL   | `metadata-and-urls.md`                       |
| Support / Help URL   | `metadata-and-urls.md`                       |
| Package (manifest)   | `../manifest.prod.xml`                       |

## Before you submit — checklist

- [ ] Upload `../manifest.prod.xml` on the **Packages** page.
- [ ] Paste the English listing (and German, if adding `de`).
- [ ] Set Privacy, Terms and Support URLs (see `metadata-and-urls.md`).
- [ ] **Deploy the updated privacy policy first** — it must publicly mention the
      platform data handling + Contact AI Profiling (the `synaplan-website`
      `/privacy-policy` change). Reviewers check this.
- [ ] Upload screenshots + logo — already in this folder: `300x300logo.jpg`,
      `synamail_screen1366x768_1.jpg`, `synamail_screen1366x768_2.jpg`.
- [ ] Fill **Notes for certification** from `reviewer-notes.md` (test account!).
- [ ] Confirm the Partner Center **publisher display name** matches the manifest
      `ProviderName` (`Synaplan`).
