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

- [ ] **Upload the PRODUCTION manifest: `../manifest.prod.xml`** on the
      **Packages** page (URLs on `https://addin.synaplan.com`, GUID
      `342cee66-dd27-471d-b2ac-fa1bbb5db54f`).
      **⚠️ NEVER upload `../manifest.xml`** — that is the local **dev** manifest
      (`https://localhost:3000`). Uploading it fails certification with
      "add-in does not load in any environment" on Windows, Mac, and web,
      because Microsoft's reviewers have nothing serving `localhost:3000`.
      Verify the file you upload contains `addin.synaplan.com`, **not**
      `localhost`, before submitting.
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

## Certification remediation — 2026-06-19 report (Product ID f21612a9-…)

The first submission came back "Attention needed". Re-check each item below
before resubmitting (Partner Center paths in brackets):

- [ ] **1120.3.1.1 / 1120.3.7.3 / 1120.3.7.8 — "does not load" (Win/Mac/WebView2).**
      Root cause: wrong/non-loading package. Fix = upload `../manifest.prod.xml`
      (see warning above) and confirm the host is live:
      `curl -sI https://addin.synaplan.com/src/taskpane/taskpane.html` → `200`.
      Sideload `manifest.prod.xml` yourself first (Outlook → Get Add-ins →
      My add-ins → Custom add-ins → **Add from file**) and confirm the pane opens.
- [ ] **100.3.3.3 — Video URL** [Offer Listing → Videos]. Use the **canonical**
      YouTube page URL, not a short link:
      `https://www.youtube.com/watch?v=h9Ouzl4AZ1E` (NOT `https://youtu.be/…`).
- [ ] **1120.2.3.1 — Apple ID / iOS** [Availability / Markets]. The manifest has
      **no mobile form factor** (`make validate` says "does not include mobile
      apps"), so set **Apple / iOS store availability = NO**. (Only set it YES if
      you add a `MobileFormFactor` and supply a valid Apple ID.)
- [ ] **100.3.2.2 — Screenshots** [Offer Listing → Screenshots]. Add at least one
      **compose-mode** screenshot (Synamail open in a new-mail / reply window).
      Compose mode is supported by the manifest (`MessageComposeCommandSurface`)
      and the taskpane, so this is just a missing image.
