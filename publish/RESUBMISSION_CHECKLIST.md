<!-- markdownlint-disable MD013 MD034 -->

# Synamail — AppSource resubmission checklist

Tracks every point from the Microsoft certification report and the action taken.
Work top to bottom; everything under **Before you resubmit** must be ticked.

## Certification report — point by point

| Policy                                             | Finding                                                 | Root cause                                                                                         | Action taken                                                                                                                                                                                                                                         | Owner action before resubmit                                                                                                                                                                                                 |
| -------------------------------------------------- | ------------------------------------------------------- | -------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **100.1.3.1** Description — Offer Length & Quality | Description references "Plugin"                         | The store description called the Synaplan companion a "plugin"                                     | Replaced "plugin/Plugin/complemento/eklenti" with "companion / Komponente / componente / bileşeni" in `listing-en/de/es/tr.md` and `metadata-and-urls.md`                                                                                            | Re-paste the updated description into Partner Center for **every** language (EN, DE, ES, TR). Confirm the word "plugin" appears nowhere in any listing field.                                                                |
| **100.3.2.2** Content & Quality                    | Offer images contain personal information               | Screenshots show a real inbox / contact data                                                       | (code: n/a)                                                                                                                                                                                                                                          | **Re-shoot or mask** all screenshots in `assets/store/` and `publish/synamail_screen*.{png,jpg}` so no real names, email addresses, or message content are visible. Add at least one **compose-mode** screenshot. Re-upload. |
| **1120.1.1** Office JavaScript APIs                | "Office.js is not able to check" / not Microsoft-hosted | Almost certainly tested against an older/broken build; current build is correct                    | All three add-in pages (`taskpane.html`, `commands.html`, `auth-relay.html`) load `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` as a direct `<script>` in `<head>`. Added an `onerror` hook so a CDN failure is surfaced, not silent. | Confirm the **live** pages serve this exact script (see verification commands below) **before** uploading the manifest.                                                                                                      |
| **1120.3.1.1** Slow Load Time                      | "Could not install or load your add-in"                 | Blank white pane until the JS bundle parsed read as "not loading"; possible cold-cache/CDN latency | `taskpane.html` now shows an immediate branded **loading spinner**, a **noscript** fallback, and — after a 12s grace period — an actionable **error message with a support link** instead of a blank/stuck pane                                      | Redeploy `addin.synaplan.com` with this build, then re-open the add-in on a cold profile to confirm a visible loading state.                                                                                                 |
| **1120.3.7.3** Office for Mac                      | Add-in not visible on Mac                               | Reviewer likely couldn't locate the per-message button                                             | No manifest change needed (Desktop form factor covers Mac). Reviewer notes now say exactly where the button is per client.                                                                                                                           | Re-test on Outlook for Mac (M365). Confirm **Open Synamail** is reachable from the message ribbon / Apps overflow.                                                                                                           |
| **1120.3.7.8** O365 Perpetual — WebView2 / Anaheim | Add-in not visible on Windows 11                        | Same as Mac: button location + possibly stale build                                                | No manifest change needed. Reviewer notes cover classic Outlook + WebView2.                                                                                                                                                                          | Re-test on classic Outlook for M365 on Windows 11 with Edge WebView2. Confirm the button + that the pane loads.                                                                                                              |

## Before you resubmit (hard gate)

- [ ] **Redeploy** `https://addin.synaplan.com` from this commit (the loading-state + Office.js `onerror` hardening must be live). Verify with `version.json`.
- [ ] **Reviewer notes:** open `publish/reviewer-notes.md` and replace every
      `[…]` placeholder with **real, working** credentials. This was a likely
      cause of "could not load" — a reviewer with placeholder credentials can't
      sign in. Paste into Partner Center → Notes for certification.
- [ ] **Descriptions:** re-paste EN/DE/ES/TR descriptions; grep yourself — no
      "plugin".
- [ ] **Screenshots:** all personal information masked; compose-mode shot added.
- [ ] **Manifest:** upload `manifest.prod.xml` (GUID
      `342cee66-dd27-471d-b2ac-fa1bbb5db54f`). Bump `<Version>` if Partner Center
      requires a higher build than the previous submission.
- [ ] **Cloudflare:** confirm `addin.synaplan.com` is **not** behind a Bot Fight
      Mode / JS-challenge rule. The Office WebView runtime must reach the pages
      and the office.js CDN without a challenge (an interactive challenge renders
      as a blank pane = "not loading"). Add a WAF skip/bypass for the host if in
      doubt.
- [ ] **Sign-in page:** confirm `https://web.synaplan.com/addin/connect` is live
      (HTTP 200) — it is the target of the sign-in dialog.

## Live verification commands (run before uploading)

```bash
# Pages serve 200 + the Microsoft-hosted office.js
curl -sS https://addin.synaplan.com/src/taskpane/taskpane.html | grep -o 'appsforoffice.microsoft.com/lib/1/hosted/office.js'
curl -sS -o /dev/null -w '%{http_code}\n' https://addin.synaplan.com/src/commands/commands.html
curl -sS -o /dev/null -w '%{http_code}\n' https://addin.synaplan.com/src/dialog/auth-relay.html

# Manifest icons (all must be 200)
for n in 16 32 64 80 128; do curl -sS -o /dev/null -w "icon-$n: %{http_code}\n" https://addin.synaplan.com/assets/icon-$n.png; done

# Sign-in dialog target is live
curl -sS -o /dev/null -w '%{http_code}\n' https://web.synaplan.com/addin/connect

# What build is live
curl -sS https://addin.synaplan.com/version.json
```

## Local gate (must be green before deploy)

```bash
make ci-local   # lint-docs + lint + check-types + test + validate + build + budget
```
