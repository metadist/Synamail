<!-- markdownlint-disable MD034 -- plain URLs are intentional: this text is copy-pasted as-is into Partner Center -->

# Synamail — Notes for certification (reviewer instructions)

Paste this into the Partner Center **"Notes for certification"** field, and
**replace every bracketed `[…]` value with a real, working credential before
submitting.** A submission left with the placeholder values below cannot be
signed in to and will fail certification.

---

## Test account (REQUIRED — fill these in)

Synamail connects to a Synaplan workspace. We have pre-seeded a reviewer
account; the API key it stores is per-mailbox and revocable at any time.

- **Synaplan sign-in:** [reviewer@example.com] / [password]
- **Server URL:** https://web.synaplan.com (this is the default — do **not**
  change the "Synaplan instance" field)
- **Microsoft 365 test mailbox** (to install/sideload the add-in):
  [m365-user@tenant] / [password]

The reviewer account already has the `synamail` Synaplan companion installed, so
the **Contact AI Profiling** feature is fully functional.

---

## Where to find the add-in (it appears on the message surface, not the main toolbar)

The Synamail button is a **per-message** command — it only appears once a mail
item is open or selected, never on the empty inbox/folder toolbar. By client:

- **Outlook on the web / new Outlook for Windows:** open any email. The
  **Synamail** group with the **Open Synamail** button appears on the message
  toolbar. If it is collapsed, click **Apps** (or the **…** / "More apps"
  overflow) and choose **Synamail**.
- **Classic Outlook for Windows (Microsoft 365):** open an email in the Reading
  Pane or in its own window → **Home** tab (Read) → the **Synamail** group →
  **Open Synamail**. In a compose/reply window it is on the **Message** tab.
- **Outlook for Mac:** open an email → the **Synamail** button is on the message
  ribbon; if hidden, use the **Apps** / **…** overflow on the message toolbar.

Clicking **Open Synamail** opens the task pane on the right. On first open it
shows the **Sign in to Synaplan** screen.

## How to test (about 5 minutes)

1. Open any email and click **Open Synamail** (see locations above). The task
   pane opens.
2. Click **Sign in to Synaplan**. A Microsoft dialog opens the Synaplan login
   page; sign in with the reviewer account above and click **Connect this
   Outlook**. The dialog closes and the task pane shows the signed-in home view.
3. With an email open, expand **Email actions** and try:
   - **Summarize** — produces a short summary of the email.
   - **Translate** — translates the email body.
   - **Draft reply** — generates a reply draft.
   - **Save to knowledge base** — stores the email in the Synaplan workspace.
   - **Ask** — type a question about the email (attachments are analysed too).
4. Open **Contact AI Profiling** for the sender, then **Update profile from this
   email** — a rolling AI profile is created. It can be deleted again from the
   same panel.

## Environments verified by us

- Outlook on the web (Edge/Chrome, Windows + Mac)
- New Outlook for Windows
- Classic Outlook for Microsoft 365 on Windows 11 (Edge WebView2 runtime)
- Outlook for Mac (Microsoft 365)

Office.js is loaded from the Microsoft-hosted CDN
(https://appsforoffice.microsoft.com/lib/1/hosted/office.js) in every add-in
HTML page. All add-in resources are served over HTTPS from
https://addin.synaplan.com.

## Privacy / data handling (for the reviewer)

- Synamail acts only on explicit user clicks; it does not read the mailbox in the
  background and does not send mail on its own.
- Email content is sent only to the user's own Synaplan workspace
  (web.synaplan.com), hosted in Germany.
- The stored API key is per-mailbox, encrypted at rest by Exchange, and revocable
  from the Synaplan web app (API Keys page).
- Contact AI Profiling is optional, user-triggered, stored only in the user's
  workspace, and user-deletable.
- Privacy policy: https://www.synaplan.com/privacy-policy
- Source code (Apache-2.0): https://github.com/metadist/Synamail
