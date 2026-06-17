<!-- markdownlint-disable MD034 -- plain URLs are intentional: this text is copy-pasted as-is into Partner Center -->

# Synamail — Notes for certification (reviewer instructions)

Paste this into the Partner Center **"Notes for certification"** field, and fill
in the bracketed test credentials before submitting.

---

## Test account

Synamail requires a Synaplan workspace. We have pre-seeded a reviewer account:

- **Synaplan sign-in:** [reviewer@example.com] / [password]
- **Server URL:** https://web.synaplan.com (this is the default; no change needed)
- A Microsoft 365 test mailbox to sideload the add-in: [m365-user@tenant] / [password]

The reviewer account already has the `synamail` server plugin installed, so the
Contact AI Profiling feature is fully functional.

## How to test (5 minutes)

1. In Outlook, open any email and click **Open Synamail** in the ribbon.
2. Click **Sign in to Synaplan**. A login window opens — sign in with the test
   account above and click **Connect this Outlook**. The window closes.
3. With an email open, try the actions:
   - **Summarize** — produces a short summary of the email.
   - **Translate** — translates the email body.
   - **Draft reply** — generates a reply draft.
   - **Save to knowledge base** — stores the email in the Synaplan workspace.
   - **Ask** — type a question about the email (attachments are analysed too).
4. Open **Contact AI Profiling** for the sender, then **Update profile from this
   email** — a rolling AI profile is created. It can be deleted again from the
   same panel.

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
