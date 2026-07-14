# Synamail — User Guide

**Synamail puts your Synaplan AI assistant right inside Outlook.** Summarise long
emails, translate them, draft replies, save them to your knowledge base, and ask
questions about a thread — without ever leaving your inbox.

This guide is for **everyday users**. You don't need any technical knowledge, and you
won't touch a command line. If you're a developer who wants to build or self-host
Synamail, see [`../INSTALL.md`](../INSTALL.md) instead.

> Each step below ends with one **Next →** so you always know what to do next.

---

## 1. What is Synamail?

It's a small panel ("add-in") that appears next to your emails in Outlook. You click a
button — _Summarise_, _Translate_, _Draft reply_ — and your Synaplan workspace does the
work. Nothing happens to your mail automatically: **Synamail only acts when you click.**

**Next →** [Install it](#2-install-synamail)

---

## 2. Install Synamail

Synamail works in Outlook on the Web (any browser, on Windows, Mac, **and Linux**), new
Outlook for Windows, classic Outlook 2024, and Outlook on Mac.

1. Open your Outlook and open any email.
2. Click the **Apps** icon (four little squares) in the email's toolbar.
   _(Older layouts: **… → Get Add-ins**.)_
3. Search for **Synamail** and click **Add**.

| Your Outlook                     | Where to find "Apps / Get Add-ins"    |
| -------------------------------- | ------------------------------------- |
| Outlook on the Web (incl. Linux) | Open an email → toolbar **Apps** icon |
| New Outlook for Windows          | **Home → Get Add-ins**                |
| Classic Outlook for Windows      | **Home → Get Add-ins**                |
| Outlook on Mac                   | Toolbar **Apps** icon                 |

> **On Linux?** Use Outlook in your browser. You can even click your browser's
> "Install app" to keep Outlook (and Synamail) as a desktop-like app. Your mailbox must
> be a Microsoft 365 or Outlook.com account.
>
> **Don't see an "Add custom add-in" option at work?** Your IT team may restrict
> add-ins. The certified AppSource version is usually allowed — ask them to enable
> **Synamail** in the Microsoft 365 admin center if needed.

**Next →** [Connect your account](#3-connect-your-synaplan-account)

---

## 3. Connect your Synaplan account

1. Open any email and click **Open Synamail** in the ribbon.
2. The panel shows a **Sign in to Synaplan** button. Click it.
3. A small Synaplan login window opens — sign in the way you normally do (Google,
   Microsoft, GitHub, or password).
4. Click **Connect this Outlook**. The window closes — you're ready.

You never type or paste a key. If your company runs its own Synaplan server, click
**"Use a self-hosted instance"** _before_ signing in and enter your server address.

**Next →** [Summarise your first email](#4-summarise-an-email)

---

## 4. Summarise an email

The fastest way to feel the magic:

1. Open a long email.
2. In the Synamail panel, click **Summarise**.
3. Read the short bullet-point summary. Click **Copy** to keep it.

**Next →** [Translate, reply, classify](#5-translate-reply-and-classify)

---

## 5. Translate, reply, and classify

- **Translate** — click **Translate**, pick a language. The email is shown in your
  language.
- **Draft reply** — click **Draft reply**, choose a tone (formal / concise / friendly).
  A pre-written reply opens in Outlook's compose window for you to review and send.
- **Classify** — click **Classify** to auto-tag the email (invoice, support, personal…)
  for your own triage.

**Next →** [Save an email to your knowledge base](#6-save-an-email-to-your-knowledge-base)

---

## 6. Save an email to your knowledge base

Your **knowledge base** is your personal, searchable memory in Synaplan. Save an email
so the AI can recall it later.

1. Open an email and click **Save to knowledge base**.
2. Pick a group (or create one), choose which attachments to include.
3. Click **Save**. Done — it's now searchable in Synaplan.

**Next →** [Build a knowledge base for a contact](#7-build-a-knowledge-base-per-contact)

---

## 7. Build a knowledge base per contact

Keep everything about one person in one place.

1. On an email, click the **Contact** pill below the subject (e.g. `alice@example.com`).
2. Click **Save current email to this contact** to add it.
3. Use the **Search** box to find anything you've saved about that contact, or click
   **Ask about this contact** to chat with the AI using only their history.

**Next →** [Write better emails](#8-write-better-emails-compose-mode)

---

## 8. Write better emails (Compose mode)

When you're writing a new email or reply, open Synamail to:

- **Draft from prompt** — type a one-line idea, get a full draft.
- **Improve / Shorten / Translate** — select some text, then click to rewrite it.
- **Insert from knowledge base** — search your saved knowledge and drop a snippet
  (with a citation) into your email.

**Next →** [Settings](#9-settings-you-might-want)

---

## 9. Settings you might want

Open **⚙ Settings** in the panel:

- See which account and server you're signed in to.
- Change the **language** of the interface.
- Switch to a **self-hosted** Synaplan (sign out first, then sign back in).
- **Log out** if you ever want a clean start.

> **Mail Routes (Preview):** you may see a "Mail Routes" tab marked _Preview_. It's an
> early feature for automating actions on incoming mail. It only ever acts when you
> set it up and confirm — feel free to explore, but expect rough edges.

**Next →** [Privacy](#10-privacy-in-plain-words)

---

## 10. Privacy in plain words

- **Nothing leaves Outlook until you click.** Synamail never reads or sends your mail
  in the background.
- An email's content reaches your Synaplan workspace **only** for the action you click
  (e.g. when you press _Summarise_).
- Your connection key is stored securely in your mailbox and can be removed any time
  from Synamail's **Sign out**, or from your Synaplan account's API-keys page.

Full policy: see your Synaplan instance's **Privacy** and **Terms** pages.

---

## 11. It didn't work? (quick fixes)

| What you see                                   | Try this                                                                                                           |
| ---------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| The **Synamail** button isn't in the ribbon    | Refresh Outlook on the Web (Ctrl/Cmd+F5), or restart new Outlook. It can take a minute after install.              |
| No **"Add a custom add-in"** option            | Your work account blocks add-ins — ask IT to enable the certified Synamail, or use a personal Outlook.com account. |
| Sign-in window opens then closes with an error | Just try again. If it keeps happening, sign out and sign in once more.                                             |
| **Summarise** does nothing or errors           | Your Synaplan server may be unreachable, or you were signed out. Open **Settings → Log out** and sign in again.    |
| Everything looks blank                         | Close and reopen the panel; if it persists, remove and re-add the add-in.                                          |

Still stuck? Use the **? Help** link in the panel, or contact support via your Synaplan
instance's support page.

---

_Looking to build, contribute, or self-host Synamail? Start at
[`../README.md`](../README.md) → the "build / host it" lane._
