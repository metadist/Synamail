# Synamail — Feature Specifications

Concrete feature specs for the Outlook add-in. Every section names: the user value, the Synaplan endpoint(s) used, the UI surface, the data model (where new), and edge cases.

This is the integration-ready feature contract. If anything here contradicts `docs/PROJECT_PLAN.md` or `docs/ARCHITECTURE.md`, this file wins for _what the feature does_; the other documents win for _when/where it's built_.

> **v1 release scope (2026-06-10).** For the AppSource submission the app was
> deliberately condensed to the features that work on **all** required Outlook
> hosts: Home chat, the read-mode actions (§1), Save to knowledge base (§3),
> Contact AI Profiling (§4 + §8, served by the `synamail` Synaplan plugin),
> find-meeting-times, and Settings. Compose-mode assistance (§2), RULE/Synapse
> editing (§5), Mail Routes (§7), and every EWS-dependent feature (mailbox-wide
> search, sender history, block sender) are **cut from v1** — EWS retires for
> Exchange Online in Oct 2026 and `makeEwsRequestAsync` is unavailable in new
> Outlook for Windows. The coverage matrix in §6 is authoritative.

## 1. Read-mode AI actions

The user opens an email in Outlook → the Synamail taskpane offers its actions on the selected message (rendered by `EmailActionsPanel.vue` inside Home's "Email actions" section).

### 1.1 Summarise

- **User value:** Get a bulleted summary of a long email without reading it.
- **Input:** Email body (`Office.context.mailbox.item.body.getAsync('text')`), plus headers (`from`, `to`, `subject`).
- **Output:** 3–7 bullet points; renders inline in the taskpane.
- **Endpoint:** `POST /api/v1/messages/send` with a `summarise` system prompt template and the email body. Language follows the user's Outlook display language (`Office.context.displayLanguage`), overridable via Settings.
- **UI surface:** "Summarise" button in `ReadMode.vue`. Loading spinner, error toast on failure, copy-to-clipboard on success.
- **Edge cases:**
  - Body > token limit → chunk and reduce-then-summarise; surface a "long email — summarising in chunks" hint.
  - HTML-only body → fall back to `body.getAsync('html')` and strip markup client-side before sending.

### 1.2 Translate

- **User value:** Read an email in the user's language regardless of the sender's.
- **Input:** Email body + target language (default: Outlook display language; selector for en/de/fr/es/it/zh/ar).
- **Output:** Translated body rendered in the taskpane; "copy" + "open as draft" actions.
- **Endpoint:** `POST /api/v1/messages/send` with a `translate(targetLang)` system prompt.
- **UI surface:** "Translate" button + target-language picker in `ReadMode.vue`.
- **Edge cases:** Auto-detect source language via the AI; if source = target, surface a hint and skip the call.

### 1.3 Draft reply

- **User value:** Pre-fill the reply window with an AI-suggested response.
- **Input:** Email body + thread context + tone selector (formal / concise / friendly) + language selector.
- **Output:** Reply opens in Outlook's compose UI with the generated HTML body pre-filled.
- **Endpoints:**
  1. `POST /api/v1/messages/send` with a `reply(tone, lang)` system prompt.
  2. `Office.context.mailbox.item.displayReplyForm({ htmlBody, attachments: [] })`.
- **UI surface:** "Draft reply" button with tone + language controls.
- **Edge cases:** If the thread has prior messages, include the last 2–3 turns in the prompt for context (capped to fit token budget).

### 1.4 Classify

- **User value:** Auto-categorise the email (invoice, support, internal, personal, spam, …) for the user's own triage.
- **Input:** Email body + headers.
- **Output:** `{ category, confidence, reasoning }` JSON; rendered as a labelled pill + tooltip.
- **Endpoint:** `POST /api/v1/messages/send` with a `classify` system prompt that enforces JSON output.
- **UI surface:** "Classify" button; result is sticky for the current message until the user dismisses it.
- **Edge cases:** If confidence < 0.5, show "unsure" + suggest manual triage. Categories list is configurable via Settings in a later iteration; v1 uses a fixed list.

### 1.5 Ask follow-ups

- **User value:** Multi-turn chat about the current email thread without leaving Outlook.
- **Input:** User-typed question + the email body as context.
- **Output:** AI answer; subsequent questions continue the same chat.
- **Endpoints:**
  - First turn per Outlook conversation: `POST /api/v1/chats` to create the chat, **keyed by `mailbox.item.conversationId`** stored in `roamingSettings.chats[<conversationId>]`.
  - Subsequent turns: `POST /api/v1/chats/{chatId}/messages` (or whichever the Synaplan API exposes for sending into an existing chat — confirm during wire-up).
  - Display: `GET /api/v1/chats/{chatId}/messages`.
- **UI surface:** Anchored input at the bottom of `ReadMode.vue`. History is scrollable.
- **Edge cases:** Reopening the same thread later reuses the existing `chatId`. If the chat was deleted server-side, fall back to creating a new one and update roaming.

## 2. Compose-mode AI actions — **partially shipped**

> **§2.1 "Draft from prompt" shipped (2026-07-06).** The compose ribbon button
> opens the Home taskpane, whose **Writing assistant** section (auto-expanded in
> compose mode, `ComposeAssistantPanel.vue`) turns a one-line intent into a full
> email body written straight into the open draft. This closes the AppSource
> review finding that the pane showed "no interaction" in compose mode.
> §2.2 (rewrite selection) and §2.3 (insert from RAG) remain deferred — the
> specs below stay as their design.

The user is writing or replying → the Synamail taskpane offers writing assistance.

### 2.1 Draft from prompt — **shipped (2026-07-06)**

- **User value:** Generate a full email body from a one-line intent.
- **Input:** User intent text + tone + language + (optional) the replied-to email
  body when composing a reply/forward.
- **Output:** HTML body written via `body.setAsync(html, { coercionType: 'html' })`.
- **Endpoint:** `POST /api/v1/messages/send` with a `compose(tone, lang)` system prompt.
- **UI surface:** Home → **Writing assistant** section (`ComposeAssistantPanel.vue`),
  shown only in compose mode; intent textarea + tone/language pickers + "Draft".
- **Edge cases:** Empty intent disables the button; a host that rejects the body
  write surfaces `compose.insertFailed`.

### 2.2 Improve / Shorten / Translate selection

- **User value:** Rewrite the currently selected text.
- **Input:** `body.getSelectedDataAsync` → selected text + transformation (`improve` / `shorten` / `translate(lang)`).
- **Output:** `body.setSelectedDataAsync` replaces the selection.
- **Endpoint:** `POST /api/v1/messages/send` with the appropriate system prompt.
- **Edge cases:** No selection → button disabled with tooltip.

### 2.3 Insert from RAG

- **User value:** Paste a snippet (with citation) from the user's Synaplan knowledge base into the email being written.
- **Input:** Search query → list of RAG results.
- **Output:** On click, formatted snippet + citation inserted via `setSelectedDataAsync`.
- **Endpoint:** `POST /api/v1/rag/search` with `{ query, threshold, limit, groups? }`.
- **UI surface:** Search bar + scrollable result list inside `ComposeMode.vue`.

## 3. RAG ingestion — "Save to knowledge base"

Note on terminology: the user's brief says "**RAG keys**"; Synaplan calls these **RAG file groups** (table `BFILEGROUPS`, endpoint `GET /api/v1/files/groups`). This document uses "RAG group" throughout.

### 3.1 Save email + attachments to a chosen group

- **User value:** Make the email and its attachments retrievable later by Synaplan AI.
- **Input:**
  - The email itself: `Office.context.mailbox.item.getAsFileAsync` (returns `.eml`) — requires Mailbox 1.8.
  - Each attachment: `getAttachmentsAsync` + `getAttachmentContentAsync` — requires Mailbox 1.8.
  - A target RAG group (user-picked or default).
  - A processing level: `Extract Only` (default) / `Extract + Vectorize` / `Full Analysis` — matches the levels in `synaplan/docs/RAG.md`.
- **Output:** File appears in the user's RAG group; success toast.
- **Endpoints:**
  1. `POST /api/v1/files/upload` for each artefact.
  2. `POST /api/v1/files/{id}/process` with the chosen processing level and group id.
- **UI surface:** "Save to knowledge base" button in `ReadMode.vue`; modal with group picker, attachment checkboxes, processing-level dropdown.
- **Default group suggestion:** `contact:<sender-email>` — see §4.

### 3.2 Group picker

- Lists existing groups via `GET /api/v1/files/groups`.
- Lets the user create a new group inline (`POST /api/v1/files/groups`).
- Pre-selects the **last-used group** for this Outlook (stored in `roamingSettings.lastRagGroupId`).
- Shows the contact-default as a separate suggestion chip at the top — clicking it picks `contact:<sender>`.

## 4. Contact AI Profiling — per-contact RAG group (search by sender / recipient)

This is the feature the user asked for: _"search by sender or recipient to create a knowledge base for these contacts"_. Implemented as a **per-contact RAG group convention**, no Synaplan-side schema changes. This per-contact corpus is the data foundation for **Contact AI Profiling** — see [`CONTACT_PROFILING.md`](CONTACT_PROFILING.md) for the rolling-profile design built on top of it (§8 has the summary).

### 4.1 Data model (convention only — no backend change)

- For every contact the user accumulates emails about, Synamail uses a RAG group with the name **`contact:<lowercased-email>`** (e.g. `contact:alice@example.com`).
- The group is created on demand the first time the user clicks "Save this email to <alice@example.com>'s knowledge base".
- Group naming is the only thing that makes a group a "contact" group. The plugin treats any group with the `contact:` prefix as a contact group when populating the ContactProfile view.

### 4.2 UI surfaces

- **In ReadMode:** a "Contact: `alice@example.com` ▾" pill below the subject. Click → opens `ContactProfile.vue` scoped to that contact. The pill is a picker for multi-recipient threads (sender + each To/Cc address).
- **`ContactProfile.vue`** shows:
  - The contact's email + name (from the original Outlook headers).
  - A search box that runs `POST /api/v1/rag/search` filtered to `groups=[contact:<email>]`.
  - A list of saved emails with subject + date + snippet; click → open the source email in Outlook (where the Office API allows it).
  - A "Save current email to this contact" button.
  - An "Ask about this contact" button that opens a chat with the contact group as the RAG scope.

### 4.3 Endpoints used

- `GET /api/v1/files/groups` — find / create the `contact:<email>` group.
- `POST /api/v1/files/groups` — create on first use.
- `POST /api/v1/files/upload` + `POST /api/v1/files/{id}/process` — ingest the email into the contact's group.
- `POST /api/v1/rag/search` with `groups: ["contact:<email>"]` — search within the contact's knowledge.
- `POST /api/v1/chats` with a RAG scope hint to the contact's group — "Ask about this contact".

### 4.4 Privacy + edge cases

- Email addresses are case-insensitive — always lowercase before forming the group name.
- Strip Outlook-internal `+suffix` aliases from contact-group names? **No** — the suffix can be meaningful (e.g. `alice+invoices@…`). The user can rename a group in Synaplan if they want to merge identities.
- For internal-only senders (e.g. `alice@yourcompany.com`), the same convention applies — there is no special-casing.
- Saving an email to a contact group **also** records the recipient list as searchable metadata on the file (via the `BFILES.BTAGS` field if available, confirm during wire-up), so a search for "alice" returns hits where alice was on the To: line too.

## 5. RULE integration — Synapse Routing rules — **deferred (not in v1)**

> **Cut from the v1 release (2026-06-10).** `RuleEditor.vue` was never built;
> the spec below is the design for a later release.

Note: in the user's brief this was called "RULE integration". Synaplan's actual rule engine is **Synapse Routing** (`docs/SYNAPSE_ROUTING.md` in `synaplan/`), and the rules are stored in `BSELECTION_RULES` on each routable topic. Synamail surfaces this engine inside Outlook so users can shape how Synaplan routes their incoming emails (e.g. via the `smart+keyword@synaplan.net` webhook flow).

### 5.1 What rules do

- Each Synaplan routing topic (e.g. `billing`, `support`, `research`) has a `BSELECTION_RULES` field — Tier-0 if/then matchers that **win immediately** over embedding and AI classification.
- Example: `IF subject contains "invoice" THEN topic=billing`.

### 5.2 What Synamail offers

- **`RuleEditor.vue`** — a view inside the taskpane that:
  - Lists the user's routable topics via `GET /api/v1/prompts` (confirmed against `PromptController`).
  - Shows the current `BSELECTION_RULES` for each topic.
  - Lets the user add / remove keyword matchers — **subject to the write-endpoint check in §5.4**.
  - Offers a **"Test against current email"** preview that calls `POST /api/v1/admin/synapse/dry-run` — **admin-only**, hidden for non-admin users with an inline hint.
- **"Apply this email's pattern as a rule"** — one-click action in ReadMode that pre-fills the RuleEditor with a candidate rule based on the current email (subject substring, sender domain).

### 5.3 Endpoints used (verified against `synaplan/backend/src/Controller/`)

- `GET /api/v1/prompts` — list the user's routable topics (`PromptController`). `BSELECTION_RULES` is a field on the `Prompt` entity (`backend/src/Entity/Prompt.php`).
- `PATCH /api/v1/prompts/{id}` (or whichever write method `PromptController` exposes) — update `BSELECTION_RULES`. Confirm exact method/shape in Sprint 2 wire-up.
- **`POST /api/v1/admin/synapse/dry-run` — admin-only** (`AdminSynapseController`). A regular Synamail user cannot call this. Therefore the "Test against current email" preview in `RuleEditor.vue` is **disabled for non-admin users in v1**; the button is hidden and an inline hint explains why. For admins, the preview is available.

### 5.4 Risks and v1 deferrals

- If `PromptController` does not expose a write endpoint for `BSELECTION_RULES` to non-admin users, RULE integration is **read-only** in v1. `RuleEditor.vue` then shows existing rules but the "+ Add rule" affordance is hidden; a Synaplan-side ticket is filed to add a user-scoped write endpoint, and editing arrives in v1.1.
- The dry-run preview is gated to admins regardless of the read/write decision above.
- **Confirm in Sprint 2 wire-up:** open `PromptController.php` and document the actual read/write surface in `STEPS.md` Step 2.x before starting Step 3.7.

### 5.4 UI surface

- Settings → "Email routing rules" tab opens `RuleEditor.vue`.
- ReadMode → "···" overflow menu → "Use this email to create a routing rule".

## 6. Feature coverage matrix — v1 release scope (authoritative)

### Shipped in v1

| Feature                                | Endpoint(s)                                                              | View                           | Status                                        |
| -------------------------------------- | ------------------------------------------------------------------------ | ------------------------------ | --------------------------------------------- |
| Home chat (general Synaplan chat)      | `chats` + `messages/send` (+ `messages/stream`)                          | Home                           | Live                                          |
| Summarise                              | `messages/send`                                                          | Home → Email actions           | Live                                          |
| Translate (email body)                 | `messages/send`                                                          | Home → Email actions           | Live                                          |
| Draft reply                            | `messages/send` + `displayReplyForm`                                     | Home → Email actions → compose | Live                                          |
| Classify                               | `messages/send` (JSON-mode)                                              | Home → Email actions           | Live                                          |
| Find meeting times                     | `messages/send` + `displayNewAppointmentForm`                            | Home → Email actions           | Live                                          |
| Ask follow-ups (incl. email images)    | `chats` + `messages/send` + `files/upload`                               | Home → Email actions           | Live                                          |
| Draft from prompt (compose mode)       | `messages/send` (`compose` prompt) + `body.setAsync`                     | Home → Writing assistant       | Live (2026-07-06)                             |
| Save to RAG (chosen group)             | `files/upload` (with `process_level`) + `files/groups`                   | Home → Email actions + modal   | Live                                          |
| Contact AI Profiling — corpus          | upload/search scoped to `contact:<email>`                                | ContactProfile                 | Live                                          |
| Contact AI Profiling — rolling profile | `synamail` plugin: `GET/POST/DELETE …/plugins/synamail/profiles/{email}` | ContactProfile profile card    | Live (needs the plugin installed server-side) |
| Sign in (Office Dialog → API key)      | `/addin/connect` bridge → `apikeys` POST                                 | SignIn                         | Live                                          |
| Sign out (revoke key)                  | `apikeys` DELETE                                                         | Settings                       | Live                                          |
| Self-hosted instance override          | n/a (client-side)                                                        | SignIn / Settings              | Live                                          |
| Model display + language pref          | `config/models` + `config/models/defaults`                               | Settings                       | Live                                          |

### Cut from v1 (design retained)

| Feature                                          | Why cut                                                  | Where the design lives |
| ------------------------------------------------ | -------------------------------------------------------- | ---------------------- |
| Compose-mode rewrite selection / insert from RAG | Deferred after §2.1 shipped (draft-from-prompt is live)  | §2.2 + §2.3            |
| RULE / Synapse routing editor                    | Never built                                              | §5                     |
| Mail Routes (per-email automations)              | Config UI existed without a runtime engine               | §7 + `MAIL_ROUTES.md`  |
| Mailbox search → knowledge base                  | EWS-only (mock fallback elsewhere); EWS retires Oct 2026 | git history            |
| "More from this sender" history                  | EWS-only with mock fallback                              | git history            |
| Block sender                                     | EWS-only, never had UI                                   | git history            |

Anything not in these tables is either deferred to a future phase (Smart Alerts, `OnNewMessageCompose` autoload, mobile) or out of scope for v1.

## 7. Mail Routes — per-email AI agent triggers (design / RFC)

> **Not** Synaplan's Synapse Routing (§5, which selects a system-prompt/topic).
> A **Mail Route** is a Synamail automation: **WHEN** a mail matches conditions
> **THEN** run one or more AI/agent actions on it.

**User value:** define rules like _"from Oliver AND requests a meeting → suggest
3 free times"_, _"from `a@b.c` AND English → translate & show"_, or _"always →
add to knowledge group `contracts`"_. Conditions can be deterministic (sender,
keyword, attachment) or AI-evaluated (language, intent, category); actions can be
silent/auto (add to KB) or interactive (show translation, pick a meeting slot).

**Constraint:** an Outlook add-in is not a background service — routes fire when
the user is in Outlook and the matching mail surfaces (pinned taskpane +
`ItemChanged`), opt-in per route, transparent + undoable. A 24/7 server-side
path (Synaplan `InboundEmailHandler`) is an optional later phase.

**Endpoints:** reuses existing actions — `messages/send` (translate / summarise /
classify / intent / draft reply), `files/upload` (add to knowledge), plus a
future Graph `Calendars.Read` for the meeting agent's free/busy lookup.

Full model, action catalog (with feasibility), data model, trust checklist, and
phased plan live in [`MAIL_ROUTES.md`](MAIL_ROUTES.md). **Not yet implemented** —
design-first per the 2026-05-31 decision.

## 8. Contact AI Profiling — rolling relationship memory (**shipped in v1, Phase 1**)

> Renames and supersedes the "Contact knowledge base" of §4. The
> `contact:<email>` RAG group stays; profiling adds a durable, recomputed
> **state object** on top of it.
>
> **Shipped 2026-06-10 (Phase 1) — via the `synamail` Synaplan plugin.** The
> implementation supersedes the client-only "Option B" below: the rolling
> prompt and the stored profile live **server-side** in the plugin
> (`synamail-plugin/`, released to `synaplan/plugins/synamail/`), keyed per
> user + contact in Synaplan's generic `plugin_data` store. The add-in rolls
> one email at a time into the profile (`POST …/profiles/{email}/update`) —
> explicitly via the profile card, and automatically whenever an email is
> saved to the contact's group. See `CONTACT_PROFILING.md` for the updated
> architecture record.

**User value:** opening a contact instantly answers "where does this
relationship stand?" — _"you haven't mailed in 6 weeks; last exchange friendly
but distanced"_, or _"you promised a demo but never mailed back"_. Built from
exchanged emails **plus manual note snippets** the user adds (calls, meetings,
verbal promises), so off-channel reality is captured. Tracks **open commitments**
that a later note (_"demo delivered over phone"_) resolves.

**Inputs:** emails (`fileUpload` into the contact group), manual notes
(`type: note` text files), deterministic Outlook signals (last contact, cadence,
who-owes), contact identity.

**Engine:** targeted `rag/search` + code-computed deterministic facts + a
`profileContact` JSON-mode prompt (`messages/send`, like `classify`). "Rolling" =
recompute on open / save / add-note (add-ins aren't background services).

**Data architecture (locked 2026-06-07): Option B — local-first, server-as-truth,
no platform changes.** The corpus (emails + notes) lives in the user's Synaplan
workspace = one synced source of truth across all their machines; the summary is
_derived_ and cached locally (roaming LRU), optionally pinned as a `profile.json`
artifact in the group for consistent cross-machine wording. **Sync the inputs,
not the summary.** Reuses the existing `files`/RAG API only — no new tables,
migrations or endpoints. GDPR: user-owned, transparent, fully deletable.

Full model, storage tiers, lifecycle, UI, privacy and phased plan live in
[`CONTACT_PROFILING.md`](CONTACT_PROFILING.md). **Not yet implemented** —
design-first.
