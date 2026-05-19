# Synamail — Feature Specifications

Concrete feature specs for the Outlook add-in. Every section names: the user value, the Synaplan endpoint(s) used, the UI surface, the data model (where new), and edge cases.

This is the integration-ready feature contract. If anything here contradicts `planning/PLAN.md` or `docs/ARCHITECTURE.md`, this file wins for _what the feature does_; the other documents win for _when/where it's built_.

## 1. Read-mode AI actions

The user opens an email in Outlook → the Synamail taskpane offers five actions on the selected message.

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

## 2. Compose-mode AI actions

The user is writing or replying → the Synamail taskpane offers writing assistance.

### 2.1 Draft from prompt

- **User value:** Generate a full email body from a one-line intent.
- **Input:** User intent text + (optional) referenced email if this is a reply.
- **Output:** HTML body inserted via `body.setAsync(html, { coercionType: 'html' })`.
- **Endpoint:** `POST /api/v1/messages/send` with a `compose(intent)` system prompt.

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

## 4. Contact knowledge base — search by sender / recipient

This is the feature the user asked for: _"search by sender or recipient to create a knowledge base for these contacts"_. Implemented as a **per-contact RAG group convention**, no Synaplan-side schema changes.

### 4.1 Data model (convention only — no backend change)

- For every contact the user accumulates emails about, Synamail uses a RAG group with the name **`contact:<lowercased-email>`** (e.g. `contact:alice@example.com`).
- The group is created on demand the first time the user clicks "Save this email to <alice@example.com>'s knowledge base".
- Group naming is the only thing that makes a group a "contact" group. The plugin treats any group with the `contact:` prefix as a contact group when populating the ContactKnowledgeBase view.

### 4.2 UI surfaces

- **In ReadMode:** a "Contact: alice@example.com ▾" pill below the subject. Click → opens `ContactKnowledgeBase.vue` scoped to that contact. The pill is a picker for multi-recipient threads (sender + each To/Cc address).
- **`ContactKnowledgeBase.vue`** shows:
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

## 5. RULE integration — Synapse Routing rules

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

## 6. Feature coverage matrix (cross-check with planning/PLAN.md)

| Feature                               | Sprint | Endpoint(s)                                                   | View                            | Status target               |
| ------------------------------------- | ------ | ------------------------------------------------------------- | ------------------------------- | --------------------------- |
| Summarise                             | 3      | `messages/send`                                               | ReadMode                        | Live                        |
| Translate (email body)                | 3      | `messages/send`                                               | ReadMode                        | Live                        |
| Draft reply                           | 3      | `messages/send` + `displayReplyForm`                          | ReadMode → Outlook compose      | Live                        |
| Classify                              | 3      | `messages/send` (JSON-mode)                                   | ReadMode                        | Live                        |
| Ask follow-ups                        | 3      | `chats` + `chats/{id}/messages`                               | ReadMode                        | Live                        |
| Save to RAG (chosen group)            | 3      | `files/upload` + `files/{id}/process` + `files/groups`        | ReadMode + modal                | Live                        |
| Contact knowledge base — save         | 3      | `files/groups` (POST) + upload/process with `contact:<email>` | ReadMode → contact pill         | Live                        |
| Contact knowledge base — search       | 3      | `rag/search` with group filter                                | ContactKnowledgeBase            | Live                        |
| Contact knowledge base — ask          | 3      | `chats` with RAG scope hint                                   | ContactKnowledgeBase            | Live                        |
| Insert from RAG (compose)             | 3      | `rag/search`                                                  | ComposeMode                     | Live                        |
| Draft / improve / translate selection | 3      | `messages/send`                                               | ComposeMode                     | Live                        |
| RULE integration — view               | 3      | `prompts` (or `synapse/topics`)                               | RuleEditor                      | Live (read)                 |
| RULE integration — edit               | 3      | `prompts` PATCH                                               | RuleEditor                      | Live or deferred (see §5.3) |
| Sign in (Office Dialog → API key)     | 2      | `/addin/connect` bridge → `apikeys` POST                      | SignIn                          | Live                        |
| Sign out (revoke key)                 | 2      | `apikeys` DELETE                                              | Settings                        | Live                        |
| Self-hosted instance override         | 2      | n/a (client-side)                                             | Settings → "Use a self-hosted…" | Live                        |

Anything not in this table is either deferred to a future phase (Smart Alerts, `OnNewMessageCompose` autoload, mobile) or out of scope for v1.
