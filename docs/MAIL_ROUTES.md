# Synamail Mail Routes — Design (per-email AI agent triggers)

**Status:** Phase 1 in progress. The **foundation has landed** — the route
data model, roaming persistence, the pure-logic core (ICS builder, sender
matcher, calendar-conflict detection, translate-before-ingest planner) with
unit tests, and the **"Mail Routes" config view** (three route kinds: meeting /
date helper, add-to-project, newsletter knowledge base). The **runtime engine**
(trigger wiring on `ItemChanged`, calendar read, attachment fetch, AI intent
evaluation, and the action outputs) is the next increment. Owner: Synamail.
Companion to [`FEATURES.md`](FEATURES.md) and [`ARCHITECTURE.md`](ARCHITECTURE.md).

> **Read this first — terminology.** A "**Mail Route**" here is a Synamail
> automation: **WHEN** a mail matches conditions **THEN** run one or more
> AI/agent actions on it. This is **NOT** Synaplan's _Synapse Routing_
> (`BSELECTION_RULES`), which only picks a system prompt/topic for the AI and is
> a separate, server-side concern (edited by `RuleEditor.vue`). Do not conflate
> the two. If anything in this repo blurs them, this doc wins for the
> _agent-trigger_ meaning.

---

## Decisions locked (2026-05-31)

1. **Naming.** The existing Synapse-Routing view (`RuleEditor.vue`) is renamed
   **"Synaplan topics"** so "Routes" only ever means Mail Routes. (Done.)
2. **Meeting agent.** Read the calendar to detect conflicts and propose open
   slots. We do **not** auto-create or override meetings — we suggest, the user
   confirms. **Calendar-read path (revised 2026-06-01): EWS first.** The
   existing-meeting check uses an EWS `CalendarView` lookup (`useOutlookMailbox`
   already uses EWS heavily), which ships **without an Azure app registration
   or admin consent** but is limited to classic Outlook desktop + Outlook on the
   Web. The conflict-detection logic (`findConflict` in
   `shared/mail-routes/match.ts`) is pure and host-agnostic, so a Graph / NAA
   provider can be swapped in later for "new Outlook for Windows" without
   touching the engine.
3. **ICS is a must.** The meeting agent generates a standards-compliant **`.ics`
   (VEVENT)** the user can attach to a reply to send the invite — alongside the
   "add to my calendar" path. (Reverses the earlier ICS-out-of-scope call.)
4. **Builder UX = hybrid.** Structured **Sender + keyword dropdowns** for the
   deterministic match, **plus a free-text "AI clarify" prompt** that guides the
   AI conditions/actions (e.g. _"only if it's really a meeting request, not just
   mentioning a past meeting"_).
5. **Phase 1 must-have = Categorize.** AI assigns the best-fitting **native
   Outlook category** to the mail, using user-defined category→meaning mappings
   (incl. repurposing the stock color names — see §4a).

---

## 1. What a Mail Route is

```
Route = WHEN <conditions on the email> THEN <ordered list of actions>
```

Concrete examples (the user's own):

- **Add to knowledge:** _always_ → add the mail (+ attachments) to KB group
  `contracts`.
- **Translate & show:** _from `a@b.c` AND language = English_ → translate to
  `XYZ` and show the translation in the taskpane.
- **Meeting agent:** _from Oliver AND intent = "requests a meeting"_ → read the
  user's calendar, propose 3 free slots, offer to draft the reply / create the
  appointment.

A route can chain multiple actions. Conditions can be **deterministic**
(sender, keyword, has-attachment) or **AI-evaluated** (language, intent,
category). Actions can be **silent/auto** (add to KB) or **interactive**
(show a translation; present 3 slots to pick).

---

## 2. The one hard constraint (unchanged, still true)

An Outlook add-in is **not a background service**. Per the 2026-05-31 decision
the trigger is **inbound + add-in-driven**: a route fires when the user is in
Outlook and the matching mail surfaces (pinned taskpane + `ItemChanged`). That's
hands-free per email, but not silent background processing — which also keeps us
within `ARCHITECTURE.md §8.6` data-minimisation. A true 24/7 path (Synaplan's
`InboundEmailHandler` IMAP connector) remains an optional later phase.

---

## 3. Conditions (the WHEN)

| Condition                                                                | Type          | How evaluated                                      |
| ------------------------------------------------------------------------ | ------------- | -------------------------------------------------- |
| `from` (address or `@domain`)                                            | deterministic | `item.from`                                        |
| `subjectContains` / `bodyContains`                                       | deterministic | string match                                       |
| `hasAttachment` / attachment type                                        | deterministic | `item.attachments`                                 |
| `language is X`                                                          | **AI**        | detect via `messages/send` (or a cheap classifier) |
| `intent is X` ("requests a meeting", "is an invoice", "asks a question") | **AI**        | one classify call returning an enum + confidence   |
| `category is X`                                                          | **AI**        | reuse the existing `classify` action               |

AI-evaluated conditions are **only run when the cheap deterministic conditions
already match** (cost control): e.g. evaluate "requests a meeting" only on mail
from Oliver, not every mail.

---

## 4. Actions (the THEN) — the agent tool catalog

Each action is an "agent tool." Feasibility today:

| Action                    | What it does                                                                                                                              | Feasibility                                                                                                                  | Reuses                                                                |
| ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| **Add to knowledge**      | Upload body + attachments to a RAG group, vectorized                                                                                      | ✅ ready                                                                                                                     | `fileUpload`, `getReadItemAsFile`, `getImageAttachments`              |
| **Translate & show**      | Translate body to a target language, render in pane                                                                                       | ✅ ready                                                                                                                     | `translate`                                                           |
| **Summarise & show**      | Bulleted summary in pane                                                                                                                  | ✅ ready                                                                                                                     | `summarise`                                                           |
| **Categorize (Outlook)**  | AI picks the best-fitting **native Outlook category** and applies it to the mail                                                          | ✅ ready (Mailbox 1.8)                                                                                                       | `classify`-style call + Office.js categories API (§4a)                |
| **Classify / show**       | Show a category pill in the pane (no mailbox change)                                                                                      | ✅ ready                                                                                                                     | `classify`                                                            |
| **Draft reply**           | Generate a reply and open the compose form                                                                                                | ✅ ready                                                                                                                     | `draftReply` + `displayReplyForm`                                     |
| **Suggest meeting times** | Extract requested times, **read calendar free/busy**, propose open slots; user picks → create appointment and/or attach `.ics` to a reply | ⚠️ Graph dep accepted — free/busy via `Calendars.Read` (NAA, ARCHITECTURE §7.5). Suggest/confirm only — never auto-override. | `extractMeetingTimes`, `displayNewAppointmentForm`, ICS builder (§4b) |
| **Forward / notify**      | Forward to an address or post a Synaplan note                                                                                             | later                                                                                                                        | —                                                                     |

### 4a. Categorize action (Phase 1 detail)

Outlook has native, color-coded **categories**. The stock "Red/Blue/Green…
Category" names are meaningless, so users **repurpose them with meaning**:

- The user maps a category to a description/example, e.g. **"Blue Category" →
  _"about project XYZ AND from the `bmw.de` domain"_**.
- Mappings live in roaming (shared across routes). The AI receives the mail plus
  the candidate categories' descriptions and returns the best match + a
  confidence; below a threshold it applies nothing.
- Apply via Office.js: `Office.context.mailbox.masterCategories.getAsync`
  (and `addAsync` to ensure the category/color exists), then
  `item.categories.addAsync([name])`. Needs `ReadWriteMailbox` (have it) +
  Mailbox 1.8. Reversible via `item.categories.removeAsync`.
- Transparent: toast "Tagged **Blue** — project XYZ ✓ — Undo".

### 4b. ICS builder (meeting agent)

- Generate a minimal RFC-5545 **VEVENT** (uid, dtstart/dtend, summary, location,
  organizer/attendee) client-side — no backend.
- Attach to a reply via `displayReplyForm({ htmlBody, attachments: [{ type:
'file', name: 'invite.ics', ... }] })` so the user sends a real invite to the
  sender. Complements (doesn't replace) "add to my calendar"
  (`displayNewAppointmentForm`).

**The meeting agent's only true gap is reading free/busy** (Graph/NAA, accepted).
Time extraction, appointment creation, and ICS generation are all in reach.

---

## 5. Data model (routes live in roaming settings)

Per-mailbox, drives a client-side trigger → store in `roamingSettings`
(`synamail.routes`), readable by the taskpane (and a future event runtime).

```ts
type ConditionType =
  | 'from'
  | 'subjectContains'
  | 'bodyContains'
  | 'hasAttachment'
  | 'language'
  | 'intent'
  | 'category'

interface RouteCondition {
  type: ConditionType
  value: string // e.g. "oliver@acme.com", "en", "requests_meeting", "pdf"
}

type ActionType =
  | 'addToKnowledge'
  | 'translateShow'
  | 'summariseShow'
  | 'classifyLabel'
  | 'draftReply'
  | 'suggestMeetingTimes'

interface RouteAction {
  type: ActionType
  params?: Record<string, string> // e.g. { groupId: "contracts" }, { targetLang: "de" }
}

interface MailRoute {
  id: string
  name: string
  enabled: boolean
  match: 'all' | 'any' // how conditions combine (default 'all')
  conditions: RouteCondition[]
  actions: RouteAction[]
  /** Auto actions run silently; interactive actions render results for the user. */
}

interface RoutesState {
  paused: boolean
  routes: MailRoute[]
  seen: Record<string, number> // internetMessageId -> lastRunAt (dedup auto actions)
}
```

---

## 6. Execution model

On `ItemChanged` for a read item (or when the taskpane opens on a mail):

1. **Cheap match** — evaluate deterministic conditions for each enabled route.
2. **AI match** — for routes still matching, evaluate AI conditions
   (language/intent/category) with the fewest calls (batch where possible).
3. **Run actions** in order:
   - **Auto** actions (add to KB) run once per message (dedup on
     `internetMessageId`), with a toast + undo.
   - **Interactive** actions (translate/show, suggest times) render a result
     card in the taskpane for the user to act on — they are **not** auto-applied
     to the mailbox without a click.
4. **Audit** — every run is logged ("Route 'Contracts' → added 2 files").

Intent detection ("requests a meeting") = one `classify`-style call returning an
enum; reuse the existing JSON-mode parsing pattern.

---

## 7. Trust / sensitivity (unchanged principles)

- Opt-in per route, default OFF; first auto-run per route asks for confirmation.
- Every action visible (toast + audit log); silent only for explicitly-auto
  actions the user enabled.
- One-click undo for KB captures; global pause; per-route enable.
- Dedup on `internetMessageId`; allow-listed attachment types; size threshold to
  skip signature logos.
- Interactive actions never modify the mailbox without a user click.
- Respect Synaplan rate limits; gate AI-condition evaluation behind cheap
  deterministic matches to control cost.

---

## 8. Relationship to existing pieces

- **`RuleEditor.vue` + `listTopics`/`updateTopicRules`/`testRouting`** — that is
  **Synapse Routing** (system-prompt/topic selection), a _different_ feature
  (`FEATURES.md §5`). **Renamed to "Synaplan topics"** in the UI so "Routes" only
  means Mail Routes. Functionality unchanged.
- **Meeting-times + appointment** (`extractMeetingTimes`,
  `displayNewAppointmentForm`) — become the building blocks of the
  `suggestMeetingTimes` action (plus the calendar-read gap).
- **`getReadItemAsFile` / `getImageAttachments` / `fileUpload`** — building
  blocks of the `addToKnowledge` action.

---

## Builder UX (locked = hybrid)

A route is created with **structured dropdowns + a free-text AI prompt**:

- **Match (structured):** Sender (address / `@domain`) and Keyword
  (subject/body contains) dropdowns + has-attachment toggle — fast, deterministic.
- **AI clarify (free text):** an optional prompt that refines the AI conditions
  and actions in natural language, e.g. _"only if this is an actual meeting
  request, not just referencing a past one"_ or _"categorize as the project it's
  about"_. This text is passed to the AI step alongside the structured match.
- **Action picker:** choose one or more actions (Categorize, Add to knowledge,
  Translate & show, Summarise, Draft reply, Suggest meeting times) with their
  params (group, target language, category map…).

This keeps the common case point-and-click while letting power users express
nuance the dropdowns can't.

---

## 9. Phased plan

- **Phase 0 — spike (0.5d):** pinning + `ItemChanged` support on target clients;
  scope the Graph free/busy path (NAA) for the meeting agent.
- **Phase 1 — Routes engine + Categorize + core actions:** data model +
  "Mail Routes" builder (hybrid: dropdowns + AI-clarify prompt) + trigger +
  actions **`categorize` (Outlook native categories — the headline must-have)**,
  `addToKnowledge`, `translateShow`, `summariseShow`, `draftReply`. Conditions:
  deterministic + `language`/`category`/`intent` (AI). Includes the category→
  meaning map UI (repurposing the stock color names).
- **Phase 2 — Meeting agent:** Graph `Calendars.Read` (NAA) → propose open
  slots; user picks → create appointment **and/or attach a generated `.ics`** to
  the reply.
- **Phase 3 — hands-off + server-side (optional):** event-based launch event;
  optional Synaplan `InboundEmailHandler` for 24/7 capture.

---

## 10. Open questions (resolved 2026-05-31)

1. ~~Naming~~ → **"Synaplan topics"** (done).
2. ~~Calendar read~~ → **Graph/NAA accepted**; suggest-only, never override.
3. ~~Builder UX~~ → **hybrid** (structured dropdowns + AI-clarify prompt).
4. ~~Phase 1 must-haves~~ → **Categorize (headline)**, Add-to-knowledge,
   Translate & show, Draft reply. (`.ics` lands with the meeting agent in
   Phase 2.)

---

## Related

- [`FEATURES.md`](FEATURES.md) §7 — feature-contract entry.
- [`ARCHITECTURE.md`](ARCHITECTURE.md) §7.5 — auth / future Graph (NAA).
- Building blocks already shipped: `useOutlookItem.ts`
  (`getReadItemAsFile`, `getImageAttachments`), `synaplan-client.ts`
  (`fileUpload`, `translate`, `summarise`, `classify`, `draftReply`,
  `extractMeetingTimes`).
