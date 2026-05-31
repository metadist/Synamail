# Synamail Mail Routes — Design (per-email AI agent triggers)

**Status:** Design / RFC. Not yet implemented. Owner: Synamail. Companion to
[`FEATURES.md`](FEATURES.md) and [`ARCHITECTURE.md`](ARCHITECTURE.md).

> **Read this first — terminology.** A "**Mail Route**" here is a Synamail
> automation: **WHEN** a mail matches conditions **THEN** run one or more
> AI/agent actions on it. This is **NOT** Synaplan's _Synapse Routing_
> (`BSELECTION_RULES`), which only picks a system prompt/topic for the AI and is
> a separate, server-side concern (edited by `RuleEditor.vue`). Do not conflate
> the two. If anything in this repo blurs them, this doc wins for the
> _agent-trigger_ meaning.

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

| Action                    | What it does                                                                                                       | Feasibility                                                                                                                                                                               | Reuses                                                           |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| **Add to knowledge**      | Upload body + attachments to a RAG group, vectorized                                                               | ✅ ready                                                                                                                                                                                  | `fileUpload`, `getReadItemAsFile`, `getImageAttachments`         |
| **Translate & show**      | Translate body to a target language, render in pane                                                                | ✅ ready                                                                                                                                                                                  | `translate`                                                      |
| **Summarise & show**      | Bulleted summary in pane                                                                                           | ✅ ready                                                                                                                                                                                  | `summarise`                                                      |
| **Classify / label**      | Tag the mail (category), show a pill, optional notify                                                              | ✅ ready                                                                                                                                                                                  | `classify`                                                       |
| **Draft reply**           | Generate a reply and open the compose form                                                                         | ✅ ready                                                                                                                                                                                  | `draftReply` + `displayReplyForm`                                |
| **Suggest meeting times** | Extract requested times, **read calendar free/busy**, propose 3 slots, optionally create appointment / draft reply | ⚠️ **partial** — extraction + `displayNewAppointmentForm` exist; **reading calendar availability needs Graph (`Calendars.Read`) via an auth token / NAA** (see ARCHITECTURE §7.5 phase 2) | `extractMeetingTimes`, `displayNewAppointmentForm` (create only) |
| **Forward / notify**      | Forward to an address or post a Synaplan note                                                                      | later                                                                                                                                                                                     | —                                                                |

**The meeting agent is the one with a real capability gap:** we can already
detect proposed times and _create_ an appointment, but **"look into the
calendar and suggest 3 free times" requires reading the user's free/busy**,
which the add-in cannot do today without a Graph token. That's a dedicated
sub-project (Graph auth / Nested App Auth). Phase it separately.

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
  (`FEATURES.md §5`). It stays as-is. **Decision needed:** keep it visible,
  rename it to "Synaplan topics" to avoid confusion with Mail Routes, or
  deprioritise it.
- **Meeting-times + appointment** (`extractMeetingTimes`,
  `displayNewAppointmentForm`) — become the building blocks of the
  `suggestMeetingTimes` action (plus the calendar-read gap).
- **`getReadItemAsFile` / `getImageAttachments` / `fileUpload`** — building
  blocks of the `addToKnowledge` action.

---

## 9. Phased plan

- **Phase 0 — spike (0.5d):** pinning + `ItemChanged` support on target clients;
  scope the Graph calendar-read path for the meeting agent.
- **Phase 1 — Routes engine + simple actions:** data model + "Mail Routes"
  Settings UI (builder: conditions + actions) + trigger + `addToKnowledge`,
  `translateShow`, `summariseShow`, `classifyLabel`, `draftReply`. Deterministic
  - `language`/`category`/`intent` AI conditions.
- **Phase 2 — Meeting agent:** Graph free/busy read → `suggestMeetingTimes`
  proposing 3 open slots; create appointment / draft reply from a pick.
- **Phase 3 — hands-off + server-side (optional):** event-based launch event;
  optional Synaplan `InboundEmailHandler` for 24/7 capture.

---

## 10. Open questions

1. **Naming** — call the existing Synapse Routing UI "Synaplan topics" so
   "Routes" unambiguously means Mail Routes?
2. **Calendar read** — accept the Graph/NAA dependency for the meeting agent, or
   ship the meeting action as "propose times the sender suggested + create
   appointment" first (no availability check) and add free/busy later?
3. **Builder UX** — fixed condition/action dropdowns (v1) vs free-form natural
   language ("if Oliver asks for a meeting…") parsed into a route by AI?
4. Which 3–4 actions are the must-haves for Phase 1?

---

## Related

- [`FEATURES.md`](FEATURES.md) §7 — feature-contract entry.
- [`ARCHITECTURE.md`](ARCHITECTURE.md) §7.5 — auth / future Graph (NAA).
- Building blocks already shipped: `useOutlookItem.ts`
  (`getReadItemAsFile`, `getImageAttachments`), `synaplan-client.ts`
  (`fileUpload`, `translate`, `summarise`, `classify`, `draftReply`,
  `extractMeetingTimes`).
