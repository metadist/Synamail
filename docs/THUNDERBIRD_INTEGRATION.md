# Thunderbird Integration — Concept & Plan

**Status: concept / not started.** This document captures the first plan for bringing
Synamail to Mozilla Thunderbird. Nothing described here is implemented yet; no code
should land against this plan until it has been reviewed and turned into sprint items.

## 1. Summary

Thunderbird cannot run Office.js add-ins, so Synamail cannot simply be "enabled" for
Thunderbird — it needs a second host integration. Thunderbird's extension model is the
**MailExtension**: a WebExtension (same foundation as Firefox add-ons) with
mail-specific APIs.

The core finding: **the majority of Synamail is already host-agnostic.** The Synaplan
client, prompts, markdown renderer, the entire Vue 3 UI (the four home boxes + chat),
and all six locales carry over unchanged. The Outlook coupling is concentrated in a
handful of composables and the manifest/commands layer. Roughly 60–70 % of the code is
reusable as-is.

The plan is a **host adapter** abstraction with two implementations (Office.js and
Thunderbird MailExtension APIs) and two build targets from this one repo.

## 2. Platform comparison

| Concern             | Outlook (today)                                        | Thunderbird (planned)                                                            |
| ------------------- | ------------------------------------------------------ | -------------------------------------------------------------------------------- |
| Extension model     | Office.js add-in                                       | MailExtension (WebExtension)                                                     |
| Manifest            | `manifest.xml` (+ generated unified JSON)              | `manifest.json` (WebExtension format)                                            |
| Main UI surface     | Taskpane (persistent side pane)                        | **Spaces tab** (TB 115+, full HTML page) + action popups                         |
| Read the open email | `Office.context.mailbox.item` + async getters          | `messenger.messageDisplay.getDisplayedMessage()`, `messenger.messages.getFull()` |
| Write a draft       | `body.setAsync` on the compose item                    | `messenger.compose.setComposeDetails()`                                          |
| Open a reply        | `displayReplyForm({ htmlBody })`                       | `messenger.compose.beginReply(messageId, { body })`                              |
| Open a new email    | `displayNewMessageForm({ htmlBody })`                  | `messenger.compose.beginNew({ body })`                                           |
| New vs. reply       | Heuristic (compose body already has quoted text)       | Explicit — the compose details carry the compose type                            |
| Settings storage    | `roamingSettings` (Exchange-synced, encrypted at rest) | `browser.storage.local` (or `storage.sync`)                                      |
| Auth dialog         | `displayDialogAsync` + `auth-relay.html`               | `windows.create()` popup + `runtime.onMessage`                                   |
| API calls / CORS    | Hosted origin (`addin.synaplan.com`)                   | Extension origin; CORS bypassed via `host_permissions`                           |
| Distribution        | AppSource / sideload manifest                          | addons.thunderbird.net (ATN) / XPI sideload                                      |

## 3. What is reused unchanged

The host-agnostic "brain" of Synamail:

- **`src/shared/`** — `synaplan-client.ts` (all API calls incl. SSE streaming),
  `prompts.ts`, `markdown.ts`, `types.ts`.
- **The Vue 3 UI** — `Home.vue` with the four function boxes (`EmailWritingBox`,
  `SummarizeBox`, `KnowledgeBaseBox`, chat), `ChatThread`, `SaveToRagDialog`,
  `Toast`, `ActionButton`, `Settings`/`SignIn` views (with adapter-backed storage
  and auth), design tokens and `app.css`.
- **i18n** — all six locales (`en`, `de`, `fr`, `es`, `it`, `pt`).
- **Web Speech dictation** in the chat composer. Web Speech support in Thunderbird
  is limited; the mic already hides itself on unsupported hosts, so no change needed.

## 4. What needs a Thunderbird counterpart

The Outlook coupling is concentrated in these files:

| Outlook-specific piece                                                                                                                       | Thunderbird counterpart                                                                                 |
| -------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| `src/taskpane/composables/useOutlookItem.ts` (item snapshot, mode detection, `setComposeBody`, reply/new-message helpers, attachment access) | New `useThunderbirdItem` built on `messageDisplay`, `messages`, `compose`, `messages.listAttachments()` |
| `src/taskpane/composables/useRoamingSettings.ts`                                                                                             | `browser.storage.local` wrapper with the same function signatures                                       |
| `src/taskpane/composables/useAuth.ts` + `src/dialog/auth-relay.*`                                                                            | Popup window to the same `/addin/connect` page; result returned via `postMessage`/`runtime` messaging   |
| `manifest.xml`, `src/commands/commands.ts`                                                                                                   | `manifest.json`, background script, `message_display_action` / `compose_action` / space registration    |

## 5. Proposed architecture: the `MailHost` adapter

Introduce one interface that captures everything the UI needs from the mail client,
and inject the host implementation at boot:

```
src/
├── hosts/
│   ├── types.ts            ← MailHost interface (the contract)
│   ├── office/             ← wraps today's useOutlookItem / roaming / auth code
│   └── thunderbird/        ← MailExtension implementation
├── shared/                 ← unchanged (client, prompts, markdown, types)
├── taskpane/               ← rename mentally to "app": views + components,
│                             consume MailHost instead of Office.js directly
└── thunderbird/
    ├── manifest.json       ← MailExtension manifest
    └── background.ts       ← action buttons, space registration, message bus
```

Sketch of the contract (final shape decided during implementation):

```ts
interface MailHost {
  /** Reactive snapshot of the currently relevant email (read or compose). */
  currentItem(): Ref<MailItemSnapshot>
  /** Write HTML into the open draft. */
  writeDraft(html: string): Promise<boolean>
  /** Open a reply window pre-filled with HTML. */
  openReply(html: string): Promise<boolean>
  /** Open a new-message window pre-filled with HTML (+ optional subject). */
  openNewMessage(html: string, subject?: string): Promise<boolean>
  /** Persistent settings (roamingSettings vs browser.storage). */
  storage: SettingsStore
  /** Open the sign-in flow and resolve with the SignInPayload. */
  openSignInDialog(url: string): Promise<SignInPayload>
}
```

The Office implementation is a thin wrapper around the existing, tested composables —
the refactor must not change Outlook behaviour. Vitest component tests then run
against a mock `MailHost`, which also simplifies today's Office stubbing in
`tests/setup.ts`.

### Build targets

Two Vite configs from one repo:

- `make build` — today's taskpane bundle (unchanged output, `dist/`).
- `make build-thunderbird` — MailExtension bundle + `web-ext build` style XPI
  packaging (`dist-thunderbird/`).

CI builds both; the existing bundle-size budget applies per target.

## 6. Auth flow (the delicate part)

Read [`AUTH_FLOW.md`](AUTH_FLOW.md) first — the sign-in flow has regressed before and
its invariants apply here too (real round-trip, no mock mode, `state` nonce
round-trip).

Planned Thunderbird flow, reusing the existing server side:

1. Extension opens `https://<instance>/addin/connect?state=<nonce>&host=thunderbird`
   in a popup via `windows.create()`.
2. `AddinConnectView.vue` (in `synaplan/frontend`) detects `host=thunderbird` and,
   instead of `Office.context.ui.messageParent(...)`, posts the `SignInPayload` to
   the opener / redirects to an extension-owned page.
3. The extension validates the `state` nonce (same rule as Outlook) and stores the
   API key via the storage adapter.

This requires one **small cross-repo change** in `synaplan/frontend`'s
`AddinConnectView.vue` (a second response channel next to the Office relay). Same
coordination rules as the Sprint-2 bridge page — see
[`SYNAPLAN_INTEGRATION.md`](SYNAPLAN_INTEGRATION.md) and run Synaplan's house
pre-commit gate in that repo.

Fallback if the popup channel proves fragile in review: manual API-key paste in the
extension's settings page (Synaplan already lets users create keys in the web UI).

## 7. Differences that simplify or complicate

**Simpler than Outlook:**

- New-vs-reply is explicit (`compose.getComposeDetails().type`), replacing the
  quoted-body heuristic in `EmailWritingBox`.
- No Office.js boot watchdog / Safari ITP issues; the extension runs first-party.
- CORS is a manifest permission, not a server headers question.

**More complicated / open questions:**

- **UI surface choice.** A Spaces tab (TB 115+) fits the 4-box home best but is not
  tied to the currently selected message the way the taskpane is; action popups are
  message-scoped but transient. Likely answer: space for the home + chat, popup for
  the message-scoped quick actions. Needs a UX decision.
- **Storage sync.** `roamingSettings` roams with the mailbox; `browser.storage.local`
  is per-profile. Chat-id continuity across devices is lost unless we move that
  mapping server-side (possible later enhancement).
- **Minimum version.** Target Thunderbird 128 ESR (or current ESR at implementation
  time); document in the manifest's `strict_min_version`.
- **Review process.** ATN review is human and can take days; release automation must
  account for it (a second release channel to maintain).

## 8. Milestones

| #   | Milestone                | Content                                                                                              | Exit criteria                                                           |
| --- | ------------------------ | ---------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------- |
| 1   | Host adapter refactor    | Introduce `MailHost`; Office implementation wraps existing composables; UI consumes the adapter      | `make ci-local` green; zero behaviour change in Outlook                 |
| 2   | Thunderbird skeleton     | `manifest.json`, background script, space/tab opening the Vue app, TB item adapter (read-only)       | App boots in Thunderbird; open email is displayed                       |
| 3   | Auth + storage           | Popup sign-in flow (incl. the `synaplan/frontend` connect-page change), storage adapter              | Sign-in round-trip works against `web.synaplan.com` and self-hosted     |
| 4   | Feature parity           | Four boxes wired: write/reply, summarize (new compose window), knowledge base, chat with attachments | Manual test matrix passes on TB ESR (Linux/Windows/macOS)               |
| 5   | Packaging & distribution | XPI build in CI, ATN listing, release documentation                                                  | Installable signed XPI; docs updated (`ARCHITECTURE.md`, `FEATURES.md`) |

Estimated effort for milestones 1–4: **1–2 weeks focused work**, plus ATN review lead
time for milestone 5. Milestone 1 pays for itself independently: it cleans up Office
stubbing in tests and decouples the UI from the host.

## 9. Non-goals (for the first iteration)

- Contact AI Profiling (currently disabled in the taskpane too).
- Whisper-backend dictation fallback (Web Speech only, hidden when unsupported).
- Server-side chat-id roaming (accept per-profile chat continuity).
- Any other mail client (Apple Mail, etc.) — but the `MailHost` seam is the
  extension point if one ever comes up.

## 10. Decision log

| Date       | Decision                                                       |
| ---------- | -------------------------------------------------------------- |
| 2026-07-13 | Concept written; no implementation until the plan is reviewed. |
