# Synamail — GUI Definitions

UI specification for the Synamail Outlook add-in. Treat this as the contract for `Sprint 2 (GUI + sideload)` and the asset list for `Sprint 4 (AppSource)`.

> Companion docs: [`../docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md) (technical), [`../docs/FEATURES.md`](../docs/FEATURES.md) (feature-by-feature contract).

## 1. Visual language

| Setting        | Value                                                                                                  |
| -------------- | ------------------------------------------------------------------------------------------------------ |
| Taskpane width | Standard Outlook taskpane (320 px nominal, scales to 450 px in pinned view).                           |
| Theme          | **Fluent UI design tokens** via CSS variables. Light + Dark mode driven by Office UI theme.            |
| Typography     | Segoe UI Variable (Windows), -apple-system (Mac), system-ui fallback.                                  |
| Density        | "Compact" Fluent density — matches Outlook's reading pane.                                             |
| Iconography    | Outline-style 16/20 px icons in Office Fluent System Icons.                                            |
| Accent colour  | Synaplan brand (consistent with the `synaplan/frontend/` SPA). Hover/pressed states use Fluent tokens. |
| Accessibility  | WCAG 2.1 AA, full keyboard nav, visible focus rings, high-contrast theme tested.                       |
| Strings        | All UI text via `vue-i18n`. **Always update `en.json` AND `de.json` together** (Synaplan house rule).  |

## 2. Routing — which view shows when

```
no apiKey in roamingSettings              →  SignIn.vue
apiKey present + Office.context.mailbox.item.itemType = "message" + reading pane (.subject available) →  ReadMode.vue
apiKey present + itemType = "message"     + compose pane               →  ComposeMode.vue
user navigates from any view              →  Settings.vue
user navigates from Settings              →  RuleEditor.vue
user navigates from ReadMode contact pill →  ContactKnowledgeBase.vue
401 from any API call                     →  clear apiKey → SignIn.vue
```

A top-level `App.vue` chooses the right view; sub-views push/pop within the taskpane (no full router needed — a small `view` ref in a Pinia store is enough).

## 3. Views

### 3.1 `SignIn.vue` (first run)

```
┌──────────────────────────────┐
│ [Synamail logo]              │
│                              │
│ Synamail                     │
│ Your Synaplan workspace in   │
│ Outlook.                     │
│                              │
│ ┌────────────────────────┐   │
│ │  Sign in to Synaplan   │   │   ← primary button
│ └────────────────────────┘   │
│                              │
│ Use a self-hosted instance → │   ← link, opens Settings.vue with base-URL editor focused
│                              │
│ Privacy · Terms · Support    │   ← footer links
└──────────────────────────────┘
```

- Single button → `useAuth.openSignInDialog()` → `Office.context.ui.displayDialogAsync(baseUrl + '/addin/connect?state=<nonce>&label=Outlook+Add-in')`.
- Footer links open in the user's default browser via `Office.context.ui.openBrowserWindow` (where supported) or a `target=_blank` anchor.
- Error toast at the bottom of the view if the dialog is cancelled, the state nonce mismatches, or the bridge page returns an error.

### 3.2 `ReadMode.vue` (selected email in reading pane)

```
┌──────────────────────────────┐
│ Synamail   ⚙ Settings        │   ← top bar
├──────────────────────────────┤
│ Subject: Q3 invoice          │
│ From: alice@example.com      │
│ Contact: alice@example.com ▾ │   ← contact pill, opens ContactKnowledgeBase
├──────────────────────────────┤
│ [ Summarise ]                │
│ [ Translate ▾ ]              │   ← target-language picker inline
│ [ Draft reply ▾ ]            │   ← tone + language popover
│ [ Classify ]                 │
│ [ Save to knowledge base ]   │
├──────────────────────────────┤
│ Result area                  │   ← summary / translation / classification result
│ (scrolls)                    │
├──────────────────────────────┤
│ Ask about this email…        │   ← anchored input, full chat thread above
│ [ send ]                     │
└──────────────────────────────┘
```

Action affordances:

- **Translate ▾** — dropdown with: auto, en, de, fr, es, it, zh, ar.
- **Draft reply ▾** — popover with tone (formal / concise / friendly) and language picker. On submit, calls `messages/send` then `mailbox.item.displayReplyForm({ htmlBody })`.
- **Save to knowledge base** — opens a modal with the group picker (see §3.7).
- **Contact pill** — shows sender by default; tap the chevron to switch to a recipient (`To` / `Cc` list). Tapping the pill body navigates to `ContactKnowledgeBase.vue`.
- **Ask** — anchored chat input keyed by `mailbox.item.conversationId`; history scrolls in the result area.

States: loading spinner per action, error toast, disabled when `apiKey` missing or `body.getAsync` not ready.

### 3.3 `ComposeMode.vue` (writing or replying)

```
┌──────────────────────────────┐
│ Synamail   ⚙ Settings        │
├──────────────────────────────┤
│ Compose mode                 │
│                              │
│ Intent: ________________     │   ← textbox
│ [ Draft from prompt ]        │
├──────────────────────────────┤
│ Selection actions            │
│ [ Improve ] [ Shorten ]      │
│ [ Translate ▾ ]              │
├──────────────────────────────┤
│ Insert from knowledge base   │
│ ⌕ search…                    │
│ ───  result list  ───        │
└──────────────────────────────┘
```

- Selection buttons are disabled until `body.getSelectedDataAsync` returns non-empty.
- Insert-from-RAG list shows snippet + group badge + score; click → `setSelectedDataAsync` with a citation footer.

### 3.4 `Settings.vue` (configuration window)

```
┌──────────────────────────────┐
│ ← Back     Settings          │
├──────────────────────────────┤
│ Signed in as                 │
│ demo@synaplan.test           │
│ on https://web.synaplan.com  │
│                              │
│ [ Sign out ]                 │   ← calls DELETE /api/v1/apikeys/{keyId}
├──────────────────────────────┤
│ Synaplan instance            │
│ ⓘ Edit before signing in to  │
│   use a self-hosted server.  │
│ [ web.synaplan.com         ] │   ← editable, only when signed out
├──────────────────────────────┤
│ Preferences                  │
│ Default RAG group:    [ ▾ ]  │
│ Language:             [ ▾ ]  │   ← auto / en / de
│ Auto-detect contact   [ on ] │
│ group on save                │
├──────────────────────────────┤
│ Advanced                     │
│ Email routing rules        → │   ← navigates to RuleEditor.vue
│ Reset roaming settings       │
└──────────────────────────────┘
```

### 3.5 `RuleEditor.vue` (RULE integration — Synapse Routing rules)

```
┌──────────────────────────────┐
│ ← Settings   Routing rules   │
├──────────────────────────────┤
│ Topic: billing               │   ← topic selector, dropdown
│                              │
│ Rules (Tier-0 matchers)      │
│  • subject contains "invoice"│  [✕]
│  • from contains "@billing." │  [✕]
│                              │
│ [ + Add rule ]               │
├──────────────────────────────┤
│ Test against current email   │
│ [ Run dry-run preview ]      │
│ → would route to: billing    │
├──────────────────────────────┤
│ ⓘ Rules are evaluated before │
│ AI classification. See the   │
│ Synapse Routing docs.        │
└──────────────────────────────┘
```

- Topic list comes from `GET /api/v1/prompts` (verified against `PromptController.php`).
- "Run dry-run preview" calls `POST /api/v1/admin/synapse/dry-run` — **admin-only**. For non-admin users the button is hidden and an inline hint says "Routing preview requires an admin account on your Synaplan instance".
- "+ Add rule" depends on whether `PromptController` exposes a user-scoped write endpoint (audit lives in `docs/FEATURES.md` §5.4). If absent, the view is read-only in v1.

### 3.6 `ContactKnowledgeBase.vue` (search by sender / recipient)

```
┌──────────────────────────────┐
│ ← Back  alice@example.com ▾  │   ← contact picker (sender + recipients)
├──────────────────────────────┤
│ ⌕ Search alice's knowledge   │
│                              │
│ Recent emails (in group)     │
│ • Q3 invoice — 2026-05-10    │
│ • Renewal terms — 2026-04-29 │
│ • Onboarding intro — 2026-03 │
├──────────────────────────────┤
│ [ Save current email to      │
│   alice's knowledge base ]   │
│                              │
│ [ Ask about alice ]          │   ← opens chat with group as RAG scope
└──────────────────────────────┘
```

- The contact picker lists the sender first, then each recipient. Switching contact reloads the view.
- "Save current email" calls upload + process with `group=contact:alice@example.com`, creating the group on first use.
- "Ask about alice" opens a chat in the same view (push-style navigation) with the contact group as the RAG scope hint.

### 3.7 Group-picker modal (shared component)

Opened from the "Save to knowledge base" button in ReadMode, and from `ContactKnowledgeBase.vue`.

```
┌──────────────────────────────┐
│ Save to knowledge base       │
├──────────────────────────────┤
│ Suggested                    │
│ ◯ contact:alice@example.com  │
│ ◯ contact:bob@example.com    │
│                              │
│ All groups                   │
│ ◯ work-notes                 │
│ ◯ project-x                  │
│ ◯ + Create new group…        │
├──────────────────────────────┤
│ Processing level             │
│ ◯ Extract only (fastest)     │
│ ◯ Extract + Vectorize        │
│ ◯ Full analysis              │
├──────────────────────────────┤
│ Attachments to include       │
│ ☑ invoice.pdf                │
│ ☑ contract.docx              │
│ ☐ signature.png              │
├──────────────────────────────┤
│ [ Cancel ]  [ Save ]         │
└──────────────────────────────┘
```

## 4. Reusable components

| Component                   | Used in                                                       |
| --------------------------- | ------------------------------------------------------------- |
| `ActionButton.vue`          | Every view — Fluent-styled primary/secondary buttons.         |
| `Spinner.vue`               | Every action with a network call.                             |
| `Toast.vue`                 | Global success / error notifications.                         |
| `LanguagePicker.vue`        | Translate buttons in ReadMode + ComposeMode.                  |
| `TonePicker.vue`            | Draft reply popover.                                          |
| `GroupPickerModal.vue`      | Save-to-RAG flow.                                             |
| `ContactPill.vue`           | ReadMode, ContactKnowledgeBase.                               |
| `ChatThread.vue`            | Ask in ReadMode, "Ask about contact" in ContactKnowledgeBase. |
| `RagResultList.vue`         | Insert-from-RAG in ComposeMode, ContactKnowledgeBase search.  |
| `ProcessingLevelPicker.vue` | Group picker modal.                                           |

## 5. Asset list

### 5.1 In-add-in icons (required by manifest)

| File                  | Size    | Used for                        |
| --------------------- | ------- | ------------------------------- |
| `assets/icon-16.png`  | 16×16   | Ribbon (small)                  |
| `assets/icon-32.png`  | 32×32   | Ribbon (medium), context menu   |
| `assets/icon-64.png`  | 64×64   | Outlook on iPad                 |
| `assets/icon-80.png`  | 80×80   | Ribbon (large, classic Outlook) |
| `assets/icon-128.png` | 128×128 | High-DPI ribbon                 |

### 5.2 AppSource store assets

| File                                   | Size          | Used for                                |
| -------------------------------------- | ------------- | --------------------------------------- |
| `assets/store/hero-256.png`            | 256×256       | AppSource listing thumbnail             |
| `assets/store/hero-512.png`            | 512×512       | AppSource listing hero                  |
| `assets/store/screenshot-signin.png`   | 1366×768      | Store screenshot — SignIn               |
| `assets/store/screenshot-read.png`     | 1366×768      | Store screenshot — ReadMode             |
| `assets/store/screenshot-compose.png`  | 1366×768      | Store screenshot — ComposeMode          |
| `assets/store/screenshot-settings.png` | 1366×768      | Store screenshot — Settings             |
| `assets/store/screenshot-contact.png`  | 1366×768      | Store screenshot — ContactKB            |
| `assets/store/screenshot-rules.png`    | 1366×768      | Store screenshot — RuleEditor           |
| `assets/store/screencast.mp4`          | ≤ 60 s, 1080p | Showcase video (summarise → save → ask) |
| `assets/store/copy.md`                 | —             | Title, descriptions, keywords (en + de) |

All assets exist as **placeholders** at the end of Sprint 1 (empty files referenced in the asset list) and as **final artwork** by Step 4.3.

## 6. Design verification checklist

To be ticked at the end of Sprint 2 (visual review) and Sprint 4 (store-ready review):

- [ ] Every view renders in light + dark Office themes.
- [ ] Tab order is logical in every view; visible focus ring everywhere.
- [ ] All strings come from `en.json` and `de.json`; no hardcoded text.
- [ ] No layout breakage at 320 px width or at the pinned 450 px width.
- [ ] All icons are visible against the Outlook ribbon in light + dark mode.
- [ ] Loading + empty + error states exist for every list and every action.
- [ ] High-contrast theme passes axe-core automated check.
- [ ] All screenshots in `assets/store/` show real, populated content (no Lorem Ipsum).
