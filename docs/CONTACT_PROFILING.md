# Synamail Contact AI Profiling — Design (rolling relationship memory)

**Status:** Design-first. Not yet implemented. This doc is the authoritative
spec for the **Contact AI Profiling** feature (formerly "Contact knowledge
base"). It supersedes the contact-KB description in
[`FEATURES.md`](FEATURES.md) §4 for _what the feature does_; that file's coverage
matrix still wins for _release timing_. Companion to
[`MAIL_ROUTES.md`](MAIL_ROUTES.md) and [`ARCHITECTURE.md`](ARCHITECTURE.md).
Owner: Synamail.

> **Read this first — terminology.** A "**profile**" here is a Synamail-owned,
> evolving summary of one **contact** (a person, identified by email; or an
> **org**, identified by domain). It is built from the emails you exchange **plus
> manual note snippets you add** (calls, meetings, verbal promises). It is **not**
> Synaplan's generic RAG search, and **not** Synapse Routing. The contact's corpus
> still lives in a `contact:<email>` RAG group; profiling adds a durable,
> recomputed **state object** on top of that corpus.

---

## Decisions locked (2026-06-07)

1. **Data architecture = Option B, "local-first, server-as-truth."** No Synaplan
   platform changes — **no new tables, no migrations, no new endpoints.** We
   reuse the existing `files`/RAG API only. (See §5 for the rejected options.)
2. **Sync the inputs, not the summary.** A profile is _derived_ data:
   `profile = f(emails + notes)`. The **inputs** (emails + manual notes) live in
   the `contact:<email>` RAG group in the user's Synaplan workspace — one shared
   source of truth across all the user's machines. Each machine **recomputes**
   the summary from that synced corpus.
3. **Roaming settings are a cache, never the store.** `Office.roamingSettings`
   already syncs small data across the user's Outlook clients, but its ~32 KB
   quota means it only holds a small, disposable snapshot cache — the corpus is
   authoritative.
4. **Pin a best-effort snapshot for consistent wording.** When a profile is
   (re)generated, store the compact snapshot JSON as a single artifact in the
   contact group so all machines show the _same text_ instead of three slightly
   different AI phrasings. Any machine may refresh it.
5. **A first-class server-side profile entity (Option C) is explicitly
   deferred.** It is the only "big platform change" and is not needed for the
   value below.

---

## 1. The problem and the core idea

Today the view is pure **retrieval** — it runs `rag/search` over the
`contact:<email>` group and lists hits. That answers "what did we say about X?"
but not "where does this relationship stand, and what do I owe whom?".

A **rolling profile** is a durable, evolving state object that is **recomputed**
as new signals arrive, so opening a contact instantly tells you things like:

- _"You haven't mailed in 6 weeks. Your last exchange was friendly but
  distanced."_
- _"You promised to deliver a demo, but never mailed back."_
- _"They owe you a reply since 3 June."_

The feature has three layers:

1. **Corpus** — every email + every note about the contact (raw material, lives
   in Synaplan as the `contact:<email>` RAG group). **The synced backbone.**
2. **Profile snapshot** — a compact, structured summary the user sees instantly
   (derived; cached locally, optionally pinned in the group — §5).
3. **Recompute pipeline** — deterministic signals + an AI synthesis pass that
   regenerates the snapshot from the corpus on each interaction.

---

## 2. Why manual note snippets are mandatory

The most valuable relationship context is **off-channel**: a phone call, a
hallway chat, a verbal "yes." Without a way to inject _"Demo delivered over
phone"_, the profile is blind to half of reality and will confidently surface a
**wrong** open commitment ("you promised a demo") that the user already
fulfilled.

So note capture is a first-class input, not a nice-to-have:

- A **"+ Add note"** composer stores a short text artifact (e.g. _"Demo
  delivered over phone — 7 Jun"_) into the **same** `contact:<email>` group,
  tagged `type: note`.
- Notes become first-class corpus: searchable, vectorized, fed into profiling
  exactly like emails — **and synced across machines for free**, because they
  live in Synaplan.
- A note can **resolve** an open commitment (see §6).

---

## 3. What gets captured (the inputs)

| Input                 | Source                                     | How                                                                     | Status         |
| --------------------- | ------------------------------------------ | ----------------------------------------------------------------------- | -------------- |
| Emails (in/out)       | Outlook item                               | `getReadItemAsFile` → `fileUpload` into `contact:<email>` with metadata | building block |
| Manual notes/snippets | User-typed                                 | upload short text file, `type: note`, vectorized                        | new            |
| Deterministic signals | Outlook `senderHistory` (EWS) + item dates | last in/out date, cadence, response latency, who owes the next reply    | building block |
| Contact identity      | Outlook headers                            | name + email; org derived from domain                                   | building block |

**Deterministic signals are computed in code, never left to the LLM** — dates,
cadence and "who owes a reply" must be exact, not hallucinated. They are recomputed
locally per machine (cheap, no sync needed).

---

## 4. Data model

A structured object per contact. The snapshot is small enough to cache and to
pin as a single artifact in the contact group; the corpus stays in Synaplan.

```ts
interface ContactProfile {
  email: string // lower-cased; the group is `contact:<email>`
  name?: string
  org?: string // derived from the email domain
  firstSeen: string // ISO
  lastInbound?: string // ISO — last mail FROM the contact
  lastOutbound?: string // ISO — last mail TO the contact
  owedBy?: 'me' | 'them' // who should send the next reply
  tone?: string // e.g. "friendly but distanced"
  stage?: string // e.g. "early sales / evaluating"
  summary: string // the rolling narrative (1–3 sentences)
  commitments: Commitment[] // open + recently resolved
  notes: ProfileNote[] // manual snippets (most-recent first)
  profiledAt: string // ISO — "as of" timestamp for the snapshot
  profiledOn?: string // machine/host hint that last refreshed it
}

interface Commitment {
  id: string
  text: string // "Deliver a demo"
  owedBy: 'me' | 'them'
  since: string // ISO — first seen
  status: 'open' | 'resolved'
  resolvedBy?: string // note id that closed it (if any)
}

interface ProfileNote {
  id: string
  text: string // "Demo delivered over phone"
  date: string // ISO
  fileId?: number // the Synaplan file the note was stored as
}
```

---

## 5. Storage — local-first, server-as-truth (Option B, locked)

The decision (2026-06-07): **lives locally, syncs via Synaplan, with no platform
changes.** Concretely, three tiers:

- **Inputs → Synaplan** (`contact:<email>` RAG group). Emails + notes. The
  **synced source of truth**, shared by every machine signed into the same
  Synaplan workspace. Written with the existing `fileUpload`; **no new tables or
  endpoints.**
- **Derived summary → recomputed per machine**, cached in
  `Office.roamingSettings` as a **bounded LRU** (e.g. last 10–20 contacts).
  Roaming already syncs small data across the user's Outlook clients, but its
  **~32 KB total** quota means it must stay a disposable cache, not the store.
- **Pinned snapshot (optional, recommended) → one JSON artifact in the contact
  group** (e.g. `profile.json`, `type: profile`). So all machines show identical
  wording rather than three slightly different AI phrasings. Any machine can
  refresh it; the newest `profiledAt` wins.

> **Why this needs no platform change.** The corpus, notes, and even the pinned
> snapshot are all just **files in a RAG group** — the `files`/RAG API already
> exists. Synaplan never needs a "profiles" concept. The summary is reproducible
> from the corpus, so losing the local cache or the pinned file is non-fatal: the
> next open regenerates it.

### Options considered (for the record)

- **A — Roaming only.** Zero platform change and syncs via Exchange, but the
  ~32 KB cap limits how many profiles fit, it's tied to one mailbox, and it's
  invisible outside Outlook. _Rejected as the store_ (kept as the cache tier).
- **B — Local-first, server-as-truth (CHOSEN).** Inputs in Synaplan, summary
  derived + cached + optionally pinned. No platform change; syncs across all the
  user's machines via the shared workspace; cache is disposable so the quota
  never bites.
- **C — First-class profile entity in Synaplan** (DB table + endpoints).
  Queryable and shareable beyond Outlook, but it is the **big platform change**
  (migrations, controllers, GDPR-delete plumbing) and is **not needed** for the
  value here. _Deferred._

### Sync follows the workspace, not the mailbox

Because the corpus is keyed to the **Synaplan workspace**, sync follows the
workspace. All of the user's machines sign into the same Synaplan account, so
they stay in sync even if the underlying Outlook mailboxes differ — more robust
than roaming alone.

---

## 6. The profiling engine

A new client method, e.g. `profileContact(input)`, that produces a
`ContactProfile`:

1. **Gather corpus** — a few **targeted** `rag/search` queries scoped to
   `contact:<email>` (recent topics, commitments/promises, tone signals), plus
   the most recent notes.
2. **Compute deterministic facts in code** — `lastInbound`, `lastOutbound`,
   cadence, `owedBy` (from the direction of the most recent message). These are
   injected into the result, not produced by the model.
3. **AI synthesis** — send the gathered snippets + the deterministic facts to
   `messages/send` in **JSON mode** (same pattern as the existing `classify`
   action) with a dedicated `profileContact` prompt in `prompts.ts`. The model
   returns `summary`, `tone`, `stage`, and a list of `commitments`.
4. **Pin + cache** — write the snapshot back as the pinned `profile.json` in the
   group (best-effort) and into the local roaming LRU.

### Commitments lifecycle (the headline behaviour)

This is what makes the profile feel intelligent instead of nagging:

- The AI extracts **promises** from the corpus ("you promised a demo").
- A later **note** (_"Demo delivered over phone"_) **resolves** the matching
  open commitment — either by the user tapping "mark resolved", or by the AI
  matching the note to the commitment on the next recompute.
- Resolved commitments are kept briefly (for context) then drop off.

---

## 7. Freshness — what "rolling" means in an add-in

An Outlook add-in is **not a background service** (same constraint as Mail
Routes, [`MAIL_ROUTES.md`](MAIL_ROUTES.md) §2). So "rolling" means **incremental
recompute on interaction**, not a daemon:

- Recompute when the user **opens** the contact, **saves an email** to it, or
  **adds a note** — and on an explicit **Refresh** button.
- On a fresh machine, the first open regenerates the summary from the synced
  corpus (or loads the pinned snapshot instantly, then refreshes).
- Always show an **"as of <time>"** line so the user knows the snapshot's age.
- A true 24/7 path (continuous server-side profiling via Synaplan's
  `InboundEmailHandler`) remains an optional later phase.

---

## 8. UI surfaces (rebuilt `ContactProfile.vue`)

- **Profile card** at the top: the rolling `summary`, freshness signal
  ("last contact 6 weeks ago — they owe you a reply"), and `tone`/`stage` chips.
- **Open-commitments list** with one-tap "mark resolved" / "add resolving note".
- **"+ Add note"** composer (the snippet capture).
- **Refresh** button + "as of" timestamp.
- Existing **search** and **"ask about this contact"** stay underneath.
- A **delete-profile** action (privacy — §9).

---

## 9. Privacy and compliance — do not skip

Profiling **identifiable people** in an EU/German product heading to AppSource
demands explicit handling:

- Data stays in the **user's own** Synaplan workspace — **no central
  aggregation** of profiles.
- **Transparency** — the user knows a profile exists and what feeds it.
- **Full delete** — purge the `contact:<email>` group (corpus + notes + pinned
  snapshot) and the cached snapshot.
- Optional **per-contact opt-out**.
- A clear line in the **privacy policy** and the **store listing**. AppSource
  reviewers will look for this.

---

## 10. Dependencies and gaps

- **Listing a group's contents.** `rag/search` requires a query and there is no
  "list everything in a group" endpoint (the same gap that produced the
  empty-query error fixed on 2026-06-07). Profiling works around it with several
  **targeted** queries; fetching the pinned `profile.json` similarly needs a
  deterministic way to read one file from the group. A Synaplan-side
  **files-by-group** listing would make both cleaner and is worth a ticket — but
  it is an _enhancement_, not a blocker (Option B still works with targeted
  queries).
- **Note / snapshot storage shape.** Notes and the pinned snapshot are stored as
  small text/JSON files in the contact group; confirm the `BFILES` metadata/tag
  fields available for `type` (`note` / `profile`) and `date` during wire-up.
- **Reusable building blocks already shipped.** `getReadItemAsFile`,
  `getImageAttachments` (`useOutlookItem.ts`); `senderHistory`
  (`useOutlookMailbox.ts`); `fileUpload`, `ragSearch`, `ragGroups`,
  `ragCreateGroup`, `chat` (`synaplan-client.ts`).

---

## 11. Phased plan

- **Phase 1 — shippable slice (recommended first):** manual **note snippets**
  (stored in Synaplan = synced) + a recomputed **profile card** (deterministic
  freshness + AI `summary`/`tone`/`stage`) + **open-commitments** list, with a
  local roaming cache. This alone delivers the "you haven't mailed in a while /
  you promised a demo" experience across all the user's machines.
- **Phase 2 — commitments lifecycle + pinned snapshot:** resolve-by-note
  matching, resolved-state decay, and the pinned `profile.json` for consistent
  cross-machine wording.
- **Phase 3 — org-level rollups:** aggregate `org:<domain>` profiles across the
  people at one company.
- **Phase 4 — server-side (optional, Option C):** continuous profiling via
  Synaplan's `InboundEmailHandler` and/or a first-class profile entity — only if
  cross-tool access is ever needed.

---

## 12. Open questions

1. **Note storage** — text file per note in the contact group (proposed) vs a
   single appended `notes.md` artifact. Per-note files are simpler to vectorize
   and delete individually.
2. **Snapshot cache size** — how many contacts to keep in the roaming LRU before
   the ~32 KB quota bites. Start small (10–20), measure.
3. **Pinned-snapshot conflicts** — newest `profiledAt` wins is the simple rule;
   confirm that's acceptable when two machines refresh close together.
4. **Recompute cost** — how aggressively to recompute on open vs only on new
   activity, given Synaplan rate limits. Cache + "Refresh" button is the
   conservative default.

---

## Related

- [`FEATURES.md`](FEATURES.md) §4 (contact KB today) and §8 (profiling summary).
- [`MAIL_ROUTES.md`](MAIL_ROUTES.md) §2 — the shared "no background service"
  constraint.
- [`ARCHITECTURE.md`](ARCHITECTURE.md) — auth, API surface, data minimisation.
