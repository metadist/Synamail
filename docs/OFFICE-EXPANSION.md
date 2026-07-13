# Office Expansion — Synaoffice (Word + PowerPoint Add-ins)

> **Status:** Estimation / planning note (July 2026). Not yet started.
>
> Assessment of how the add-in know-how built in **Synamail** (the Outlook add-in,
> `github.com/metadist/Synamail`) can be expanded into a new repository, **Synaoffice**,
> providing plugins for Microsoft **Word** and **PowerPoint** in the same way, with new
> host-specific features.

## Starting point

Synamail's architecture: Vue 3 + TypeScript + Vite + Office.js, OAuth via the Office
Dialog API against the existing Synaplan login, a Zod-validated HTTP client generated
from the Synaplan OpenAPI spec, and a full local CI gate (`make ci-local`) mirroring
GitHub Actions. All features run on existing Synaplan endpoints — no Synaplan-side API
changes were needed.

## How well does the know-how transfer?

Roughly **60–70% of Synamail carries over directly**, because most of it is
host-agnostic:

- **Auth flow** — the Office Dialog API + `/addin/connect` bridge page + `state`-nonce
  flow works identically in Word and PowerPoint. The Synaplan-side bridge needs
  **zero changes**. This was the hardest, most regression-prone part of Synamail
  (see Synamail's `docs/AUTH_FLOW.md`), and it comes for free.
- **Synaplan client** — `synaplan-client.ts`, Zod schemas generated from the OpenAPI
  spec, and the same endpoints (`messages/send`, `files/upload`, `rag/search`,
  `apikeys`).
- **Toolchain and process** — Vite/Vue setup, the `tokens.css` design system, vue-i18n
  with six locales, Makefile targets, `.githooks` + `ci-local` gate, manifest
  validation, Vitest/Playwright test infrastructure, and the whole AppSource
  submission checklist.

## What is genuinely new

1. **Different Office.js API model.** Outlook uses the direct `Office.context.mailbox`
   API; Word and PowerPoint use the batch/proxy model (`Word.run(...)`,
   `PowerPoint.run(...)` with `context.sync()`). This is the main new learning curve —
   the equivalent of Synamail's `useOutlookItem` composable has to be written twice,
   once per host.
2. **Settings storage.** `Office.context.roamingSettings` (where Synamail keeps the
   API key) is **Outlook-only**. Word/PowerPoint need
   `Office.context.document.settings` (per-document — wrong for credentials) or
   `localStorage` / partitioned storage instead. This touches `useRoamingSettings` and
   `useAuth` and needs a deliberate design decision early.
3. **Host-specific features**, for example:
   - **Word:** rewrite / improve / translate selection, generate sections from a
     prompt, insert content from RAG with citations, summarise the document, proofread
     against the knowledge base.
   - **PowerPoint:** generate slides / outline from a prompt or RAG content, write
     speaker notes, summarise a deck, rewrite slide text for tone / length.
4. **Manifest:** one manifest _can_ target both hosts (`Document` + `Presentation`),
   which is the recommended path — one repo, one AppSource listing "Synaoffice", one
   deployment.

## Recommended shape

One new `Synaoffice` repo cloned structurally from Synamail: a shared core
(`src/shared/`, auth, client, i18n, styles, dialog relay) plus `src/word/` and
`src/powerpoint/` taskpane entry points.

Shared code should be **copied, not extracted into an npm package** for now — three
repos sharing a package adds release overhead that is not yet worth it.

## Effort estimate

| Phase              | Scope                                                                                                                | Effort                                         |
| ------------------ | -------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------- |
| 0 — Scaffold       | Repo from Synamail skeleton, dual-host manifest, sideload into Word + PowerPoint, port auth (incl. storage decision) | ~1 week                                        |
| 1 — Word MVP       | Selection rewrite, summarise, generate-from-prompt, insert-from-RAG                                                  | 2–3 weeks                                      |
| 2 — PowerPoint MVP | Slide / outline generation, speaker notes, deck summary                                                              | 2–3 weeks                                      |
| 3 — Polish + store | Unified manifest, store assets, submission iterations                                                                | 1–2 weeks dev + 4–8 weeks Microsoft turnaround |

Overall: **roughly 6–8 weeks of development** to a solid dual-host MVP, versus the
~10+ weeks Synamail took from zero. The discount comes almost entirely from reusing
the auth flow, the HTTP client, and the CI/process scaffolding.

## Recommended first step (de-risking spike)

Before committing to the plan, run a **1–2 day spike**: sideload a hello-world
dual-host manifest into Word and PowerPoint (web + desktop) and run the existing
dialog auth flow end-to-end there. The dialog behaves slightly differently per host,
and that is where Synamail historically had its most painful regressions.
