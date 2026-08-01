/**
 * Task-directive templates for Outlook AI actions. Pure functions; each takes
 * typed input and returns a deterministic string. Determinism is verified by
 * unit tests.
 *
 * These are short format/behaviour constraints prepended to the user-visible
 * content (intent, email body, question). They must NOT use a "You are…"
 * persona preamble — that string is persisted as the Synaplan user turn and
 * pollutes History (see metadist/Synamail#56).
 *
 * Keep prompts boring and explicit — Sprint 3's E2E suite snapshots their
 * output and a single character of drift fails the build.
 */

export const summarise = (lang: string): string =>
  `Summarise the email in 3-7 concise bullet points, ` +
  `in language code "${lang}". Output plain markdown bullets, nothing else.`

export const translate = (targetLang: string): string =>
  `Translate the following text to language code "${targetLang}". ` +
  `Preserve formatting, line breaks, and meaning. Output the translation ` +
  `only — no commentary, no quotation marks.`

export const reply = (tone: 'formal' | 'concise' | 'friendly' | 'detailed', lang: string): string =>
  `Write a reply in ${tone} tone, in language code "${lang}". Return ` +
  `well-formed HTML body content only — no doctype, no <html>, no <head>. ` +
  `Use <p>, <br>, <strong> as needed.`

export const compose = (
  tone: 'formal' | 'concise' | 'friendly' | 'detailed',
  lang: string,
): string =>
  `Turn the short [intent] into a ready-to-send email body in ${tone} tone. ` +
  `Write in the SAME language as the "[intent]" text; only if that language ` +
  `cannot be determined, fall back to language code "${lang}". If a ` +
  `"[replying to]" block is present, write a fitting response to it. Return ` +
  `well-formed HTML body content only — no doctype, no <html>, no <head>, ` +
  `no subject line. Use <p>, <br>, <strong> as needed.`

export const classify = (categories: string[]): string =>
  `Classify the email into exactly one of: ${categories.join(', ')}. ` +
  `Return a JSON object: ` +
  `{"category": "<one of the above>", "confidence": <0..1>, "reasoning": "<short>"}. ` +
  `No prose, no markdown, just the JSON.`

/**
 * Grounding line for follow-ups about an open email. Kept short so History
 * stays readable; the email body is paired under `[email context]`.
 */
export const ask = (): string =>
  `Answer the question using the email context below. If the email does not ` + `cover it, say so.`

/**
 * Extract proposed meeting / call times from an email. The reference "now"
 * and IANA timezone let the model resolve relative phrases ("next Tuesday at
 * 3pm"). Times are returned as LOCAL wall-clock ISO 8601 WITHOUT a timezone
 * offset, so `new Date(start)` yields the intended local moment for Outlook's
 * appointment form.
 */
export const meetingProposals = (nowIso: string, timezone: string): string =>
  `Extract proposed meeting or call times from the email. The current date ` +
  `and time is ${nowIso} in timezone "${timezone}". Resolve relative dates ` +
  `("tomorrow", "next Tuesday", "this afternoon") against that reference. ` +
  `Return a JSON array; each element is ` +
  `{"title": "<short meeting title>", "start": "<ISO 8601 local datetime, no ` +
  `offset, e.g. 2026-06-03T15:00:00>", "end": "<ISO 8601 local datetime>", ` +
  `"location": "<optional, omit if none>"}. If no end time is given, assume a ` +
  `30-minute slot. If the email proposes no specific time, return []. ` +
  `Output the JSON array only — no prose, no markdown fences.`
