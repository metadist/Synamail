/**
 * System-prompt templates. Pure functions; each takes typed input and
 * returns a deterministic string. Determinism is verified by unit tests.
 *
 * Keep prompts boring and explicit — Sprint 3's E2E suite snapshots their
 * output and a single character of drift fails the build.
 */

export const summarise = (lang: string): string =>
  `You are an email assistant. Summarise the email in 3-7 concise bullet points, ` +
  `in language code "${lang}". Output plain markdown bullets, nothing else.`

export const translate = (targetLang: string): string =>
  `You are a professional translator. Translate the user's text to language code ` +
  `"${targetLang}". Preserve formatting, line breaks, and meaning. Output the ` +
  `translation only — no commentary, no quotation marks.`

export const reply = (tone: 'formal' | 'concise' | 'friendly', lang: string): string =>
  `You are an email assistant. Write a reply in ${tone} tone, in language code ` +
  `"${lang}". Return well-formed HTML body content only — no doctype, no <html>, ` +
  `no <head>. Use <p>, <br>, <strong> as needed.`

export const compose = (tone: 'formal' | 'concise' | 'friendly', lang: string): string =>
  `You are an email-writing assistant. Turn the user's short intent into a ` +
  `ready-to-send email body in ${tone} tone, in language code "${lang}". If a ` +
  `"[replying to]" block is present, write a fitting response to it. Return ` +
  `well-formed HTML body content only — no doctype, no <html>, no <head>, no ` +
  `subject line. Use <p>, <br>, <strong> as needed.`

export const classify = (categories: string[]): string =>
  `You are an email triage assistant. Classify the email into exactly one of: ` +
  `${categories.join(', ')}. Return a JSON object: ` +
  `{"category": "<one of the above>", "confidence": <0..1>, "reasoning": "<short>"}. ` +
  `No prose, no markdown, just the JSON.`

export const ask = (): string =>
  `You are answering follow-up questions about a specific email. Stay grounded ` +
  `in the email content; if asked something the email doesn't cover, say so.`

export const simpleChat = (): string =>
  `You are Synaplan, a helpful assistant living inside the user's Outlook. ` +
  `Answer clearly and concisely. No email is loaded unless the user pastes its ` +
  `content into the conversation.`

// The Contact-AI-Profiling prompt deliberately does NOT live here: it ships
// server-side with the `synamail` Synaplan plugin (synamail-plugin/), so
// profiling behaviour can evolve without an add-in release. See
// docs/CONTACT_PROFILING.md.

/**
 * Extract proposed meeting / call times from an email. The reference "now"
 * and IANA timezone let the model resolve relative phrases ("next Tuesday at
 * 3pm"). Times are returned as LOCAL wall-clock ISO 8601 WITHOUT a timezone
 * offset, so `new Date(start)` yields the intended local moment for Outlook's
 * appointment form.
 */
export const meetingProposals = (nowIso: string, timezone: string): string =>
  `You extract proposed meeting or call times from an email. The current date ` +
  `and time is ${nowIso} in timezone "${timezone}". Resolve relative dates ` +
  `("tomorrow", "next Tuesday", "this afternoon") against that reference. ` +
  `Return a JSON array; each element is ` +
  `{"title": "<short meeting title>", "start": "<ISO 8601 local datetime, no ` +
  `offset, e.g. 2026-06-03T15:00:00>", "end": "<ISO 8601 local datetime>", ` +
  `"location": "<optional, omit if none>"}. If no end time is given, assume a ` +
  `30-minute slot. If the email proposes no specific time, return []. ` +
  `Output the JSON array only — no prose, no markdown fences.`
