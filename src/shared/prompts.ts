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

export const classify = (categories: string[]): string =>
  `You are an email triage assistant. Classify the email into exactly one of: ` +
  `${categories.join(', ')}. Return a JSON object: ` +
  `{"category": "<one of the above>", "confidence": <0..1>, "reasoning": "<short>"}. ` +
  `No prose, no markdown, just the JSON.`

/**
 * Categorize an email against user-defined categories (each "name": meaning).
 * `clarify` is optional free-text guidance. The model must pick exactly one
 * listed name, or "none" if nothing fits.
 */
export const categorize = (categoryList: string, clarify?: string): string =>
  `You assign exactly ONE category to an email, choosing from this list ` +
  `(name: meaning):\n${categoryList}\n` +
  (clarify ? `Extra guidance from the user: ${clarify}\n` : '') +
  `Pick the single best-fitting category by its meaning. If none fit, use ` +
  `"none". Return a JSON object: {"category": "<exact name from the list, or ` +
  `none>", "confidence": <0..1>, "reasoning": "<short>"}. JSON only — no prose, ` +
  `no markdown fences.`

export const ask = (): string =>
  `You are answering follow-up questions about a specific email. Stay grounded ` +
  `in the email content; if asked something the email doesn't cover, say so.`

export const simpleChat = (): string =>
  `You are Synaplan, a helpful assistant living inside the user's Outlook. ` +
  `Answer clearly and concisely. No email is loaded unless the user pastes its ` +
  `content into the conversation.`

export const newMail = (lang: string): string =>
  `You are an email-writing assistant. From the user's description, compose a ` +
  `brand-new email in language code "${lang}". Return a single JSON object: ` +
  `{"subject": "<concise subject line>", "htmlBody": "<HTML body using <p> ` +
  `paragraphs and <br> breaks; no doctype, no <html>, no <head>>"}. ` +
  `Output the JSON only — no markdown fences, no commentary.`

export const compose = (): string =>
  `You are an email-writing assistant. Generate a complete email body matching ` +
  `the user's intent. Return well-formed HTML — paragraphs in <p>, line breaks ` +
  `via <br>. No doctype, no <html>, no <head>.`

export const improveSelection = (): string =>
  `Rewrite the user's selected text to be clearer and more polished without ` +
  `changing its meaning. Preserve the original language. Return only the ` +
  `rewritten text.`

export const shortenSelection = (): string =>
  `Shorten the user's selected text. Preserve the meaning, the tone, and the ` +
  `language. Return only the shortened text.`

export const ragInsertContext = (snippet: string, citation: string): string =>
  `Reference snippet from the user's knowledge base:\n${snippet}\n\nCitation: ${citation}`

// Sprint 3 additions (Tier 1 from the Copilot-parity check).
export const extractActionItems = (lang: string): string =>
  `Extract action items from the email as a markdown checklist in language ` +
  `code "${lang}". Each item starts with "- [ ] ". Output the checklist only.`

export const sentiment = (): string =>
  `Classify the email's overall sentiment and urgency. Return JSON: ` +
  `{"sentiment": "positive" | "neutral" | "negative", ` +
  `"urgency": "low" | "medium" | "high", "reasoning": "<short>"}. JSON only.`

export const preSendCheck = (lang: string): string =>
  `You are a writing coach. Review the user's draft email and return JSON: ` +
  `{"tone": "<one word>", "clarity": <1-5>, ` +
  `"issues": ["<short issue>"], "missingAttachmentHint": <bool>, ` +
  `"sensitiveData": ["<short hint>"]}. ` +
  `Answer in language code "${lang}". JSON only.`

export const outOfOffice = (lang: string): string =>
  `Generate a polite out-of-office reply in language code "${lang}". Include ` +
  `placeholders {{return_date}} and {{backup_contact}} the user can fill in. ` +
  `Return HTML body content only.`

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
