/**
 * Email body cleaning for knowledge-base ingestion.
 *
 * Raw email text pulled from Outlook (`body.getAsync('text')`) is noisy: it
 * carries tracking URLs, zero-width / control garbage, opaque routing codes,
 * decoration rules, and the entire quoted reply history duplicated below the
 * new message. Feeding that straight into RAG dilutes the vectors and pollutes
 * search results, so we strip the obvious noise before uploading.
 *
 * This is intentionally conservative — it removes things that are almost never
 * meaningful prose (long tracking URLs, 60+ char opaque tokens, control chars,
 * Outlook's `-----Original Message-----` quoted tail) and normalises
 * whitespace, but it never tries to guess which sentences are "important". The
 * goal is cleaner vectors, not summarisation.
 *
 * `cleanEmailText` in `synaplan-client.ts` is the lighter cleaner used for the
 * AI-prompt path (summarise / translate / ask); this one is the stronger
 * variant used only for what we persist to the knowledge base.
 */

// Zero-width spaces / joiners / BOM — common in HTML-derived plain text.
const ZERO_WIDTH_RE = /[\u200B-\u200D\u2060\uFEFF]/g

// Non-printable control chars EXCEPT tab (\t) and newline (\n). `\r` is
// normalised to `\n` before this runs, so it isn't in the class.
// eslint-disable-next-line no-control-regex
const CONTROL_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g

// Long tracking URLs / mailto links, with or without autolink `< >` brackets.
// The char class excludes `< > ` and whitespace so it stops at the closing
// bracket and never eats adjacent text.
const LONG_URL_RE = /<?\b(?:https?:\/\/|mailto:)[^\s<>]{30,}>?/gi

// Outlook / common clients delimit the quoted reply history with a rule like
// `-----Original Message-----`. Everything from there on is a duplicate of an
// email that (usually) is already in the thread, so we drop it. Localised
// variants for the languages Synamail ships are covered.
const QUOTED_REPLY_RE =
  /\n[ \t]*-{2,}[ \t]*(?:Original Message|Ursprüngliche Nachricht|Message d['’]origine|Mensaje original|Messaggio originale|Mensagem original)[ \t]*-{2,}[\s\S]*$/i

// A whole line that is nothing but decoration (rules, dot-leaders, etc.).
const DECORATION_LINE_RE = /^[ \t]*[=_~*·•\-–—]{4,}[ \t]*$/gm

// Opaque, space-free blobs of 60+ chars: base64 payloads, hashes, message-ids,
// signing tokens. Real prose words (even long German compounds) don't reach 60
// characters without a space, so this is safe.
const LONG_TOKEN_RE = /\S{60,}/g

/**
 * Clean a raw email body for knowledge-base ingestion. Returns cleaned text
 * (never `null`); an empty / whitespace-only input yields `''`.
 */
export function cleanEmailForIngest(input: string): string {
  if (!input) return ''
  let s = input.replace(/\r\n?/g, '\n')

  // 1. Drop invisible / garbled characters.
  s = s.replace(ZERO_WIDTH_RE, '').replace(CONTROL_RE, '')

  // 2. Cut the quoted reply history (before URL stripping so the whole tail
  //    goes in one shot).
  s = s.replace(QUOTED_REPLY_RE, '\n')

  // 3. Strip tracking URLs and any leftover empty autolink brackets.
  s = s.replace(LONG_URL_RE, '').replace(/<\s*>/g, '')

  // 4. Strip opaque routing / signing codes and decoration-only lines.
  s = s.replace(LONG_TOKEN_RE, '')
  s = s.replace(DECORATION_LINE_RE, '')

  // 5. Normalise whitespace: trailing spaces, runs of spaces, blank-line runs.
  s = s
    .replace(/[ \t]+\n/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')

  return s.trim()
}
