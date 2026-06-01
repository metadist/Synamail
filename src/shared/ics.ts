/**
 * Minimal RFC-5545 (iCalendar) VEVENT builder.
 *
 * Pure and dependency-free so it runs identically in the Office WebView, in
 * Vitest, and in a plain browser. Produces a single-event calendar the user
 * can attach to a reply (an invite the sender can accept) or import.
 *
 * Scope is deliberately small: a meeting request with start/end, summary,
 * optional location/description, an organizer and a single attendee. We emit
 * `METHOD:REQUEST` so Outlook treats the file as an invitation rather than a
 * plain calendar export.
 */

export interface IcsEventInput {
  /** Stable unique id. Callers should pass a generated one for idempotency. */
  uid: string
  /** Local wall-clock ISO 8601 WITHOUT offset, e.g. `2026-06-03T15:00:00`. */
  startIso: string
  /** Local wall-clock ISO 8601 WITHOUT offset. */
  endIso: string
  summary: string
  location?: string
  description?: string
  /** Organizer email (typically the Outlook user). */
  organizerEmail?: string
  /** Invitee email (typically the sender we're replying to). */
  attendeeEmail?: string
  /** `DTSTAMP` reference; defaults to now. Injectable for deterministic tests. */
  nowIso?: string
}

/**
 * Format a local wall-clock ISO datetime as an iCalendar "floating" local
 * time (`YYYYMMDDTHHMMSS`, no trailing `Z`). Floating local time avoids
 * timezone-conversion surprises: the event lands at the wall-clock time the
 * email proposed, in whatever timezone the recipient opens it.
 */
export function toIcsLocalDate(iso: string): string {
  // Accept both "2026-06-03T15:00:00" and a Date-parseable value; normalise to
  // the literal Y/M/D/H/M/S components without applying any offset shift.
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/)
  if (m) {
    const [, y, mo, d, h, mi, s] = m
    return `${y}${mo}${d}T${h}${mi}${s ?? '00'}`
  }
  // Fallback: derive components from a Date in local time.
  const dt = new Date(iso)
  if (Number.isNaN(dt.getTime())) {
    throw new Error(`ics: unparseable datetime "${iso}"`)
  }
  const pad = (n: number): string => String(n).padStart(2, '0')
  return (
    `${dt.getFullYear()}${pad(dt.getMonth() + 1)}${pad(dt.getDate())}` +
    `T${pad(dt.getHours())}${pad(dt.getMinutes())}${pad(dt.getSeconds())}`
  )
}

/** Escape a value for an iCalendar text property (RFC 5545 §3.3.11). */
export function escapeIcsText(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\r?\n/g, '\\n')
}

/**
 * Fold a content line to 75 octets per RFC 5545 §3.1 using CRLF + a single
 * leading space on continuation lines. We approximate "octets" with JS string
 * length, which is correct for the ASCII-heavy output we emit.
 */
function foldLine(line: string): string {
  if (line.length <= 75) return line
  const parts: string[] = []
  let rest = line
  parts.push(rest.slice(0, 75))
  rest = rest.slice(75)
  while (rest.length > 74) {
    parts.push(' ' + rest.slice(0, 74))
    rest = rest.slice(74)
  }
  if (rest.length > 0) parts.push(' ' + rest)
  return parts.join('\r\n')
}

/** Build a complete VCALENDAR document (CRLF-delimited) with one VEVENT. */
export function buildIcsEvent(input: IcsEventInput): string {
  const dtstamp = toIcsLocalDate(input.nowIso ?? new Date().toISOString())
  const lines: string[] = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Synaplan//Synamail//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    `UID:${escapeIcsText(input.uid)}`,
    `DTSTAMP:${dtstamp}`,
    `DTSTART:${toIcsLocalDate(input.startIso)}`,
    `DTEND:${toIcsLocalDate(input.endIso)}`,
    `SUMMARY:${escapeIcsText(input.summary)}`,
  ]
  if (input.location) lines.push(`LOCATION:${escapeIcsText(input.location)}`)
  if (input.description) lines.push(`DESCRIPTION:${escapeIcsText(input.description)}`)
  if (input.organizerEmail) lines.push(`ORGANIZER:mailto:${input.organizerEmail}`)
  if (input.attendeeEmail) {
    lines.push(
      `ATTENDEE;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:${input.attendeeEmail}`,
    )
  }
  lines.push('STATUS:CONFIRMED', 'SEQUENCE:0', 'END:VEVENT', 'END:VCALENDAR')
  return lines.map(foldLine).join('\r\n')
}

/** UTF-8 base64 of an ICS document, for `fileUpload`/attachment payloads. */
export function icsToBase64(ics: string): string {
  return globalThis.btoa(unescape(encodeURIComponent(ics)))
}

/** A `data:` URL for the ICS, usable as a download link or preview. */
export function icsToDataUrl(ics: string): string {
  return `data:text/calendar;charset=utf-8;base64,${icsToBase64(ics)}`
}
