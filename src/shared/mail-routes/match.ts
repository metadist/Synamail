/**
 * Deterministic matching helpers for Mail Routes.
 *
 * Pure functions only — these run as the cheap "Tier-0" gate before any AI
 * call. Keeping them side-effect-free makes them trivially unit-testable and
 * safe to run on every `ItemChanged`.
 */

import type { MailRoute } from './types'

/** Lowercase + trim an email-ish string for case-insensitive comparison. */
function normalizeAddress(value: string): string {
  return value.trim().toLowerCase()
}

/**
 * Does `from` match a single sender pattern?
 *
 * Patterns:
 *   - `oliver@acme.com` → exact address match.
 *   - `@acme.com` / `acme.com` → any address at that domain (incl. subdomains
 *     are NOT matched — exact domain only, to avoid `acme.com.evil.com`).
 *
 * Empty / malformed inputs never match.
 */
export function senderMatchesPattern(from: string, pattern: string): boolean {
  const addr = normalizeAddress(from)
  const pat = normalizeAddress(pattern)
  if (!addr || !pat) return false

  if (pat.includes('@') && !pat.startsWith('@')) {
    // Full address pattern.
    return addr === pat
  }
  // Domain pattern: strip a leading '@' if present.
  const domain = pat.startsWith('@') ? pat.slice(1) : pat
  if (!domain) return false
  const at = addr.lastIndexOf('@')
  if (at === -1) return false
  return addr.slice(at + 1) === domain
}

/** Does `from` match ANY of the route's sender patterns? */
export function senderMatches(from: string | undefined, patterns: string[]): boolean {
  if (!from) return false
  return patterns.some((p) => senderMatchesPattern(from, p))
}

/**
 * Filter enabled routes whose sender list matches the given `from`. Disabled
 * routes and routes with no sender patterns are excluded (a route with no
 * senders would match everything, which is never the intent for v1).
 */
export function routesForSender(routes: MailRoute[], from: string | undefined): MailRoute[] {
  return routes.filter((r) => r.enabled && r.senders.length > 0 && senderMatches(from, r.senders))
}

// ---------------------------------------------------------------------------
// Calendar conflict detection (rule 1)
// ---------------------------------------------------------------------------

export interface TimeSlot {
  /** Epoch ms (inclusive start). */
  start: number
  /** Epoch ms (exclusive end). */
  end: number
}

export interface CalendarEvent extends TimeSlot {
  subject: string
  /** Display name / email of the other party, when known. */
  organizer?: string
}

/** Parse a local-or-offset ISO string to epoch ms; NaN-safe (returns null). */
export function isoToEpoch(iso: string): number | null {
  const t = new Date(iso).getTime()
  return Number.isNaN(t) ? null : t
}

/** Half-open interval overlap: [aStart,aEnd) intersects [bStart,bEnd). */
export function slotsOverlap(a: TimeSlot, b: TimeSlot): boolean {
  return a.start < b.end && b.start < a.end
}

/**
 * Find the first existing calendar event that conflicts with `proposed`.
 * Returns `null` when the slot is free.
 */
export function findConflict(proposed: TimeSlot, events: CalendarEvent[]): CalendarEvent | null {
  for (const ev of events) {
    if (slotsOverlap(proposed, ev)) return ev
  }
  return null
}
