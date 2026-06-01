import { describe, expect, it } from 'vitest'
import {
  findConflict,
  isoToEpoch,
  routesForSender,
  senderMatches,
  senderMatchesPattern,
  slotsOverlap,
} from '@shared/mail-routes/match'
import type { MailRoute } from '@shared/mail-routes/types'

describe('senderMatchesPattern', () => {
  it('matches a full address case-insensitively', () => {
    expect(senderMatchesPattern('Oliver@Acme.com', 'oliver@acme.com')).toBe(true)
    expect(senderMatchesPattern('other@acme.com', 'oliver@acme.com')).toBe(false)
  })

  it('matches a domain pattern with or without the leading @', () => {
    expect(senderMatchesPattern('anyone@acme.com', '@acme.com')).toBe(true)
    expect(senderMatchesPattern('anyone@acme.com', 'acme.com')).toBe(true)
    expect(senderMatchesPattern('anyone@other.com', '@acme.com')).toBe(false)
  })

  it('does not match a look-alike domain suffix', () => {
    expect(senderMatchesPattern('a@acme.com.evil.com', '@acme.com')).toBe(false)
  })

  it('never matches empty inputs', () => {
    expect(senderMatchesPattern('', 'acme.com')).toBe(false)
    expect(senderMatchesPattern('a@acme.com', '')).toBe(false)
  })
})

describe('senderMatches', () => {
  it('matches when any pattern matches', () => {
    expect(senderMatches('a@acme.com', ['@nope.com', '@acme.com'])).toBe(true)
    expect(senderMatches('a@acme.com', ['@nope.com'])).toBe(false)
    expect(senderMatches(undefined, ['@acme.com'])).toBe(false)
  })
})

describe('routesForSender', () => {
  const base = { senders: ['@acme.com'], name: 'r', durationMinutes: 30 } as const
  const routes: MailRoute[] = [
    { ...base, id: '1', kind: 'meeting', enabled: true },
    { ...base, id: '2', kind: 'meeting', enabled: false },
    { id: '3', kind: 'meeting', enabled: true, name: 'no-senders', senders: [], durationMinutes: 30 },
  ]

  it('keeps only enabled routes with a matching, non-empty sender list', () => {
    const got = routesForSender(routes, 'x@acme.com')
    expect(got.map((r) => r.id)).toEqual(['1'])
  })
})

describe('calendar overlap', () => {
  it('isoToEpoch returns null for bad input', () => {
    expect(isoToEpoch('nope')).toBeNull()
    expect(isoToEpoch('2026-06-03T15:00:00Z')).toBeTypeOf('number')
  })

  it('slotsOverlap uses half-open intervals (touching ≠ overlap)', () => {
    expect(slotsOverlap({ start: 0, end: 10 }, { start: 5, end: 15 })).toBe(true)
    expect(slotsOverlap({ start: 0, end: 10 }, { start: 10, end: 20 })).toBe(false)
  })

  it('findConflict returns the first conflicting event or null', () => {
    const events = [
      { start: 100, end: 200, subject: 'A' },
      { start: 300, end: 400, subject: 'B' },
    ]
    expect(findConflict({ start: 150, end: 160 }, events)?.subject).toBe('A')
    expect(findConflict({ start: 250, end: 290 }, events)).toBeNull()
  })
})
