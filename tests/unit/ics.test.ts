import { describe, expect, it } from 'vitest'
import { buildIcsEvent, escapeIcsText, icsToBase64, toIcsLocalDate } from '@shared/ics'

describe('toIcsLocalDate', () => {
  it('formats a local wall-clock ISO without applying any offset', () => {
    expect(toIcsLocalDate('2026-06-03T15:00:00')).toBe('20260603T150000')
  })

  it('defaults seconds to 00 when omitted', () => {
    expect(toIcsLocalDate('2026-06-03T15:30')).toBe('20260603T153000')
  })

  it('throws on an unparseable value', () => {
    expect(() => toIcsLocalDate('not-a-date')).toThrow()
  })
})

describe('escapeIcsText', () => {
  it('escapes commas, semicolons, backslashes and newlines', () => {
    expect(escapeIcsText('a, b; c\\d\ne')).toBe('a\\, b\\; c\\\\d\\ne')
  })
})

describe('buildIcsEvent', () => {
  const ics = buildIcsEvent({
    uid: 'evt-1@synamail',
    startIso: '2026-06-03T15:00:00',
    endIso: '2026-06-03T15:30:00',
    summary: 'Project sync',
    location: 'Online',
    organizerEmail: 'me@firma.de',
    attendeeEmail: 'oliver@acme.com',
    nowIso: '2026-06-01T09:00:00',
  })

  it('emits a single VEVENT inside a VCALENDAR with METHOD:REQUEST', () => {
    expect(ics).toContain('BEGIN:VCALENDAR')
    expect(ics).toContain('METHOD:REQUEST')
    expect(ics).toContain('BEGIN:VEVENT')
    expect(ics).toContain('END:VEVENT')
    expect(ics).toContain('END:VCALENDAR')
    expect(ics.match(/BEGIN:VEVENT/g)?.length).toBe(1)
  })

  it('carries the start, end, summary, organizer and attendee', () => {
    expect(ics).toContain('DTSTART:20260603T150000')
    expect(ics).toContain('DTEND:20260603T153000')
    expect(ics).toContain('SUMMARY:Project sync')
    expect(ics).toContain('ORGANIZER:mailto:me@firma.de')
    expect(ics).toContain('ATTENDEE;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:oliver@acme.com')
    expect(ics).toContain('UID:evt-1@synamail')
  })

  it('uses CRLF line endings', () => {
    expect(ics).toContain('\r\n')
  })

  it('omits optional lines when not provided', () => {
    const minimal = buildIcsEvent({
      uid: 'u',
      startIso: '2026-06-03T15:00:00',
      endIso: '2026-06-03T15:30:00',
      summary: 'x',
      nowIso: '2026-06-01T09:00:00',
    })
    expect(minimal).not.toContain('LOCATION:')
    expect(minimal).not.toContain('ORGANIZER:')
    expect(minimal).not.toContain('ATTENDEE')
  })

  it('icsToBase64 round-trips to the original text', () => {
    const b64 = icsToBase64('BEGIN:VCALENDAR')
    expect(decodeURIComponent(escape(globalThis.atob(b64)))).toBe('BEGIN:VCALENDAR')
  })
})
