import { describe, expect, it } from 'vitest'
import { titleWithSubject, truncateTitle } from '@shared/chat-title'

describe('truncateTitle', () => {
  it('leaves short titles intact', () => {
    expect(truncateTitle('Summarise: Q3 plan')).toBe('Summarise: Q3 plan')
  })

  it('collapses whitespace and truncates with an ellipsis', () => {
    const long = 'A'.repeat(80)
    const out = truncateTitle(long)
    expect(out.length).toBe(60)
    expect(out.endsWith('…')).toBe(true)
  })
})

describe('titleWithSubject', () => {
  it('uses Prefix: subject when a subject is present', () => {
    expect(titleWithSubject('Summarise', 'Invoice #42')).toBe('Summarise: Invoice #42')
  })

  it('falls back to intent/question when there is no subject', () => {
    expect(titleWithSubject('Compose', undefined, 'Invite Alice to lunch')).toBe(
      'Invite Alice to lunch',
    )
  })

  it('uses the prefix alone when nothing else is available', () => {
    expect(titleWithSubject('Translate')).toBe('Translate')
  })
})
