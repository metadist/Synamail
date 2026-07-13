import { describe, expect, it } from 'vitest'
import { cleanEmailForIngest } from '@shared/email-clean'

describe('cleanEmailForIngest', () => {
  it('returns empty string for empty / whitespace input', () => {
    expect(cleanEmailForIngest('')).toBe('')
    expect(cleanEmailForIngest('   \n\n  ')).toBe('')
  })

  it('leaves clean prose untouched', () => {
    expect(cleanEmailForIngest('Hello world')).toBe('Hello world')
  })

  it('removes zero-width characters and BOM', () => {
    expect(cleanEmailForIngest('a\u200Bb\uFEFFc\u2060d')).toBe('abcd')
  })

  it('removes non-printable control characters but keeps tabs/newlines', () => {
    expect(cleanEmailForIngest('a\u0000\u0007b')).toBe('ab')
    expect(cleanEmailForIngest('line1\n\tindented')).toBe('line1\n\tindented')
  })

  it('strips long tracking URLs and autolink brackets without eating neighbours', () => {
    const dirty = `Read more <https://track.example.com/${'x'.repeat(50)}>KEEPME`
    const clean = cleanEmailForIngest(dirty)
    expect(clean).toContain('Read more')
    expect(clean).toContain('KEEPME')
    expect(clean).not.toContain('track.example.com')
    expect(clean).not.toContain('http')
  })

  it('keeps short, human-meaningful URLs', () => {
    expect(cleanEmailForIngest('see https://syn.io for details')).toContain('https://syn.io')
  })

  it('cuts the Outlook quoted reply history at the separator', () => {
    const text = [
      'Thanks, that works for me.',
      '',
      '-----Original Message-----',
      'From: Bob <bob@example.com>',
      'Sent: Monday',
      'Subject: Re: plan',
      '',
      'Here is the original very long message...',
    ].join('\n')
    const clean = cleanEmailForIngest(text)
    expect(clean).toBe('Thanks, that works for me.')
    expect(clean).not.toContain('Original Message')
    expect(clean).not.toContain('bob@example.com')
  })

  it('cuts the German quoted reply history', () => {
    const text = 'Danke!\n\n-----Ursprüngliche Nachricht-----\nVon: Bob\nBetreff: Re: Plan'
    expect(cleanEmailForIngest(text)).toBe('Danke!')
  })

  it('removes long opaque routing / signing tokens', () => {
    const token = 'A1b2C3d4'.repeat(10) // 80 chars, no spaces
    const clean = cleanEmailForIngest(`Message body\nX-Token: ${token}\nEnd`)
    expect(clean).not.toContain(token)
    expect(clean).toContain('Message body')
    expect(clean).toContain('End')
  })

  it('drops decoration-only rules and collapses blank-line runs', () => {
    const text = 'Section A\n=========\n\n\n\nSection B'
    const clean = cleanEmailForIngest(text)
    expect(clean).not.toContain('====')
    expect(clean).toContain('Section A')
    expect(clean).toContain('Section B')
    expect(clean).not.toMatch(/\n{3,}/)
  })

  it('normalises CRLF and trailing spaces', () => {
    expect(cleanEmailForIngest('a  \r\nb   ')).toBe('a\nb')
  })
})
