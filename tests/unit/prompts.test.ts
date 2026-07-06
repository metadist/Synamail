import { describe, expect, it } from 'vitest'
import * as prompts from '@shared/prompts'

describe('prompts', () => {
  it('summarise renders the target language code into the system prompt', () => {
    expect(prompts.summarise('de')).toContain('"de"')
  })

  it('translate renders the target language', () => {
    expect(prompts.translate('fr')).toContain('"fr"')
  })

  it('reply combines tone and language deterministically', () => {
    const a = prompts.reply('concise', 'en')
    const b = prompts.reply('concise', 'en')
    expect(a).toBe(b)
    expect(a).toMatch(/concise/)
    expect(a).toMatch(/"en"/)
  })

  it('compose combines tone and language and asks for HTML body only', () => {
    const a = prompts.compose('friendly', 'de')
    const b = prompts.compose('friendly', 'de')
    expect(a).toBe(b)
    expect(a).toMatch(/friendly/)
    expect(a).toContain('"de"')
    expect(a).toMatch(/HTML/i)
    expect(a).toMatch(/no subject line/i)
  })

  it('classify lists the allowed categories verbatim', () => {
    const p = prompts.classify(['billing', 'support', 'general'])
    expect(p).toContain('billing, support, general')
  })

  it('meetingProposals embeds the reference now + timezone and asks for a JSON array', () => {
    const p = prompts.meetingProposals('2026-06-01T09:00:00', 'Europe/Berlin')
    expect(p).toContain('2026-06-01T09:00:00')
    expect(p).toContain('Europe/Berlin')
    expect(p).toMatch(/JSON array/i)
  })
})
