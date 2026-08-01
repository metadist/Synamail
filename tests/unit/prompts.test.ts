import { describe, expect, it } from 'vitest'
import * as prompts from '@shared/prompts'

describe('prompts', () => {
  it('summarise renders the target language code into the task directive', () => {
    expect(prompts.summarise('de')).toContain('"de"')
    expect(prompts.summarise('de')).not.toMatch(/you are/i)
  })

  it('translate renders the target language', () => {
    expect(prompts.translate('fr')).toContain('language code "fr"')
    expect(prompts.translate('fr')).not.toMatch(/you are/i)
  })

  it('reply combines tone and language deterministically', () => {
    const a = prompts.reply('concise', 'en')
    const b = prompts.reply('concise', 'en')
    expect(a).toBe(b)
    expect(a).toMatch(/concise/)
    expect(a).toMatch(/"en"/)
    expect(a).not.toMatch(/you are/i)
  })

  it('compose combines tone and language and asks for HTML body only', () => {
    const a = prompts.compose('friendly', 'de')
    const b = prompts.compose('friendly', 'de')
    expect(a).toBe(b)
    expect(a).toMatch(/friendly/)
    expect(a).toContain('"de"')
    expect(a).toMatch(/HTML/i)
    expect(a).toMatch(/no subject line/i)
    expect(a).not.toMatch(/you are/i)
  })

  it('compose writes in the intent language and treats the code as a fallback', () => {
    const p = prompts.compose('formal', 'en')
    // Follows the user's intent language, like chat…
    expect(p).toMatch(/same language/i)
    expect(p).toMatch(/\[intent\]/)
    // …and the language code is only the fallback.
    expect(p).toMatch(/fall back to language code "en"/i)
  })

  it('classify lists the allowed categories verbatim', () => {
    const p = prompts.classify(['billing', 'support', 'general'])
    expect(p).toContain('billing, support, general')
    expect(p).not.toMatch(/you are/i)
  })

  it('ask grounding line stays short and persona-free', () => {
    expect(prompts.ask()).toMatch(/email context/i)
    expect(prompts.ask()).not.toMatch(/you are/i)
  })

  it('meetingProposals embeds the reference now + timezone and asks for a JSON array', () => {
    const p = prompts.meetingProposals('2026-06-01T09:00:00', 'Europe/Berlin')
    expect(p).toContain('2026-06-01T09:00:00')
    expect(p).toContain('Europe/Berlin')
    expect(p).toMatch(/JSON array/i)
    expect(p).not.toMatch(/you are/i)
  })
})
