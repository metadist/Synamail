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

  it('classify lists the allowed categories verbatim', () => {
    const p = prompts.classify(['billing', 'support', 'general'])
    expect(p).toContain('billing, support, general')
  })

  it('extractActionItems renders the markdown checklist instruction', () => {
    expect(prompts.extractActionItems('en')).toContain('- [ ]')
  })

  it('preSendCheck produces JSON-only contract', () => {
    expect(prompts.preSendCheck('en')).toContain('JSON only')
  })
})
