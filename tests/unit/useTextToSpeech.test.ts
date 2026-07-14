import { describe, expect, it } from 'vitest'
import { splitForTts } from '@/taskpane/composables/useTextToSpeech'

describe('splitForTts', () => {
  it('returns nothing for empty/blank text', () => {
    expect(splitForTts('')).toEqual([])
    expect(splitForTts('   ')).toEqual([])
  })

  it('keeps short text as a single chunk', () => {
    expect(splitForTts('Hello world.')).toEqual(['Hello world.'])
  })

  it('splits long text into <= 480-char chunks on sentence boundaries', () => {
    const sentence = `${'a'.repeat(200)}. `
    const text = sentence.repeat(5).trim() // ~1000 chars across 5 sentences
    const chunks = splitForTts(text)
    expect(chunks.length).toBeGreaterThan(1)
    for (const chunk of chunks) {
      expect(chunk.length).toBeLessThanOrEqual(480)
    }
    // No text is lost (ignoring whitespace joins).
    expect(chunks.join(' ').replace(/\s+/g, '')).toBe(text.replace(/\s+/g, ''))
  })

  it('hard-splits a single sentence longer than the limit', () => {
    const text = 'b'.repeat(1000)
    const chunks = splitForTts(text)
    expect(chunks.length).toBe(3)
    for (const chunk of chunks) {
      expect(chunk.length).toBeLessThanOrEqual(480)
    }
  })
})
