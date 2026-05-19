import { describe, expect, it } from 'vitest'
import {
  clearSettings,
  loadSettings,
  patchSettings,
  saveSettings,
} from '@/taskpane/composables/useRoamingSettings'

describe('useRoamingSettings', () => {
  it('round-trips save → load', async () => {
    await saveSettings({
      apiKey: 'sk_x',
      keyId: 7,
      email: 'a@b.test',
      baseUrl: 'https://example.test',
    })
    expect(loadSettings()).toEqual({
      apiKey: 'sk_x',
      keyId: 7,
      email: 'a@b.test',
      baseUrl: 'https://example.test',
    })
  })

  it('returns null when empty', () => {
    expect(loadSettings()).toBeNull()
  })

  it('clear removes the entry', async () => {
    await saveSettings({ apiKey: 'sk_x', keyId: 1, email: 'a@b.test', baseUrl: 'https://x' })
    await clearSettings()
    expect(loadSettings()).toBeNull()
  })

  it('patch merges fields', async () => {
    await saveSettings({ apiKey: 'sk_x', keyId: 1, email: 'a@b.test', baseUrl: 'https://x' })
    await patchSettings({ language: 'de' })
    expect(loadSettings()?.language).toBe('de')
    expect(loadSettings()?.apiKey).toBe('sk_x')
  })

  it('patch throws when nothing is stored yet', () => {
    expect(() => patchSettings({ language: 'de' })).toThrow()
  })
})
