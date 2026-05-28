import { describe, expect, it } from 'vitest'
import {
  clearSettings,
  getChatIdForConversation,
  getLastRagGroupId,
  loadSettings,
  patchSettings,
  saveSettings,
  setChatIdForConversation,
  setLastRagGroupId,
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

  it('chatId helpers round-trip per conversation and do nothing when unauthenticated', async () => {
    expect(getChatIdForConversation('thread-1')).toBeUndefined()
    // No-op when no settings exist yet.
    await setChatIdForConversation('thread-1', 99)
    expect(loadSettings()).toBeNull()

    await saveSettings({
      apiKey: 'sk_x',
      keyId: 1,
      email: 'a@b.test',
      baseUrl: 'https://x',
    })
    await setChatIdForConversation('thread-1', 42)
    await setChatIdForConversation('thread-2', 99)

    expect(getChatIdForConversation('thread-1')).toBe(42)
    expect(getChatIdForConversation('thread-2')).toBe(99)
    expect(loadSettings()?.apiKey).toBe('sk_x')
  })

  it('lastRagGroupId helpers round-trip', async () => {
    expect(getLastRagGroupId()).toBeUndefined()
    await saveSettings({
      apiKey: 'sk_x',
      keyId: 1,
      email: 'a@b.test',
      baseUrl: 'https://x',
    })
    await setLastRagGroupId('contact:alice@example.com')
    expect(getLastRagGroupId()).toBe('contact:alice@example.com')
  })
})
