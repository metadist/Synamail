import { describe, expect, it, vi } from 'vitest'
import {
  buildDialogUrl,
  generateStateNonce,
  hydrateAuthState,
  isSignedIn,
  signIn,
  signOut,
} from '@/taskpane/composables/useAuth'
import { saveSettings } from '@/taskpane/composables/useRoamingSettings'

describe('generateStateNonce', () => {
  it('produces 32-char hex strings', () => {
    const n = generateStateNonce()
    expect(n).toMatch(/^[0-9a-f]{32}$/)
  })

  it('returns different values across calls', () => {
    expect(generateStateNonce()).not.toBe(generateStateNonce())
  })
})

describe('buildDialogUrl', () => {
  it('appends state and label', () => {
    const u = new URL(buildDialogUrl('https://web.synaplan.com', 'abc'))
    expect(u.pathname).toBe('/addin/connect')
    expect(u.searchParams.get('state')).toBe('abc')
    expect(u.searchParams.get('label')).toBe('Outlook Add-in')
  })
})

describe('hydrateAuthState', () => {
  it('reflects saved settings into reactive refs', async () => {
    await saveSettings({ apiKey: 'sk', keyId: 1, email: 'a@b.test', baseUrl: 'https://x' })
    hydrateAuthState()
    expect(isSignedIn.value).toBe(true)
  })
})

describe('signIn (dialog interaction)', () => {
  it('rejects when dialog open fails', async () => {
    const fakeOffice = (
      globalThis as unknown as {
        Office: { context: { ui: { displayDialogAsync: ReturnType<typeof vi.fn> } } }
      }
    ).Office
    fakeOffice.context.ui.displayDialogAsync = vi.fn((_url, _opts, cb) =>
      cb({ status: 'failed', error: { message: 'blocked' } }),
    )
    await expect(signIn({ baseUrl: 'https://x' })).rejects.toThrow(/blocked/)
  })
})

describe('signOut', () => {
  it('clears state without calling the network when revokeRemote=false', async () => {
    await saveSettings({ apiKey: 'sk', keyId: 1, email: 'a@b.test', baseUrl: 'https://x' })
    await signOut()
    expect(isSignedIn.value).toBe(false)
  })
})
