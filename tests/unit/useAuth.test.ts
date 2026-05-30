import { afterEach, describe, expect, it, vi } from 'vitest'
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
  afterEach(() => vi.unstubAllEnvs())

  it('targets <baseUrl>/addin/connect with a same-origin relay redirect (real flow)', () => {
    // Force the real flow regardless of the dev env / .env.local on the box.
    vi.stubEnv('VITE_DEV_MOCK_AUTH', 'false')
    const u = new URL(buildDialogUrl('https://web.synaplan.com', 'abc'))
    expect(u.pathname).toBe('/addin/connect')
    expect(u.searchParams.get('state')).toBe('abc')
    expect(u.searchParams.get('label')).toBe('Outlook Add-in')
    const redirect = u.searchParams.get('redirect')
    expect(redirect).toBeTruthy()
    expect(redirect).toContain('/src/dialog/auth-relay.html')
  })

  it('uses the local mock relay in dev when mock auth is not disabled', () => {
    vi.stubEnv('VITE_DEV_MOCK_AUTH', '')
    const u = new URL(buildDialogUrl('https://localhost:5174', 'abc'))
    expect(u.pathname).toContain('/src/dialog/auth-relay.html')
    expect(u.searchParams.get('mock')).toBe('1')
    expect(u.searchParams.get('state')).toBe('abc')
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
