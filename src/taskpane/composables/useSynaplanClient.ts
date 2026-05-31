/**
 * Vue composable that returns a memoised SynaplanClient bound to the
 * current roaming settings. On 401 from any call, automatically clears
 * the key and bounces back to SignIn (via a global event).
 */

import { computed } from 'vue'
import { createSynaplanClient, isApiError } from '@shared/synaplan-client'
import type { SynaplanClient } from '@shared/synaplan-client'
import { isSignedIn, signedInBaseUrl } from './useAuth'
import { loadSettings } from './useRoamingSettings'

export const AUTH_INVALIDATED_EVENT = 'synamail:auth-invalidated'

export function useSynaplanClient() {
  const client = computed<SynaplanClient | null>(() => {
    if (!isSignedIn.value) return null
    const settings = loadSettings()
    if (!settings) return null
    return createSynaplanClient({
      baseUrl: settings.baseUrl,
      apiKey: settings.apiKey,
    })
  })

  /**
   * Run a Synaplan call with automatic 401 handling.
   * Returns null on 401 so callers can render a friendly "please sign in"
   * state; throws other errors for the caller to display.
   */
  async function call<T>(fn: (c: SynaplanClient) => Promise<T>): Promise<T | null> {
    const c = client.value
    if (!c) return null
    try {
      return await fn(c)
    } catch (err) {
      if (isApiError(err) && err.status === 401) {
        window.dispatchEvent(new CustomEvent(AUTH_INVALIDATED_EVENT))
        return null
      }
      throw err
    }
  }

  return { client, call, baseUrl: signedInBaseUrl }
}
