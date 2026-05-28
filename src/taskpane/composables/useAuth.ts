/**
 * Sign-in / sign-out flow.
 *
 * Sign-in uses the Office Dialog API:
 *   1. Open `<baseUrl>/addin/connect?state=<nonce>` in a popup.
 *   2. Receive a `messageParent` payload back containing `apiKey` + metadata.
 *   3. Validate the `state` nonce.
 *   4. Persist via useRoamingSettings.
 *
 * For local development (no live bridge page) Vite serves
 * `/src/dialog/auth-relay.html` which immediately echoes a mock key back —
 * see auth-relay.ts.
 */

import { ref } from 'vue'
import { createSynaplanClient } from '@shared/synaplan-client'
import type { SignInPayload } from '@shared/types'
import { clearSettings, loadSettings, saveSettings } from './useRoamingSettings'

export interface UseAuthOptions {
  /** Override the base URL for the next sign-in (does not persist). */
  baseUrl?: string
  /** For tests. */
  dialogUrl?: string
}

interface DialogLike {
  addEventHandler: (
    eventType: Office.EventType,
    handler: (arg: { message?: string; error?: number }) => void,
  ) => void
  close: () => void
}

export const isSignedIn = ref(false)
export const signedInEmail = ref<string | null>(null)
export const signedInBaseUrl = ref<string | null>(null)

export function hydrateAuthState(): void {
  const s = loadSettings()
  isSignedIn.value = !!s
  signedInEmail.value = s?.email ?? null
  signedInBaseUrl.value = s?.baseUrl ?? null
}

export function generateStateNonce(): string {
  const arr = new Uint8Array(16)
  globalThis.crypto.getRandomValues(arr)
  return Array.from(arr, (b) => b.toString(16).padStart(2, '0')).join('')
}

export function buildDialogUrl(baseUrl: string, state: string): string {
  // Dev shortcut: in Vite dev mode point at the local mock relay so the
  // sign-in loop closes in <100 ms with no real Synaplan round-trip. The
  // mock fires Office.context.ui.messageParent same-origin from
  // https://localhost:3000, sidestepping the cross-origin messageParent
  // issue with the production https://web.synaplan.com/addin/connect
  // bridge. Set DEV_MOCK_AUTH=false in .env.local to force the real flow.
  if (import.meta.env.DEV && import.meta.env.VITE_DEV_MOCK_AUTH !== 'false') {
    const u = new URL('/src/dialog/auth-relay.html', 'https://localhost:3000')
    u.searchParams.set('state', state)
    u.searchParams.set('baseUrl', baseUrl)
    return u.toString()
  }
  const u = new URL('/addin/connect', baseUrl)
  u.searchParams.set('state', state)
  u.searchParams.set('label', 'Outlook Add-in')
  return u.toString()
}

/**
 * Open the sign-in dialog. Resolves with the validated payload, or rejects
 * with an Error explaining the failure (cancel, bad state nonce, dialog
 * couldn't open).
 *
 * Listens on TWO channels in parallel:
 *   1. Office.EventType.DialogMessageReceived — the proper Office channel.
 *   2. window 'message' events from the bridge origin — fallback for OWA,
 *      where Office's cross-domain messageParent is unreliable when the
 *      taskpane and bridge live on different origins. The bridge fires
 *      BOTH messageParent and window.opener.postMessage; whichever
 *      channel delivers first wins.
 */
export function openSignInDialog(opts: UseAuthOptions = {}): Promise<SignInPayload> {
  const baseUrl = opts.baseUrl ?? defaultBaseUrl()
  const state = generateStateNonce()
  const url = opts.dialogUrl ?? buildDialogUrl(baseUrl, state)
  const expectedOrigin = originOf(baseUrl)

  return new Promise<SignInPayload>((resolve, reject) => {
    if (typeof Office === 'undefined' || !Office.context?.ui?.displayDialogAsync) {
      reject(new Error('Office Dialog API not available — sideload the manifest first.'))
      return
    }

    Office.context.ui.displayDialogAsync(
      url,
      { height: 60, width: 40, displayInIframe: false },
      (asyncResult: Office.AsyncResult<Office.Dialog>) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(asyncResult.error?.message ?? 'Failed to open sign-in dialog'))
          return
        }
        const dialog = asyncResult.value as unknown as DialogLike
        let settled = false

        const cleanupWindowListener = (): void => {
          window.removeEventListener('message', onWindowMessage)
        }

        const finalizeError = (err: Error): void => {
          if (settled) return
          settled = true
          cleanupWindowListener()
          try {
            dialog.close()
          } catch {
            // Dialog may already be closing — ignore.
          }
          reject(err)
        }

        const finalizeSuccess = (payload: SignInPayload): void => {
          if (settled) return
          settled = true
          cleanupWindowListener()
          try {
            dialog.close()
          } catch {
            // Same as above — best-effort close.
          }
          resolve(payload)
        }

        const handlePayload = (raw: string): void => {
          let payload: SignInPayload
          try {
            payload = JSON.parse(raw) as SignInPayload
          } catch (err) {
            finalizeError(err instanceof Error ? err : new Error(String(err)))
            return
          }
          if (!payload.state || payload.state !== state) {
            finalizeError(new Error('State nonce mismatch — sign-in rejected'))
            return
          }
          if (!payload.apiKey || !payload.email || !payload.baseUrl) {
            finalizeError(new Error('Sign-in payload missing required fields'))
            return
          }
          finalizeSuccess(payload)
        }

        function onWindowMessage(event: MessageEvent): void {
          if (settled) return
          // Only accept messages from the expected bridge origin.
          if (event.origin !== expectedOrigin) return
          if (typeof event.data !== 'string') return
          handlePayload(event.data)
        }

        window.addEventListener('message', onWindowMessage)

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          if (settled) return
          handlePayload(arg.message ?? '{}')
        })

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (settled) return
          finalizeError(new Error('Sign-in cancelled'))
        })
      },
    )
  })
}

/** Extract the URL origin (scheme + host + port) for postMessage validation. */
function originOf(url: string): string {
  try {
    return new URL(url).origin
  } catch {
    return url
  }
}

export async function signIn(opts: UseAuthOptions = {}): Promise<void> {
  const payload = await openSignInDialog(opts)
  await saveSettings({
    apiKey: payload.apiKey,
    keyId: payload.keyId,
    email: payload.email,
    baseUrl: payload.baseUrl,
  })
  hydrateAuthState()
}

export async function signOut(opts: { revokeRemote?: boolean } = {}): Promise<void> {
  const settings = loadSettings()
  if (settings && opts.revokeRemote) {
    try {
      const client = createSynaplanClient({
        baseUrl: settings.baseUrl,
        apiKey: settings.apiKey,
      })
      await client.revokeApiKey(settings.keyId)
    } catch {
      // Best-effort remote revoke: tolerate failure so the user is never
      // stuck locally with stale roaming settings if Synaplan is offline.
    }
  }
  await clearSettings()
  hydrateAuthState()
}

export function defaultBaseUrl(): string {
  return loadSettings()?.baseUrl ?? 'https://web.synaplan.com'
}
