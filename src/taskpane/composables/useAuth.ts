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
  const u = new URL('/addin/connect', baseUrl)
  u.searchParams.set('state', state)
  u.searchParams.set('label', 'Outlook Add-in')
  return u.toString()
}

/**
 * Open the sign-in dialog. Resolves with the validated payload, or rejects
 * with an Error explaining the failure (cancel, bad state nonce, dialog
 * couldn't open).
 */
export function openSignInDialog(opts: UseAuthOptions = {}): Promise<SignInPayload> {
  const baseUrl = opts.baseUrl ?? defaultBaseUrl()
  const state = generateStateNonce()
  const url = opts.dialogUrl ?? buildDialogUrl(baseUrl, state)

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

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          if (settled) return
          settled = true
          try {
            const payload = JSON.parse(arg.message ?? '{}') as SignInPayload
            if (!payload.state || payload.state !== state) {
              dialog.close()
              reject(new Error('State nonce mismatch — sign-in rejected'))
              return
            }
            if (!payload.apiKey || !payload.email || !payload.baseUrl) {
              dialog.close()
              reject(new Error('Sign-in payload missing required fields'))
              return
            }
            dialog.close()
            resolve(payload)
          } catch (err) {
            dialog.close()
            reject(err instanceof Error ? err : new Error(String(err)))
          }
        })

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (settled) return
          settled = true
          reject(new Error('Sign-in cancelled'))
        })
      },
    )
  })
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
      const { createSynaplanClient } = await import('@shared/synaplan-client')
      const client = createSynaplanClient({
        baseUrl: settings.baseUrl,
        apiKey: settings.apiKey,
      })
      await client.revokeApiKey(settings.keyId)
    } catch {
      // Sprint 3 surfaces this as a toast. In Sprint 2 we tolerate failure
      // and proceed with local clear so the user is never stuck.
    }
  }
  await clearSettings()
  hydrateAuthState()
}

export function defaultBaseUrl(): string {
  return loadSettings()?.baseUrl ?? 'https://web.synaplan.com'
}
