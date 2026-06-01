/**
 * Strongly-typed wrapper around `Office.context.roamingSettings`.
 *
 * Roaming settings are JSON-serialised and stored encrypted at rest by
 * Exchange, per-mailbox. We namespace everything under `synamail.*` so we
 * don't clash with other add-ins.
 *
 * Outside the Office runtime (e.g. tests, browser dev), we fall back to
 * `localStorage` so the dev loop still works.
 */

import { EMPTY_ROUTES_STATE, type MailRoutesState } from '@shared/mail-routes/types'
import type { RoamingSettings } from '@shared/types'

const KEY = 'synamail.settings'

type RoamingApi = Pick<Office.RoamingSettings, 'get' | 'set' | 'remove' | 'saveAsync'>

interface FallbackStore {
  get(key: string): unknown
  set(key: string, value: unknown): void
  remove(key: string): void
  saveAsync(cb?: (result: Office.AsyncResult<void>) => void): void
}

function isOfficeReady(): boolean {
  return (
    typeof Office !== 'undefined' &&
    Office.context !== undefined &&
    Office.context.roamingSettings !== undefined
  )
}

function fallbackStore(): FallbackStore {
  return {
    get(key) {
      try {
        const raw = globalThis.localStorage?.getItem(key)
        return raw ? JSON.parse(raw) : undefined
      } catch {
        return undefined
      }
    },
    set(key, value) {
      try {
        globalThis.localStorage?.setItem(key, JSON.stringify(value))
      } catch {
        /* no-op */
      }
    },
    remove(key) {
      try {
        globalThis.localStorage?.removeItem(key)
      } catch {
        /* no-op */
      }
    },
    saveAsync(cb) {
      cb?.({
        status: Office.AsyncResultStatus.Succeeded,
        value: undefined,
        asyncContext: undefined,
        diagnostics: undefined,
        error: undefined as unknown as Office.Error,
      } as unknown as Office.AsyncResult<void>)
    },
  }
}

function store(): RoamingApi | FallbackStore {
  return isOfficeReady() ? Office.context.roamingSettings : fallbackStore()
}

export function loadSettings(): RoamingSettings | null {
  const value = store().get(KEY) as RoamingSettings | undefined
  if (!value || typeof value !== 'object' || !('apiKey' in value)) return null
  return value
}

export function saveSettings(settings: RoamingSettings): Promise<void> {
  const s = store()
  s.set(KEY, settings)
  return new Promise((resolve, reject) => {
    s.saveAsync((res?: Office.AsyncResult<void>) => {
      if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(res?.error ?? new Error('saveAsync failed'))
    })
  })
}

export function clearSettings(): Promise<void> {
  const s = store()
  s.remove(KEY)
  return new Promise((resolve, reject) => {
    s.saveAsync((res?: Office.AsyncResult<void>) => {
      if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(res?.error ?? new Error('saveAsync failed'))
    })
  })
}

/**
 * Wipe ALL Synamail-owned roaming + local state. Removes the auth payload,
 * the preferred base URL, and any `synamail.*` keys in localStorage. Used by
 * the Settings → "Reset saved settings" button so a developer can switch
 * cleanly between dev/live servers without leftover identity baggage.
 */
export function clearAllSettings(): Promise<void> {
  const s = store()
  s.remove(KEY)
  s.remove(BASE_URL_KEY)
  try {
    const ls = globalThis.localStorage
    if (ls) {
      const stale: string[] = []
      for (let i = 0; i < ls.length; i += 1) {
        const k = ls.key(i)
        if (k && k.startsWith('synamail.')) stale.push(k)
      }
      for (const k of stale) ls.removeItem(k)
    }
  } catch {
    // localStorage may be unavailable inside the Office sandbox — ignore.
  }
  return new Promise((resolve, reject) => {
    s.saveAsync((res?: Office.AsyncResult<void>) => {
      if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(res?.error ?? new Error('saveAsync failed'))
    })
  })
}

export function patchSettings(patch: Partial<RoamingSettings>): Promise<void> {
  const current = loadSettings()
  if (!current) {
    throw new Error('No roaming settings to patch — sign in first')
  }
  return saveSettings({ ...current, ...patch })
}

// ---------------------------------------------------------------------------
// Per-conversation chat-id cache
//
// The Ask feature creates a Synaplan chat on the first turn and then sends
// follow-ups via `trackId = chatId`. To make the same Outlook thread reuse
// the same chat across taskpane reloads, we persist the mapping inside
// `roamingSettings.chats[conversationId]`.
// ---------------------------------------------------------------------------

export function getChatIdForConversation(conversationId: string): number | undefined {
  return loadSettings()?.chats?.[conversationId]
}

export async function setChatIdForConversation(
  conversationId: string,
  chatId: number,
): Promise<void> {
  const current = loadSettings()
  if (!current) return
  const chats = { ...(current.chats ?? {}), [conversationId]: chatId }
  await saveSettings({ ...current, chats })
}

export async function clearChatIdForConversation(conversationId: string): Promise<void> {
  const current = loadSettings()
  if (!current?.chats?.[conversationId]) return
  const chats = { ...current.chats }
  delete chats[conversationId]
  await saveSettings({ ...current, chats })
}

// ---------------------------------------------------------------------------
// Last-used RAG group id (so the save-to-RAG picker pre-selects it).
// ---------------------------------------------------------------------------

export function getLastRagGroupId(): string | undefined {
  return loadSettings()?.lastRagGroupId
}

export async function setLastRagGroupId(id: string): Promise<void> {
  const current = loadSettings()
  if (!current) return
  await saveSettings({ ...current, lastRagGroupId: id })
}

// ---------------------------------------------------------------------------
// Mail Routes — per-email AI automations (docs/MAIL_ROUTES.md).
//
// Stored inside the auth-gated settings record under `routes`. Returns a fresh
// empty state when the user hasn't created any route yet (or isn't signed in),
// so callers never have to null-check the shape.
// ---------------------------------------------------------------------------

export function getMailRoutes(): MailRoutesState {
  const current = loadSettings()
  return current?.routes ?? { ...EMPTY_ROUTES_STATE }
}

export async function saveMailRoutes(state: MailRoutesState): Promise<void> {
  const current = loadSettings()
  if (!current) {
    throw new Error('No roaming settings to store routes in — sign in first')
  }
  await saveSettings({ ...current, routes: state })
}

// ---------------------------------------------------------------------------
// Preferred Synaplan instance URL.
//
// Stored independently of the auth-gated settings (it has its own key, no
// `apiKey` required) so the SignIn screen remembers the instance the user
// last chose — even before they have signed in, and after sign-out.
// ---------------------------------------------------------------------------

const BASE_URL_KEY = 'synamail.preferredBaseUrl'

export function getPreferredBaseUrl(): string | undefined {
  const v = store().get(BASE_URL_KEY)
  return typeof v === 'string' && v.length > 0 ? v : undefined
}

export function setPreferredBaseUrl(url: string): Promise<void> {
  const s = store()
  s.set(BASE_URL_KEY, url)
  return new Promise((resolve, reject) => {
    s.saveAsync((res?: Office.AsyncResult<void>) => {
      if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(res?.error ?? new Error('saveAsync failed'))
    })
  })
}
