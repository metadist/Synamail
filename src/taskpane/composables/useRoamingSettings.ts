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

export function patchSettings(patch: Partial<RoamingSettings>): Promise<void> {
  const current = loadSettings()
  if (!current) {
    throw new Error('No roaming settings to patch — sign in first')
  }
  return saveSettings({ ...current, ...patch })
}
