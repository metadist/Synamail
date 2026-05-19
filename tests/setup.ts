/**
 * Vitest setup. Stubs `Office.context` so composables and views can be
 * instantiated outside the Office runtime.
 */

import { afterEach, beforeEach, vi } from 'vitest'

interface RoamingSettingsStore {
  [key: string]: unknown
}

function makeOfficeStub() {
  const roaming: RoamingSettingsStore = {}
  return {
    onReady: (cb: () => void) => cb(),
    EventType: {
      ItemChanged: 'olkItemSelectedChanged',
      DialogMessageReceived: 'dialogMessageReceived',
      DialogEventReceived: 'dialogEventReceived',
    },
    MailboxEnums: {
      ItemType: { Message: 'message', Appointment: 'appointment' },
    },
    CoercionType: { Text: 'text', Html: 'html' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      displayLanguage: 'en-US',
      mailbox: {
        item: undefined,
        addHandlerAsync: vi.fn(),
        removeHandlerAsync: vi.fn(),
      },
      roamingSettings: {
        get: (key: string) => roaming[key],
        set: (key: string, value: unknown) => {
          roaming[key] = value
        },
        remove: (key: string) => {
          delete roaming[key]
        },
        saveAsync: (cb: (r: { status: string }) => void) => cb({ status: 'succeeded' }),
      },
      ui: {
        displayDialogAsync: vi.fn(),
        messageParent: vi.fn(),
      },
    },
  }
}

beforeEach(() => {
  ;(globalThis as unknown as Record<string, unknown>).Office = makeOfficeStub()
})

afterEach(() => {
  delete (globalThis as unknown as Record<string, unknown>).Office
})
