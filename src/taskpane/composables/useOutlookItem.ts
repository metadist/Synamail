/**
 * Reactive accessor for the currently-selected Outlook item.
 *
 * Wraps the async Office.js item APIs in a stable Vue-friendly shape.
 * In tests we inject a mocked `Office.context.mailbox.item` via setup.ts.
 */

import { onBeforeUnmount, onMounted, ref } from 'vue'

export type ItemMode = 'read' | 'compose' | 'none'

export interface OutlookItemSnapshot {
  mode: ItemMode
  subject: string
  from?: string
  to: string[]
  cc: string[]
  conversationId?: string
  bodyText: string
  attachments: Office.AttachmentDetails[]
}

const EMPTY: OutlookItemSnapshot = {
  mode: 'none',
  subject: '',
  to: [],
  cc: [],
  bodyText: '',
  attachments: [],
}

export function useOutlookItem() {
  const item = ref<OutlookItemSnapshot>(EMPTY)
  const loading = ref(false)
  const error = ref<string | null>(null)
  let itemChangedToken: string | null = null

  async function refresh(): Promise<void> {
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      item.value = EMPTY
      return
    }
    loading.value = true
    error.value = null
    try {
      item.value = await snapshot(Office.context.mailbox.item)
    } catch (err) {
      error.value = err instanceof Error ? err.message : String(err)
    } finally {
      loading.value = false
    }
  }

  onMounted(() => {
    void refresh()
    if (typeof Office !== 'undefined' && Office.context?.mailbox) {
      try {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          () => void refresh(),
          (res: Office.AsyncResult<void>) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              itemChangedToken = 'attached'
            }
          },
        )
      } catch {
        // Some clients don't support ItemChanged — refresh on demand only.
      }
    }
  })

  onBeforeUnmount(() => {
    if (itemChangedToken && typeof Office !== 'undefined' && Office.context?.mailbox) {
      try {
        Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged)
      } catch {
        /* no-op */
      }
    }
  })

  return { item, loading, error, refresh }
}

async function snapshot(raw: Office.Item): Promise<OutlookItemSnapshot> {
  const itemType = (raw as Office.MessageRead | Office.MessageCompose).itemType
  if (itemType !== Office.MailboxEnums.ItemType.Message) {
    return EMPTY
  }

  const mode = detectMode(raw)
  const subject = await readField(raw, 'subject')
  const from = await readField(raw, 'from')
  const to = await readRecipients(raw, 'to')
  const cc = await readRecipients(raw, 'cc')
  const conversationId = (raw as Office.MessageRead).conversationId
  const bodyText = await readBody(raw)
  const attachments = await readAttachments(raw)

  return { mode, subject, from, to, cc, conversationId, bodyText, attachments }
}

function detectMode(raw: Office.Item): ItemMode {
  // Compose-mode items expose async setters on body; read-mode items don't.
  const body = (raw as Office.MessageRead | Office.MessageCompose).body as
    | { setAsync?: unknown; getAsync?: unknown }
    | undefined
  if (body && typeof (body as { setAsync?: unknown }).setAsync === 'function') {
    return 'compose'
  }
  if (body && typeof (body as { getAsync?: unknown }).getAsync === 'function') {
    return 'read'
  }
  return 'none'
}

async function readField(raw: Office.Item, field: 'subject' | 'from'): Promise<string> {
  // Read-mode `subject` is a plain string; compose-mode is async.
  const v = (raw as unknown as Record<string, unknown>)[field]
  if (typeof v === 'string') return v
  if (field === 'from' && v && typeof v === 'object' && 'emailAddress' in v) {
    return (v as { emailAddress: string }).emailAddress
  }
  if (v && typeof v === 'object' && 'getAsync' in v) {
    return new Promise((resolve) => {
      ;(v as { getAsync: (cb: (r: Office.AsyncResult<unknown>) => void) => void }).getAsync((r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return resolve('')
        const val = r.value
        if (typeof val === 'string') resolve(val)
        else if (val && typeof val === 'object' && 'emailAddress' in val) {
          resolve((val as { emailAddress: string }).emailAddress)
        } else resolve('')
      })
    })
  }
  return ''
}

async function readRecipients(raw: Office.Item, field: 'to' | 'cc'): Promise<string[]> {
  const v = (raw as unknown as Record<string, unknown>)[field]
  if (Array.isArray(v)) {
    return v
      .map((r) =>
        r && typeof r === 'object' && 'emailAddress' in r
          ? (r as { emailAddress: string }).emailAddress
          : '',
      )
      .filter(Boolean)
  }
  if (v && typeof v === 'object' && 'getAsync' in v) {
    return new Promise((resolve) => {
      ;(
        v as {
          getAsync: (cb: (r: Office.AsyncResult<unknown[]>) => void) => void
        }
      ).getAsync((r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return resolve([])
        resolve(
          (r.value ?? [])
            .map((x: unknown) =>
              x && typeof x === 'object' && 'emailAddress' in x
                ? (x as { emailAddress: string }).emailAddress
                : '',
            )
            .filter(Boolean),
        )
      })
    })
  }
  return []
}

async function readBody(raw: Office.Item): Promise<string> {
  const body = (raw as Office.MessageRead | Office.MessageCompose).body
  if (!body) return ''
  return new Promise((resolve) => {
    try {
      body.getAsync(Office.CoercionType.Text, (r: Office.AsyncResult<string>) => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value ?? '') : '')
      })
    } catch {
      resolve('')
    }
  })
}

async function readAttachments(raw: Office.Item): Promise<Office.AttachmentDetails[]> {
  const direct = (raw as Office.MessageRead).attachments
  if (Array.isArray(direct)) return direct
  return []
}
