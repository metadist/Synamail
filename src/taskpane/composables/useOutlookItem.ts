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

// ---------------------------------------------------------------------------
// Compose-mode body writing ("Draft from prompt", docs/FEATURES.md §2.1)
// ---------------------------------------------------------------------------

/**
 * Replace the open draft's body with `html`. Resolves `true` on success and
 * `false` when there is no editable compose item (e.g. read mode) or the host
 * rejects the write. Uses `body.setAsync` with HTML coercion per §2.1 — this
 * replaces the current draft body (the intended behaviour for generating a
 * full email from an intent).
 */
export function setComposeBody(html: string): Promise<boolean> {
  const item = (typeof Office !== 'undefined' ? Office.context?.mailbox?.item : undefined) as
    | Office.MessageCompose
    | undefined
  const body = item?.body as
    | {
        setAsync?: (
          data: string,
          options: { coercionType: Office.CoercionType },
          cb: (r: Office.AsyncResult<void>) => void,
        ) => void
      }
    | undefined
  if (!body || typeof body.setAsync !== 'function') return Promise.resolve(false)
  return new Promise((resolve) => {
    try {
      body.setAsync!(html, { coercionType: Office.CoercionType.Html }, (r) => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded)
      })
    } catch {
      resolve(false)
    }
  })
}

/**
 * Reply to the currently-open (read-mode) message, pre-filling the reply body
 * with `html`. Opens Outlook's reply compose window. Resolves `true` on
 * success, `false` when the host has no `displayReplyForm` (e.g. compose mode
 * or a non-message item).
 */
export function displayReplyWithBody(html: string): boolean {
  const item = (typeof Office !== 'undefined' ? Office.context?.mailbox?.item : undefined) as
    | { displayReplyForm?: (arg: { htmlBody: string }) => void }
    | undefined
  if (!item || typeof item.displayReplyForm !== 'function') return false
  try {
    item.displayReplyForm({ htmlBody: html })
    return true
  } catch {
    return false
  }
}

/**
 * Open a brand-new message compose window pre-filled with `html` (and an
 * optional subject). Resolves `true` on success, `false` when the host lacks
 * `displayNewMessageForm` (pre-Mailbox 1.6 or an unsupported client).
 */
export function displayNewMessageWithBody(html: string, subject?: string): boolean {
  const mb = (typeof Office !== 'undefined' ? Office.context?.mailbox : undefined) as
    | {
        displayNewMessageForm?: (arg: { htmlBody: string; subject?: string }) => void
      }
    | undefined
  if (!mb || typeof mb.displayNewMessageForm !== 'function') return false
  try {
    mb.displayNewMessageForm(subject ? { htmlBody: html, subject } : { htmlBody: html })
    return true
  } catch {
    return false
  }
}

// ---------------------------------------------------------------------------
// File extraction for "Save to knowledge base"
// ---------------------------------------------------------------------------

export interface MessageFile {
  filename: string
  contentBase64: string
  mimeType: string
}

/**
 * Capture the currently-selected read-mode message as an uploadable file.
 *
 * We export the message as PLAIN TEXT (`.txt`) — a small header block
 * (subject / participants) followed by the body — NOT `.eml`. Synaplan's file
 * ingestion only accepts document/media types (pdf, docx, txt, csv, …) and
 * rejects `message/rfc822`; the body text is also what's actually useful for
 * RAG (an `.eml`'s MIME envelope is noise). Building from the snapshot keeps
 * this working across every Outlook host without depending on
 * `getAsFileAsync`.
 */
export function getReadItemAsFile(snapshot: OutlookItemSnapshot): MessageFile {
  const baseName = sanitizeFilename(snapshot.subject || 'message')
  return {
    filename: `${baseName}.txt`,
    contentBase64: toBase64Utf8(buildEmailText(snapshot)),
    mimeType: 'text/plain',
  }
}

/** Plain-text export of a message: header lines + a blank line + the body. */
export function buildEmailText(snapshot: OutlookItemSnapshot): string {
  const header: string[] = []
  if (snapshot.subject) header.push(`Subject: ${snapshot.subject}`)
  if (snapshot.from) header.push(`From: ${snapshot.from}`)
  if (snapshot.to.length) header.push(`To: ${snapshot.to.join(', ')}`)
  if (snapshot.cc.length) header.push(`Cc: ${snapshot.cc.join(', ')}`)
  const body = snapshot.bodyText ?? ''
  return header.length ? `${header.join('\n')}\n\n${body}` : body
}

function sanitizeFilename(s: string): string {
  const cleaned = s.replace(/[^a-zA-Z0-9-_]+/g, '_').replace(/^_+|_+$/g, '')
  return cleaned.slice(0, 80) || 'message'
}

// ---------------------------------------------------------------------------
// Image attachments (inline screenshots + file attachments)
//
// An image-only email yields empty `bodyText`, so the AI has nothing to work
// with. We pull the image bytes here and upload them to Synaplan, whose Vision
// AI extracts the visible text/description on upload — letting the user "ask"
// about a screenshot exactly like a photo attachment in chat.
// ---------------------------------------------------------------------------

/** True when an attachment is an image we can send to Synaplan's Vision AI. */
function isSupportedImage(att: Office.AttachmentDetails): boolean {
  const ct = (att.contentType ?? '').toLowerCase()
  if (ct.startsWith('image/')) {
    return /image\/(jpeg|jpg|png|gif|webp)/.test(ct)
  }
  // Some hosts leave contentType blank for inline images — fall back to name.
  const name = (att.name ?? '').toLowerCase()
  return /\.(jpe?g|png|gif|webp)$/.test(name)
}

/** Count image attachments without fetching their content (for UI hints). */
export function countImageAttachments(snapshot: OutlookItemSnapshot): number {
  return snapshot.attachments.filter(isSupportedImage).length
}

function mimeForImage(att: Office.AttachmentDetails): string {
  const ct = (att.contentType ?? '').toLowerCase()
  if (ct.startsWith('image/')) return ct
  const name = (att.name ?? '').toLowerCase()
  if (name.endsWith('.png')) return 'image/png'
  if (name.endsWith('.gif')) return 'image/gif'
  if (name.endsWith('.webp')) return 'image/webp'
  return 'image/jpeg'
}

/**
 * Fetch all image attachments of the current read item as base64 payloads
 * suitable for `synaplanClient.fileUpload(...)`. Skips attachments the host
 * only exposes as a URL (e.g. cloud attachments) since we can't read bytes.
 * Returns [] when the host lacks `getAttachmentContentAsync` (pre-1.8).
 */
export async function getImageAttachments(snapshot: OutlookItemSnapshot): Promise<MessageFile[]> {
  const raw = (typeof Office !== 'undefined' ? Office.context?.mailbox?.item : undefined) as
    | (Office.MessageRead & {
        getAttachmentContentAsync?: (
          id: string,
          cb: (r: Office.AsyncResult<Office.AttachmentContent>) => void,
        ) => void
      })
    | undefined
  if (!raw || typeof raw.getAttachmentContentAsync !== 'function') return []

  const images = snapshot.attachments.filter(isSupportedImage)
  const out: MessageFile[] = []
  for (const att of images) {
    const content = await new Promise<Office.AttachmentContent | null>((resolve) => {
      try {
        raw.getAttachmentContentAsync!(att.id, (r) => {
          resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)
        })
      } catch {
        resolve(null)
      }
    })
    // Only Base64 content is directly uploadable; Url/Eml/ICalendar are skipped.
    if (!content || content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
      continue
    }
    out.push({
      filename: sanitizeFilename(att.name || 'image') + extForMime(mimeForImage(att)),
      contentBase64: content.content,
      mimeType: mimeForImage(att),
    })
  }
  return out
}

function extForMime(mime: string): string {
  switch (mime) {
    case 'image/png':
      return '.png'
    case 'image/gif':
      return '.gif'
    case 'image/webp':
      return '.webp'
    default:
      return '.jpg'
  }
}

function toBase64Utf8(s: string): string {
  // `btoa` exists in browsers, Office WebViews, and jsdom (test runtime).
  return globalThis.btoa(unescape(encodeURIComponent(s)))
}
