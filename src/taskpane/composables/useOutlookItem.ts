/**
 * Reactive accessor for the currently-selected Outlook item.
 *
 * Wraps the async Office.js item APIs in a stable Vue-friendly shape.
 * In tests we inject a mocked `Office.context.mailbox.item` via setup.ts.
 */

import { onBeforeUnmount, onMounted, ref } from 'vue'
import { cleanEmailForIngest } from '@shared/email-clean'

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

/**
 * Plain-text export of a message: a small, clean header block + a blank line +
 * the cleaned body. The body is run through `cleanEmailForIngest` so the noise
 * (tracking URLs, control/zero-width garbage, routing codes, the quoted reply
 * tail) never reaches the knowledge base — only what's useful for RAG does.
 */
export function buildEmailText(snapshot: OutlookItemSnapshot): string {
  const header: string[] = []
  if (snapshot.subject) header.push(`Subject: ${snapshot.subject}`)
  if (snapshot.from) header.push(`From: ${snapshot.from}`)
  if (snapshot.to.length) header.push(`To: ${snapshot.to.join(', ')}`)
  if (snapshot.cc.length) header.push(`Cc: ${snapshot.cc.join(', ')}`)
  const body = cleanEmailForIngest(snapshot.bodyText ?? '')
  return header.length ? `${header.join('\n')}\n\n${body}` : body
}

function sanitizeFilename(s: string): string {
  const cleaned = s.replace(/[^a-zA-Z0-9-_]+/g, '_').replace(/^_+|_+$/g, '')
  return cleaned.slice(0, 80) || 'message'
}

// ---------------------------------------------------------------------------
// Attachments (inline screenshots + file attachments)
//
// An image-only email yields empty `bodyText`, so the AI has nothing to work
// with. We pull the image bytes here and upload them to Synaplan, whose Vision
// AI extracts the visible text/description on upload — letting the user "ask"
// about a screenshot exactly like a photo attachment in chat. For the
// knowledge base we also pull document attachments (pdf / word / ppt / …) so
// Synaplan can extract their text and vectorise it alongside the email body.
// ---------------------------------------------------------------------------

/** Skip anything larger than this (base64 balloons memory + upload time). */
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024

/** Inline images below this size are almost always signature logos / spacers. */
const INLINE_IMAGE_NOISE_BYTES = 8 * 1024

/** Document attachment types Synaplan can extract text from. */
const SUPPORTED_DOC_EXT_RE = /\.(pdf|docx?|pptx?|xlsx?|txt|csv|md|rtf|odt|odp|ods)$/i

const DOC_MIME_BY_EXT: Record<string, string> = {
  pdf: 'application/pdf',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  txt: 'text/plain',
  csv: 'text/csv',
  md: 'text/markdown',
  rtf: 'application/rtf',
  odt: 'application/vnd.oasis.opendocument.text',
  odp: 'application/vnd.oasis.opendocument.presentation',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
}

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

/** True when an attachment is a document whose text Synaplan can extract. */
function isSupportedDocument(att: Office.AttachmentDetails): boolean {
  if (isSupportedImage(att)) return false
  const name = (att.name ?? '').toLowerCase()
  return SUPPORTED_DOC_EXT_RE.test(name)
}

/** Attachments too large to upload, or that the host only exposes as a URL. */
function isSkippableAttachment(att: Office.AttachmentDetails): boolean {
  const size = typeof att.size === 'number' ? att.size : 0
  if (size > MAX_ATTACHMENT_BYTES) return true
  // Cloud / URL attachments carry no bytes we can read.
  return att.attachmentType === Office.MailboxEnums.AttachmentType.Cloud
}

/**
 * Tiny inline images are overwhelmingly signature logos, spacers, or tracking
 * pixels — vectorising them just adds noise, so we drop them from the KB set.
 * Real pasted screenshots are far larger than the threshold.
 */
function isNoiseImage(att: Office.AttachmentDetails): boolean {
  const size = typeof att.size === 'number' ? att.size : 0
  return !!att.isInline && size > 0 && size < INLINE_IMAGE_NOISE_BYTES
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

function fileExtension(name: string): string {
  const dot = name.lastIndexOf('.')
  return dot > 0 ? name.slice(dot + 1).toLowerCase() : ''
}

/** MIME type for a document attachment: the host's, else guessed by extension. */
function mimeForDocument(att: Office.AttachmentDetails): string {
  const ct = (att.contentType ?? '').toLowerCase()
  if (ct && ct !== 'application/octet-stream') return ct
  return DOC_MIME_BY_EXT[fileExtension(att.name ?? '')] ?? 'application/octet-stream'
}

/** Preserve a document's original extension while sanitising the base name. */
function documentFilename(att: Office.AttachmentDetails): string {
  const name = att.name || 'attachment'
  const dot = name.lastIndexOf('.')
  const ext = dot > 0 ? '.' + fileExtension(name) : ''
  const base = sanitizeFilename(dot > 0 ? name.slice(0, dot) : name)
  return base + ext
}

/** How to turn one selected attachment into an uploadable file. */
interface AttachmentPlan {
  id: string
  filename: string
  mimeType: string
}

/** The current read item, iff the host can hand us attachment bytes (1.8+). */
function itemWithAttachmentContent():
  | (Office.MessageRead & {
      getAttachmentContentAsync?: (
        id: string,
        cb: (r: Office.AsyncResult<Office.AttachmentContent>) => void,
      ) => void
    })
  | null {
  const raw = (typeof Office !== 'undefined' ? Office.context?.mailbox?.item : undefined) as
    | (Office.MessageRead & {
        getAttachmentContentAsync?: (
          id: string,
          cb: (r: Office.AsyncResult<Office.AttachmentContent>) => void,
        ) => void
      })
    | undefined
  if (!raw || typeof raw.getAttachmentContentAsync !== 'function') return null
  return raw
}

/** Fetch a set of planned attachments as base64 payloads, skipping failures. */
async function fetchPlannedAttachments(plans: AttachmentPlan[]): Promise<MessageFile[]> {
  const raw = itemWithAttachmentContent()
  if (!raw || plans.length === 0) return []
  const out: MessageFile[] = []
  for (const plan of plans) {
    const content = await new Promise<Office.AttachmentContent | null>((resolve) => {
      try {
        raw.getAttachmentContentAsync!(plan.id, (r) => {
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
    out.push({ filename: plan.filename, contentBase64: content.content, mimeType: plan.mimeType })
  }
  return out
}

function imagePlan(att: Office.AttachmentDetails): AttachmentPlan {
  const mimeType = mimeForImage(att)
  return {
    id: att.id,
    filename: sanitizeFilename(att.name || 'image') + extForMime(mimeType),
    mimeType,
  }
}

function documentPlan(att: Office.AttachmentDetails): AttachmentPlan {
  return { id: att.id, filename: documentFilename(att), mimeType: mimeForDocument(att) }
}

/**
 * Fetch all image attachments of the current read item as base64 payloads
 * suitable for `synaplanClient.fileUpload(...)`. Skips attachments the host
 * only exposes as a URL (e.g. cloud attachments) since we can't read bytes.
 * Returns [] when the host lacks `getAttachmentContentAsync` (pre-1.8).
 */
export async function getImageAttachments(snapshot: OutlookItemSnapshot): Promise<MessageFile[]> {
  const plans = snapshot.attachments.filter(isSupportedImage).map(imagePlan)
  return fetchPlannedAttachments(plans)
}

/**
 * Fetch document attachments (pdf / word / ppt / excel / text) as base64
 * payloads for upload. Synaplan extracts their text server-side. Oversized and
 * cloud/URL-only attachments are skipped.
 */
export async function getDocumentAttachments(
  snapshot: OutlookItemSnapshot,
): Promise<MessageFile[]> {
  const plans = snapshot.attachments
    .filter((att) => isSupportedDocument(att) && !isSkippableAttachment(att))
    .map(documentPlan)
  return fetchPlannedAttachments(plans)
}

/**
 * All attachments worth adding to the knowledge base for this email: real
 * images (screenshots/photos, minus tiny inline logos) plus supported
 * documents. Fetched in one pass so a single email save uploads the body and
 * every meaningful artefact together.
 */
export async function getKnowledgeBaseAttachments(
  snapshot: OutlookItemSnapshot,
): Promise<MessageFile[]> {
  const plans: AttachmentPlan[] = []
  for (const att of snapshot.attachments) {
    if (isSkippableAttachment(att)) continue
    if (isSupportedImage(att)) {
      if (isNoiseImage(att)) continue
      plans.push(imagePlan(att))
    } else if (isSupportedDocument(att)) {
      plans.push(documentPlan(att))
    }
  }
  return fetchPlannedAttachments(plans)
}

/** Count attachments (images + documents) that a KB save would upload. */
export function countKnowledgeBaseAttachments(snapshot: OutlookItemSnapshot): number {
  return snapshot.attachments.filter(
    (att) =>
      !isSkippableAttachment(att) &&
      ((isSupportedImage(att) && !isNoiseImage(att)) || isSupportedDocument(att)),
  ).length
}

function toBase64Utf8(s: string): string {
  // `btoa` exists in browsers, Office WebViews, and jsdom (test runtime).
  return globalThis.btoa(unescape(encodeURIComponent(s)))
}
