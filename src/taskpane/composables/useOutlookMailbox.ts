/**
 * Live Outlook mailbox access via EWS (`makeEwsRequestAsync`).
 *
 * Used by the "Search & vectorize" command to find messages about a topic
 * across the user's mailbox (complementing the Synaplan knowledge-base search).
 *
 * EWS notes:
 *   - `makeEwsRequestAsync` requires `ReadWriteMailbox` in manifest.xml and is
 *     available on classic Outlook desktop + Outlook on the web (NOT new
 *     Outlook for Windows). When unavailable we fall back to canned mock data
 *     so the dev loop / Vitest still exercise the UI end-to-end.
 *   - FindItem with a `QueryString` runs Outlook's AQS full-text search.
 *   - GetItem with `IncludeMimeContent` returns the base64 `.eml` we upload.
 */

import type { MessageFile } from './useOutlookItem'
import type { CreateSpamRuleResult, SenderHistoryItem, SenderHistoryResult } from '@shared/types'

export interface MailboxHit {
  /** EWS ItemId, usable for GetItem. */
  id: string
  subject: string
  from: string
  /** ISO 8601 received date (best-effort). */
  date: string
  source: 'mailbox'
}

export interface MailboxSearchResult {
  hits: MailboxHit[]
  /** True when results came from EWS; false when the mock stub answered. */
  fromOutlook: boolean
}

interface EwsHost {
  makeEwsRequestAsync?: (
    data: string,
    callback: (result: Office.AsyncResult<string>) => void,
  ) => void
}

function ewsHost(): EwsHost | undefined {
  return (typeof Office !== 'undefined' ? Office.context?.mailbox : undefined) as
    | EwsHost
    | undefined
}

export function ewsAvailable(): boolean {
  return typeof ewsHost()?.makeEwsRequestAsync === 'function'
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

export function buildFindItemRequest(query: string, limit: number): string {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    '<soap:Body>' +
    '<m:FindItem Traversal="Shallow">' +
    '<m:ItemShape><t:BaseShape>IdOnly</t:BaseShape>' +
    '<t:AdditionalProperties>' +
    '<t:FieldURI FieldURI="item:Subject"/>' +
    '<t:FieldURI FieldURI="item:DateTimeReceived"/>' +
    '<t:FieldURI FieldURI="message:From"/>' +
    '</t:AdditionalProperties></m:ItemShape>' +
    `<m:IndexedPageItemView MaxEntriesReturned="${limit}" Offset="0" BasePoint="Beginning"/>` +
    '<m:ParentFolderIds><t:DistinguishedFolderId Id="inbox"/></m:ParentFolderIds>' +
    `<m:QueryString>${escapeXml(query)}</m:QueryString>` +
    '</m:FindItem></soap:Body></soap:Envelope>'
  )
}

export function buildGetItemMimeRequest(id: string): string {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    '<soap:Body><m:GetItem>' +
    '<m:ItemShape><t:BaseShape>IdOnly</t:BaseShape>' +
    '<t:IncludeMimeContent>true</t:IncludeMimeContent></m:ItemShape>' +
    `<m:ItemIds><t:ItemId Id="${escapeXml(id)}"/></m:ItemIds>` +
    '</m:GetItem></soap:Body></soap:Envelope>'
  )
}

// AQS participant search — "from:alice@example.com" matches the sender field.
export function senderQuery(email: string): string {
  return `from:${email}`
}

// FindItem variant that also pulls IsRead so the history can mark unread items.
export function buildSenderFindRequest(email: string, limit: number): string {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    '<soap:Body>' +
    '<m:FindItem Traversal="Shallow">' +
    '<m:ItemShape><t:BaseShape>IdOnly</t:BaseShape>' +
    '<t:AdditionalProperties>' +
    '<t:FieldURI FieldURI="item:Subject"/>' +
    '<t:FieldURI FieldURI="item:DateTimeReceived"/>' +
    '<t:FieldURI FieldURI="message:From"/>' +
    '<t:FieldURI FieldURI="message:IsRead"/>' +
    '</t:AdditionalProperties></m:ItemShape>' +
    `<m:IndexedPageItemView MaxEntriesReturned="${limit}" Offset="0" BasePoint="Beginning"/>` +
    '<m:SortOrder><t:FieldOrder Order="Descending">' +
    '<t:FieldURI FieldURI="item:DateTimeReceived"/></t:FieldOrder></m:SortOrder>' +
    '<m:ParentFolderIds><t:DistinguishedFolderId Id="inbox"/></m:ParentFolderIds>' +
    `<m:QueryString>${escapeXml(senderQuery(email))}</m:QueryString>` +
    '</m:FindItem></soap:Body></soap:Envelope>'
  )
}

// UpdateInboxRules — create a server-side rule that moves future mail from the
// sender into Junk Email. Identified for the user by the display name.
export function buildBlockRuleRequest(email: string, displayName: string): string {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    '<soap:Body><m:UpdateInboxRules>' +
    '<m:RemoveOutlookRuleBlob>false</m:RemoveOutlookRuleBlob>' +
    '<m:Operations><t:CreateRuleOperation><t:Rule>' +
    `<t:DisplayName>${escapeXml(displayName)}</t:DisplayName>` +
    '<t:Priority>1</t:Priority>' +
    '<t:IsEnabled>true</t:IsEnabled>' +
    '<t:Conditions><t:ContainsSenderStrings>' +
    `<t:String>${escapeXml(email)}</t:String>` +
    '</t:ContainsSenderStrings></t:Conditions>' +
    '<t:Actions><t:MoveToFolder>' +
    '<t:DistinguishedFolderId Id="junkemail"/>' +
    '</t:MoveToFolder></t:Actions>' +
    '</t:Rule></t:CreateRuleOperation></m:Operations>' +
    '</m:UpdateInboxRules></soap:Body></soap:Envelope>'
  )
}

export function buildMoveToJunkRequest(ids: string[]): string {
  const itemIds = ids.map((id) => `<t:ItemId Id="${escapeXml(id)}"/>`).join('')
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    '<soap:Body><m:MoveItem>' +
    '<m:ToFolderId><t:DistinguishedFolderId Id="junkemail"/></m:ToFolderId>' +
    `<m:ItemIds>${itemIds}</m:ItemIds>` +
    '</m:MoveItem></soap:Body></soap:Envelope>'
  )
}

function localName(tag: string): string {
  const i = tag.indexOf(':')
  return i === -1 ? tag : tag.slice(i + 1)
}

function firstByLocalName(root: Element | Document, name: string): Element | null {
  const all = root.getElementsByTagName('*')
  for (let i = 0; i < all.length; i++) {
    if (localName(all[i].tagName) === name) return all[i]
  }
  return null
}

function allByLocalName(root: Element | Document, name: string): Element[] {
  const out: Element[] = []
  const all = root.getElementsByTagName('*')
  for (let i = 0; i < all.length; i++) {
    if (localName(all[i].tagName) === name) out.push(all[i])
  }
  return out
}

export function parseFindItemResponse(xml: string): MailboxHit[] {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  const messages = allByLocalName(doc, 'Message')
  return messages.map((msg) => {
    const idEl = firstByLocalName(msg, 'ItemId')
    const subjectEl = firstByLocalName(msg, 'Subject')
    const dateEl = firstByLocalName(msg, 'DateTimeReceived')
    const fromEmailEl = (() => {
      const from = firstByLocalName(msg, 'From')
      return from ? firstByLocalName(from, 'EmailAddress') : null
    })()
    return {
      id: idEl?.getAttribute('Id') ?? '',
      subject: subjectEl?.textContent ?? '',
      from: fromEmailEl?.textContent ?? '',
      date: dateEl?.textContent ?? '',
      source: 'mailbox' as const,
    }
  })
}

export function parseGetItemMime(xml: string): string {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  return firstByLocalName(doc, 'MimeContent')?.textContent ?? ''
}

export function parseSenderHistory(xml: string): SenderHistoryItem[] {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  return allByLocalName(doc, 'Message').map((msg) => {
    const idEl = firstByLocalName(msg, 'ItemId')
    const subject = firstByLocalName(msg, 'Subject')?.textContent ?? ''
    const date = firstByLocalName(msg, 'DateTimeReceived')?.textContent ?? ''
    const isRead = firstByLocalName(msg, 'IsRead')?.textContent
    return {
      date,
      subject,
      snippet: '',
      unread: isRead === 'false',
      messageId: idEl?.getAttribute('Id') ?? undefined,
    }
  })
}

// Count how many MoveItem operations the server reported as successful.
export function parseMovedCount(xml: string): number {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  return allByLocalName(doc, 'MoveItemResponseMessage').filter(
    (el) => el.getAttribute('ResponseClass') === 'Success',
  ).length
}

// True when the SOAP body reports an overall success response class.
export function isEwsSuccess(xml: string): boolean {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  const all = doc.getElementsByTagName('*')
  for (let i = 0; i < all.length; i++) {
    const rc = all[i].getAttribute('ResponseClass')
    if (rc === 'Success') return true
    if (rc === 'Error') return false
  }
  // Fall back to the textual ResponseCode element.
  return firstByLocalName(doc, 'ResponseCode')?.textContent === 'NoError'
}

function callEws(request: string): Promise<string> {
  const host = ewsHost()
  if (!host?.makeEwsRequestAsync) {
    return Promise.reject(new Error('EWS is not available in this Outlook host.'))
  }
  return new Promise<string>((resolve, reject) => {
    host.makeEwsRequestAsync!(request, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value ?? '')
      else reject(new Error(r.error?.message ?? 'EWS request failed.'))
    })
  })
}

function sanitizeFilename(s: string): string {
  const cleaned = s.replace(/[^a-zA-Z0-9-_]+/g, '_').replace(/^_+|_+$/g, '')
  return cleaned.slice(0, 80) || 'message'
}

function mockSenderHistory(email: string, limit: number): SenderHistoryItem[] {
  const subjects = [
    'Re: Q3 review prep',
    'Lunch on Thursday?',
    'Re: vendor onboarding docs',
    'FYI: pricing changes effective May',
    'Quick question about the brief',
    'Re: photoshoot timeline',
  ]
  const base = Date.now()
  return Array.from({ length: Math.min(limit, subjects.length) }, (_, i) => ({
    date: new Date(base - (i + 1) * 3 * 86_400_000).toISOString(),
    subject: subjects[i % subjects.length],
    snippet: `(mock) Latest from ${email} — message #${i + 1} preview…`,
    unread: i < 2,
    messageId: `mock-msg-${i + 1}`,
  }))
}

function mockHits(query: string, limit: number): MailboxHit[] {
  const subjects = [
    `Re: ${query} planning`,
    `${query} — vendor quote`,
    `Notes on ${query}`,
    `FW: ${query} follow-up`,
    `${query} kickoff recap`,
  ]
  const base = Date.now()
  return Array.from({ length: Math.min(limit, subjects.length) }, (_, i) => ({
    id: `mock-ews-${i + 1}`,
    subject: subjects[i],
    from: `person${i + 1}@example.test`,
    date: new Date(base - (i + 1) * 86_400_000).toISOString(),
    source: 'mailbox' as const,
  }))
}

export function useOutlookMailbox() {
  async function searchMailbox(query: string, limit = 10): Promise<MailboxSearchResult> {
    const q = query.trim()
    if (!q) return { hits: [], fromOutlook: ewsAvailable() }
    if (!ewsAvailable()) {
      return { hits: mockHits(q, limit), fromOutlook: false }
    }
    const xml = await callEws(buildFindItemRequest(q, limit))
    return { hits: parseFindItemResponse(xml), fromOutlook: true }
  }

  async function getMessageMime(hit: MailboxHit): Promise<MessageFile> {
    const filename = `${sanitizeFilename(hit.subject || 'message')}.eml`
    if (!ewsAvailable() || hit.id.startsWith('mock-ews-')) {
      const fake = `Subject: ${hit.subject}\r\nFrom: ${hit.from}\r\n\r\n(mock body for "${hit.subject}")`
      return {
        filename,
        contentBase64: globalThis.btoa(unescape(encodeURIComponent(fake))),
        mimeType: 'message/rfc822',
      }
    }
    const xml = await callEws(buildGetItemMimeRequest(hit.id))
    return { filename, contentBase64: parseGetItemMime(xml), mimeType: 'message/rfc822' }
  }

  /** Recent messages from one sender, newest first (live mailbox via EWS). */
  async function senderHistory(email: string, limit = 12): Promise<SenderHistoryResult> {
    const cap = Math.max(1, Math.min(limit, 50))
    if (!ewsAvailable()) {
      const items = mockSenderHistory(email, cap)
      return { email, total: items.length, items, fromOutlook: false }
    }
    const xml = await callEws(buildSenderFindRequest(email, cap))
    const items = parseSenderHistory(xml)
    return { email, total: items.length, items, fromOutlook: true }
  }

  /**
   * Block a sender: create a server-side rule moving their future mail to Junk,
   * and (optionally) move the messages they've already sent. When EWS isn't
   * available the dev/mock loop still gets a deterministic result.
   */
  async function blockSender(email: string, alsoClean = false): Promise<CreateSpamRuleResult> {
    const displayName = `Synamail: block ${email}`
    if (!ewsAvailable()) {
      return {
        ruleId: `mock-rule-${Math.random().toString(36).slice(2, 8)}`,
        movedCount: alsoClean ? Math.floor(Math.random() * 6) : 0,
        serverSide: false,
      }
    }
    const ruleXml = await callEws(buildBlockRuleRequest(email, displayName))
    if (!isEwsSuccess(ruleXml)) {
      throw new Error('Outlook rejected the block-sender rule.')
    }
    let movedCount = 0
    if (alsoClean) {
      const found = parseFindItemResponse(
        await callEws(buildFindItemRequest(senderQuery(email), 50)),
      )
      const ids = found.map((h) => h.id).filter(Boolean)
      if (ids.length > 0) {
        movedCount = parseMovedCount(await callEws(buildMoveToJunkRequest(ids)))
      }
    }
    return { ruleId: displayName, movedCount, serverSide: true }
  }

  return { searchMailbox, getMessageMime, senderHistory, blockSender, ewsAvailable }
}
