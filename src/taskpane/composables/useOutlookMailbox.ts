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

  return { searchMailbox, getMessageMime, ewsAvailable }
}
