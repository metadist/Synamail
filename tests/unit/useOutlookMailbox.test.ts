import { describe, expect, it } from 'vitest'
import {
  buildFindItemRequest,
  buildGetItemMimeRequest,
  ewsAvailable,
  parseFindItemResponse,
  parseGetItemMime,
  useOutlookMailbox,
} from '@/taskpane/composables/useOutlookMailbox'

const FIND_ITEM_XML = `<?xml version="1.0"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <s:Body>
    <m:FindItemResponse>
      <m:ResponseMessages>
        <m:FindItemResponseMessage ResponseClass="Success">
          <m:RootFolder>
            <t:Items>
              <t:Message>
                <t:ItemId Id="AAA111" ChangeKey="CK1"/>
                <t:Subject>Contract 2026</t:Subject>
                <t:DateTimeReceived>2026-05-01T10:00:00Z</t:DateTimeReceived>
                <t:From><t:Mailbox><t:Name>Alice</t:Name><t:EmailAddress>alice@example.test</t:EmailAddress></t:Mailbox></t:From>
              </t:Message>
              <t:Message>
                <t:ItemId Id="BBB222" ChangeKey="CK2"/>
                <t:Subject>Re: Contract 2026</t:Subject>
                <t:DateTimeReceived>2026-05-02T11:00:00Z</t:DateTimeReceived>
                <t:From><t:Mailbox><t:EmailAddress>bob@example.test</t:EmailAddress></t:Mailbox></t:From>
              </t:Message>
            </t:Items>
          </m:RootFolder>
        </m:FindItemResponseMessage>
      </m:ResponseMessages>
    </m:FindItemResponse>
  </s:Body>
</s:Envelope>`

const GET_ITEM_XML = `<?xml version="1.0"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <s:Body>
    <m:GetItemResponse>
      <m:ResponseMessages>
        <m:GetItemResponseMessage ResponseClass="Success">
          <m:Items>
            <t:Message>
              <t:MimeContent CharacterSet="UTF-8">QkFTRTY0RU1M</t:MimeContent>
            </t:Message>
          </m:Items>
        </m:GetItemResponseMessage>
      </m:ResponseMessages>
    </m:GetItemResponse>
  </s:Body>
</s:Envelope>`

describe('EWS request builders', () => {
  it('escapes the query and sets the page size in FindItem', () => {
    const req = buildFindItemRequest('a & b <c>', 5)
    expect(req).toContain('MaxEntriesReturned="5"')
    expect(req).toContain('a &amp; b &lt;c&gt;')
    expect(req).toContain('DistinguishedFolderId Id="inbox"')
  })

  it('includes the item id and requests MIME content in GetItem', () => {
    const req = buildGetItemMimeRequest('AAA111')
    expect(req).toContain('ItemId Id="AAA111"')
    expect(req).toContain('<t:IncludeMimeContent>true</t:IncludeMimeContent>')
  })
})

describe('EWS response parsers', () => {
  it('parses FindItem messages into hits', () => {
    const hits = parseFindItemResponse(FIND_ITEM_XML)
    expect(hits).toHaveLength(2)
    expect(hits[0]).toEqual({
      id: 'AAA111',
      subject: 'Contract 2026',
      from: 'alice@example.test',
      date: '2026-05-01T10:00:00Z',
      source: 'mailbox',
    })
    expect(hits[1].from).toBe('bob@example.test')
  })

  it('extracts the MIME content from GetItem', () => {
    expect(parseGetItemMime(GET_ITEM_XML)).toBe('QkFTRTY0RU1M')
  })
})

describe('useOutlookMailbox mock fallback', () => {
  it('reports EWS unavailable in the test stub', () => {
    expect(ewsAvailable()).toBe(false)
  })

  it('returns mock hits when EWS is unavailable', async () => {
    const { searchMailbox } = useOutlookMailbox()
    const r = await searchMailbox('supplier contract', 3)
    expect(r.fromOutlook).toBe(false)
    expect(r.hits.length).toBe(3)
    expect(r.hits[0].subject).toContain('supplier contract')
  })

  it('returns an empty result for a blank query', async () => {
    const { searchMailbox } = useOutlookMailbox()
    const r = await searchMailbox('   ')
    expect(r.hits).toEqual([])
  })

  it('builds a base64 eml for a mock hit', async () => {
    const { getMessageMime } = useOutlookMailbox()
    const file = await getMessageMime({
      id: 'mock-ews-1',
      subject: 'Topic!',
      from: 'x@y.test',
      date: '2026-05-01T00:00:00Z',
      source: 'mailbox',
    })
    expect(file.filename).toBe('Topic.eml')
    expect(file.mimeType).toBe('message/rfc822')
    expect(atob(file.contentBase64)).toContain('Subject: Topic!')
  })
})
