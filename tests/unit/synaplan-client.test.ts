import { describe, expect, it, vi } from 'vitest'
import {
  MockSynaplanClient,
  RealSynaplanClient,
  createSynaplanClient,
  isApiError,
} from '@shared/synaplan-client'

// ---------------------------------------------------------------------------
// Helpers — build typed fetch mocks that return canned JSON responses.
// ---------------------------------------------------------------------------

interface MockResponseStep {
  status?: number
  body?: unknown
}

function mockFetchSequence(steps: MockResponseStep[]): ReturnType<typeof vi.fn> {
  const fn = vi.fn()
  for (const step of steps) {
    const status = step.status ?? 200
    // 204/205 disallow a body per the Fetch spec; pass null to keep the
    // Response constructor happy.
    const body =
      status === 204 || status === 205
        ? null
        : typeof step.body === 'string'
          ? step.body
          : JSON.stringify(step.body ?? {})
    fn.mockResolvedValueOnce(
      new Response(body, {
        status,
        headers: { 'Content-Type': 'application/json' },
      }),
    )
  }
  return fn
}

function buildClient(
  fetchImpl: typeof fetch,
  opts: { maxRetries?: number } = {},
): RealSynaplanClient {
  return new RealSynaplanClient({
    baseUrl: 'https://api.test',
    apiKey: 'sk_test',
    fetchImpl,
    maxRetries: opts.maxRetries ?? 2,
  })
}

function lastCallBody(fetchImpl: ReturnType<typeof vi.fn>): unknown {
  const init = fetchImpl.mock.calls.at(-1)?.[1] as RequestInit | undefined
  if (!init?.body) return undefined
  return typeof init.body === 'string' ? JSON.parse(init.body) : init.body
}

// ---------------------------------------------------------------------------
// MockSynaplanClient — unchanged surface, smoke check the canned shapes
// ---------------------------------------------------------------------------

describe('MockSynaplanClient', () => {
  it('returns canned summarise output with bullet points', async () => {
    const c = new MockSynaplanClient(0)
    const r = await c.summarise({ subject: 'Hello', body: 'World', from: 'a@b.test', to: [] })
    expect(r.bullets.length).toBeGreaterThan(0)
    expect(r.summary).toContain('Hello')
  })

  it('classify routes invoice subjects to billing', async () => {
    const c = new MockSynaplanClient(0)
    const r = await c.classify({ subject: 'Q3 invoice', body: '...', from: 'x@y.test' })
    expect(r.category).toBe('billing')
  })

  it('reuses chatId across turns when caller supplies one', async () => {
    const c = new MockSynaplanClient(0)
    const a = await c.ask({ conversationId: 'abc', question: '?', chatId: 42 })
    const b = await c.ask({ conversationId: 'abc', question: 'again?', chatId: 42 })
    expect(a.chatId).toBe(42)
    expect(b.chatId).toBe(42)
  })
})

// ---------------------------------------------------------------------------
// Factory selection
// ---------------------------------------------------------------------------

describe('createSynaplanClient factory', () => {
  it('returns MockSynaplanClient when apiKey starts with mock-key-', () => {
    const c = createSynaplanClient({ baseUrl: 'https://x', apiKey: 'mock-key-abc' })
    expect(c).toBeInstanceOf(MockSynaplanClient)
  })

  it('returns MockSynaplanClient when useMock is true', () => {
    const c = createSynaplanClient({ baseUrl: 'https://x', apiKey: 'real', useMock: true })
    expect(c).toBeInstanceOf(MockSynaplanClient)
  })

  it('returns RealSynaplanClient otherwise', () => {
    const c = createSynaplanClient({ baseUrl: 'https://x', apiKey: 'sk_real' })
    expect(c).toBeInstanceOf(RealSynaplanClient)
  })
})

// ---------------------------------------------------------------------------
// RealSynaplanClient — transport behaviour
// ---------------------------------------------------------------------------

describe('RealSynaplanClient — transport', () => {
  it('injects the X-API-Key header on every request', async () => {
    const fetchImpl = mockFetchSequence([{ body: { success: true, user: { email: 'a@b.test' } } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await c.ping()
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    const headers = (fetchImpl.mock.calls[0][1] as RequestInit).headers as Headers
    expect(headers.get('X-API-Key')).toBe('sk_test')
    expect(headers.get('Accept')).toBe('application/json')
  })

  it('throws ApiError with status 401 on auth failure', async () => {
    const fetchImpl = mockFetchSequence([{ status: 401, body: 'nope' }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await expect(c.ping()).rejects.toMatchObject({ status: 401 })
  })

  it('retries 5xx with backoff up to maxRetries', async () => {
    const fetchImpl = mockFetchSequence([
      { status: 503, body: 'boom' },
      { status: 503, body: 'boom' },
      { body: { success: true, user: { email: 'a@b.test' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch, { maxRetries: 2 })
    const r = await c.ping()
    expect(r).toEqual({ ok: true, email: 'a@b.test' })
    expect(fetchImpl).toHaveBeenCalledTimes(3)
  })
})

// ---------------------------------------------------------------------------
// RealSynaplanClient — endpoint contracts
// ---------------------------------------------------------------------------

describe('RealSynaplanClient — ping', () => {
  it('reads email from auth/me', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, user: { email: 'admin@synaplan.test', isAdmin: true } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.ping()
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/auth/me')
    expect(r).toEqual({ ok: true, email: 'admin@synaplan.test' })
  })
})

describe('RealSynaplanClient — summarise / translate / draftReply / classify', () => {
  it('summarise posts a combined system+email message to /messages/send', async () => {
    const fetchImpl = mockFetchSequence([
      {
        body: {
          success: true,
          outgoingMessage: { text: '- Point one\n- Point two\n- Point three' },
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.summarise({
      subject: 'Q3 plan',
      body: 'Lorem ipsum',
      from: 'a@b.test',
    })
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/messages/send')
    const body = lastCallBody(fetchImpl) as { message: string }
    expect(body.message).toMatch(/summarise/i) // system prompt prefix
    expect(body.message).toContain('Subject: Q3 plan')
    expect(body.message).toContain('Lorem ipsum')
    expect(r.bullets).toEqual(['Point one', 'Point two', 'Point three'])
    expect(r.summary).toContain('Point one')
  })

  it('translate puts the AI text into result.translation', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'Hallo Welt' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.translate({ text: 'Hello world', targetLanguage: 'de' })
    expect(r.translation).toBe('Hallo Welt')
    const body = lastCallBody(fetchImpl) as { message: string }
    expect(body.message).toContain('language code "de"')
  })

  it('draftReply returns the AI text as htmlBody', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: '<p>Hi</p>' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.draftReply({
      subject: 'Re: order',
      body: 'previous',
      tone: 'concise',
      language: 'en',
    })
    expect(r.htmlBody).toBe('<p>Hi</p>')
  })

  it('classify parses JSON in the AI reply', async () => {
    const fetchImpl = mockFetchSequence([
      {
        body: {
          success: true,
          outgoingMessage: {
            text: '```json\n{"category":"billing","confidence":0.91,"reasoning":"contains invoice"}\n```',
          },
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.classify({ subject: 'Invoice', body: '...', from: 'a@b.test' })
    expect(r).toEqual({ category: 'billing', confidence: 0.91, reasoning: 'contains invoice' })
  })

  it('classify falls back to unknown when the AI reply has no parseable JSON', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'I think this is billing.' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.classify({ subject: 'q', body: 'b' })
    expect(r.category).toBe('unknown')
    expect(r.confidence).toBe(0)
    expect(r.reasoning).toMatch(/parse_failed/)
  })

  it('extracts AI text from alternative outgoingMessage field names', async () => {
    // The OpenAPI says `outgoingMessage` is `type: object` — the precise
    // field name isn't pinned yet. Cover the common alternatives.
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { content: 'via .content' } } },
      { body: { success: true, outgoingMessage: { message: 'via .message' } } },
      { body: { success: true, outgoingMessage: 'bare string' } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    expect((await c.translate({ text: 'x', targetLanguage: 'en' })).translation).toBe(
      'via .content',
    )
    expect((await c.translate({ text: 'x', targetLanguage: 'en' })).translation).toBe(
      'via .message',
    )
    expect((await c.translate({ text: 'x', targetLanguage: 'en' })).translation).toBe('bare string')
  })

  it('throws AI_FAILED when the response has success=false', async () => {
    const fetchImpl = mockFetchSequence([{ body: { success: false } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await expect(c.translate({ text: 'x', targetLanguage: 'en' })).rejects.toMatchObject({
      code: 'AI_FAILED',
    })
  })
})

describe('RealSynaplanClient — ask (chat round-trip)', () => {
  it('creates a chat on the first turn then sends with trackId', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, chat: { id: 17 } } },
      { body: { success: true, outgoingMessage: { text: 'answer' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.ask({ conversationId: 'thread-abc', question: 'why?' })
    expect(r).toEqual({ chatId: 17, answer: 'answer' })
    // Call 0: create chat
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/chats')
    // Call 1: send message with trackId
    expect(fetchImpl.mock.calls[1][0]).toBe('https://api.test/api/v1/messages/send')
    const sendBody = lastCallBody(fetchImpl) as { message: string; trackId: number }
    expect(sendBody.trackId).toBe(17)
    expect(sendBody.message).toContain('why?')
  })

  it('reuses an existing chatId without creating a new chat', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'follow-up answer' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.ask({ conversationId: 'thread-abc', question: 'more?', chatId: 99 })
    expect(r).toEqual({ chatId: 99, answer: 'follow-up answer' })
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/messages/send')
  })

  it('embeds emailContext into the prompt when provided', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'answer' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await c.ask({
      conversationId: 'tid',
      question: 'q',
      emailContext: 'CTX',
      chatId: 1,
    })
    const body = lastCallBody(fetchImpl) as { message: string }
    expect(body.message).toContain('[email context]')
    expect(body.message).toContain('CTX')
  })
})

describe('RealSynaplanClient — rag', () => {
  it('translates input.groups[0]/threshold to group_key/min_score', async () => {
    const fetchImpl = mockFetchSequence([
      {
        body: {
          success: true,
          results: [
            {
              id: 1,
              text: 'snip',
              score: 0.9,
              file_id: 7,
              file_name: 'a.pdf',
              group_key: 'default',
            },
          ],
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const hits = await c.ragSearch({
      query: 'find me',
      threshold: 0.7,
      limit: 5,
      groups: ['contact:alice@example.com', 'ignored-by-server'],
    })
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/rag/search')
    const body = lastCallBody(fetchImpl) as Record<string, unknown>
    expect(body).toEqual({
      query: 'find me',
      limit: 5,
      min_score: 0.7,
      group_key: 'contact:alice@example.com',
    })
    expect(hits).toEqual([
      { fileId: 7, filename: 'a.pdf', snippet: 'snip', score: 0.9, group: 'default' },
    ])
  })

  it('ragGroups normalises both wrapped and bare-array responses', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { groups: [{ key: 'default', name: 'default' }] } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.ragGroups()
    expect(r).toEqual([{ id: 'default', name: 'default', description: undefined }])
  })

  it('ragCreateGroup is a client-side reservation (no HTTP call)', async () => {
    const fetchImpl = vi.fn()
    const c = new RealSynaplanClient({
      baseUrl: 'https://api.test',
      apiKey: 'sk_t',
      fetchImpl: fetchImpl as unknown as typeof fetch,
    })
    const g = await c.ragCreateGroup('contact:bob@example.com')
    expect(g).toEqual({ id: 'contact:bob@example.com', name: 'contact:bob@example.com' })
    expect(fetchImpl).not.toHaveBeenCalled()
  })
})

describe('RealSynaplanClient — fileUpload', () => {
  it('posts multipart with files[], group_key and process_level', async () => {
    const fetchImpl = mockFetchSequence([{ body: { files: [{ id: 123 }] } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.fileUpload({
      filename: 'test.eml',
      contentBase64: btoa('hello'),
      mimeType: 'message/rfc822',
      groupId: 'contact:alice@example.com',
      processLevel: 'vectorize',
    })
    expect(r).toEqual({ fileId: 123 })
    const init = fetchImpl.mock.calls[0][1] as RequestInit
    expect(init.body).toBeInstanceOf(FormData)
    const form = init.body as FormData
    expect(form.get('group_key')).toBe('contact:alice@example.com')
    expect(form.get('process_level')).toBe('vectorize')
    // Don't set JSON content-type for multipart — the browser does it.
    expect((init.headers as Headers).get('Content-Type')).toBeNull()
  })

  it('defaults process_level to extract when omitted', async () => {
    const fetchImpl = mockFetchSequence([{ body: { files: [{ file_id: 9 }] } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.fileUpload({
      filename: 'x.eml',
      contentBase64: btoa('y'),
      mimeType: 'message/rfc822',
    })
    expect(r).toEqual({ fileId: 9 })
    const form = (fetchImpl.mock.calls[0][1] as RequestInit).body as FormData
    expect(form.get('process_level')).toBe('extract')
  })
})

describe('RealSynaplanClient — revokeApiKey', () => {
  it('DELETEs /apikeys/{id}', async () => {
    const fetchImpl = mockFetchSequence([{ status: 204 }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await c.revokeApiKey(42)
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/apikeys/42')
    const init = fetchImpl.mock.calls[0][1] as RequestInit
    expect(init.method).toBe('DELETE')
  })
})

describe('isApiError', () => {
  it('returns true for shape-conforming objects', () => {
    expect(isApiError({ status: 401, code: 'AUTH', message: 'no' })).toBe(true)
  })
  it('returns false for plain Error', () => {
    expect(isApiError(new Error('x'))).toBe(false)
  })
})
