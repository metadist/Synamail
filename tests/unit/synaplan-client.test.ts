import { describe, expect, it, vi } from 'vitest'
import {
  RealSynaplanClient,
  cleanEmailText,
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

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  })
}

/** An SSE response whose body streams the given `data:` frames then closes. */
function sseResponse(frames: string[]): Response {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(new TextEncoder().encode(frames.join('')))
      controller.close()
    },
  })
  return new Response(stream, {
    status: 200,
    headers: { 'Content-Type': 'text/event-stream' },
  })
}

function sseFrame(payload: Record<string, unknown>): string {
  return `data: ${JSON.stringify(payload)}\n\n`
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

describe('createSynaplanClient factory', () => {
  it('always returns a RealSynaplanClient (no mock mode)', () => {
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

  it('composeDraft routes through the chat stream pipeline (not /messages/send)', async () => {
    // Full pipeline = create a chat, then stream — so Synaplan's classifier
    // runs and web search follows the user's config. `/messages/send` bypasses
    // all of that, so it must NOT be used here.
    const fetchImpl = mockFetchSequence([{ body: { success: true, chat: { id: 55 } } }])
    fetchImpl.mockResolvedValueOnce(
      sseResponse([
        sseFrame({ status: 'data', chunk: '<p>Hi Alice, lunch Friday?</p>' }),
        sseFrame({ status: 'complete' }),
      ]),
    )
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.composeDraft({
      intent: 'Invite Alice to lunch on Friday',
      tone: 'friendly',
      language: 'en',
    })
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/chats')
    const streamUrl = String(fetchImpl.mock.calls[1][0])
    expect(streamUrl).toContain('/api/v1/messages/stream')
    const message = new URL(streamUrl).searchParams.get('message') ?? ''
    expect(message).toContain('[intent]')
    expect(message).toContain('Invite Alice to lunch on Friday')
    expect(message).toMatch(/friendly/)
    expect(r.htmlBody).toBe('<p>Hi Alice, lunch Friday?</p>')
  })

  it('composeDraft includes the replied-to body when composing a reply', async () => {
    const fetchImpl = mockFetchSequence([{ body: { success: true, chat: { id: 7 } } }])
    fetchImpl.mockResolvedValueOnce(
      sseResponse([
        sseFrame({ status: 'data', chunk: '<p>Sounds good.</p>' }),
        sseFrame({ status: 'complete' }),
      ]),
    )
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await c.composeDraft({
      intent: 'Accept the proposal',
      tone: 'concise',
      language: 'en',
      referenceBody: 'Can we meet Tuesday?',
    })
    const message = new URL(String(fetchImpl.mock.calls[1][0])).searchParams.get('message') ?? ''
    expect(message).toContain('[replying to]')
    expect(message).toContain('Can we meet Tuesday?')
  })

  it('tts fetches the configured TTS stream and returns an audio blob', async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(
      new Response(new Uint8Array([1, 2, 3]), {
        status: 200,
        headers: { 'Content-Type': 'audio/mpeg' },
      }),
    )
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const blob = await c.tts({ text: 'Hello world', language: 'en' })
    const url = new URL(String(fetchImpl.mock.calls[0][0]))
    expect(url.pathname).toBe('/api/v1/tts/stream')
    expect(url.searchParams.get('text')).toBe('Hello world')
    expect(url.searchParams.get('language')).toBe('en')
    expect(blob.type).toBe('audio/mpeg')
    expect(blob.size).toBe(3)
  })

  it('chat surfaces generated media from stream file/audio events as absolute URLs', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        sseResponse([
          sseFrame({ status: 'data', chunk: 'Generated image: a cat' }),
          sseFrame({ status: 'file', type: 'image', url: '/api/v1/files/uploads/13/x.png' }),
          sseFrame({ status: 'audio', url: '/api/v1/files/uploads/13/tts_1.mp3' }),
          sseFrame({ status: 'complete' }),
        ]),
      )
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.chat({ conversationId: 'home', question: '/pic a cat', chatId: 5 }, () => {})
    expect(r.answer).toContain('Generated image: a cat')
    expect(r.media).toEqual([
      { kind: 'image', url: 'https://api.test/api/v1/files/uploads/13/x.png' },
      { kind: 'audio', url: 'https://api.test/api/v1/files/uploads/13/tts_1.mp3' },
    ])
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

  it('forwards attached fileIds (email images) to /messages/send', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'I see wines in a basket' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.ask({
      conversationId: 'thread-img',
      question: 'what is in this screenshot?',
      chatId: 5,
      fileIds: [101, 102],
    })
    expect(r.answer).toBe('I see wines in a basket')
    const body = lastCallBody(fetchImpl) as { trackId: number; fileIds: number[] }
    expect(body.trackId).toBe(5)
    expect(body.fileIds).toEqual([101, 102])
  })

  it('omits fileIds from the body when none are attached', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'ok' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await c.ask({ conversationId: 't', question: 'q', chatId: 1 })
    const body = lastCallBody(fetchImpl) as Record<string, unknown>
    expect('fileIds' in body).toBe(false)
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
  it('translates input.groups[0]/threshold to group_key/min_score and maps the real hit shape', async () => {
    // Real wire shape verified against the live API on 2026-06-04:
    // results carry { chunk_id, message_id, text, score, start_line, end_line }
    // — NOT the file_id/file_name/group_key the OpenAPI annotation advertises.
    const fetchImpl = mockFetchSequence([
      {
        body: {
          success: true,
          query: 'find me',
          results: [
            {
              chunk_id: 'doc_7_1_0',
              message_id: 7,
              text: 'snip',
              score: 0.9,
              start_line: 0,
              end_line: 0,
            },
          ],
          total_results: 1,
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
    // message_id → fileId; no filename in the hit; group filled from the request.
    expect(hits).toEqual([
      {
        fileId: 7,
        filename: '',
        snippet: 'snip',
        score: 0.9,
        group: 'contact:alice@example.com',
      },
    ])
  })

  it('still tolerates the legacy/spec-advertised file_id/file_name/group_key fields', async () => {
    const fetchImpl = mockFetchSequence([
      {
        body: {
          success: true,
          results: [
            { id: 1, text: 'snip', score: 0.8, file_id: 7, file_name: 'a.pdf', group_key: 'g1' },
          ],
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const hits = await c.ragSearch({ query: 'q', groups: ['g-req'] })
    expect(hits).toEqual([
      { fileId: 7, filename: 'a.pdf', snippet: 'snip', score: 0.8, group: 'g1' },
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
      filename: 'x.txt',
      contentBase64: btoa('y'),
      mimeType: 'text/plain',
    })
    expect(r).toEqual({ fileId: 9 })
    const form = (fetchImpl.mock.calls[0][1] as RequestInit).body as FormData
    expect(form.get('process_level')).toBe('extract')
  })

  it('throws the server error when the upload is rejected (no file id)', async () => {
    // Synaplan answers 206 with success:false + errors for an unsupported type.
    const fetchImpl = mockFetchSequence([
      {
        status: 206,
        body: {
          success: false,
          files: [],
          errors: [
            { filename: 'mail.eml', error: 'File type not allowed. Allowed: pdf, docx, txt' },
          ],
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await expect(
      c.fileUpload({ filename: 'mail.eml', contentBase64: btoa('z'), mimeType: 'message/rfc822' }),
    ).rejects.toMatchObject({ code: 'UPLOAD_FAILED', message: /File type not allowed/ })
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

describe('RealSynaplanClient — extractMeetingTimes', () => {
  it('parses a JSON array of proposals and defaults a missing end to +30min', async () => {
    const fetchImpl = mockFetchSequence([
      {
        body: {
          success: true,
          outgoingMessage: {
            text: '[{"title":"Project call","start":"2026-06-03T15:00:00","end":"2026-06-03T15:30:00","location":"Teams"},{"title":"Sync","start":"2026-06-04T09:00:00"}]',
          },
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.extractMeetingTimes({
      subject: 'Can we meet?',
      body: 'How about Wednesday 3pm or Thursday 9am?',
      nowIso: '2026-06-01T09:00:00',
      timezone: 'Europe/Berlin',
    })
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/messages/send')
    expect(r).toEqual([
      {
        title: 'Project call',
        startIso: '2026-06-03T15:00:00',
        endIso: '2026-06-03T15:30:00',
        location: 'Teams',
      },
      {
        title: 'Sync',
        startIso: '2026-06-04T09:00:00',
        endIso: '2026-06-04T09:30:00',
        location: undefined,
      },
    ])
  })

  it('returns [] when the AI reports no proposals', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: '[]' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.extractMeetingTimes({
      subject: 'FYI',
      body: 'No times here.',
      nowIso: '2026-06-01T09:00:00',
      timezone: 'UTC',
    })
    expect(r).toEqual([])
  })

  it('returns [] on unparseable output rather than throwing', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'sorry, no idea' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.extractMeetingTimes({
      subject: 'x',
      body: 'y',
      nowIso: '2026-06-01T09:00:00',
      timezone: 'UTC',
    })
    expect(r).toEqual([])
  })
})

describe('RealSynaplanClient — Contact AI Profiling (synamail plugin)', () => {
  const me = { body: { success: true, user: { email: 'a@b.test', id: 7 } } }
  const profile = {
    email: 'alice@example.com',
    summary: 'Alice runs procurement at Example Corp.',
    tone: 'friendly',
    facts: ['Works at Example Corp'],
    openLoops: ['me: send the demo'],
    emailCount: 3,
    updatedAt: '2026-06-10T12:00:00+00:00',
  }

  it('resolves the user id from /auth/me, then GETs the plugin profile route', async () => {
    const fetchImpl = mockFetchSequence([me, { body: { success: true, profile } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.getContactProfile('Alice@Example.com')
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/auth/me')
    expect(fetchImpl.mock.calls[1][0]).toBe(
      'https://api.test/api/v1/user/7/plugins/synamail/profiles/alice%40example.com',
    )
    expect(r).toEqual(profile)
  })

  it('returns null when no profile exists yet, and caches the user id', async () => {
    const fetchImpl = mockFetchSequence([
      me,
      { body: { success: true, profile: null } },
      { body: { success: true, profile: null } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    expect(await c.getContactProfile('alice@example.com')).toBeNull()
    expect(await c.getContactProfile('alice@example.com')).toBeNull()
    // auth/me once + two profile GETs — the id is cached after the first call.
    expect(fetchImpl).toHaveBeenCalledTimes(3)
  })

  it('updateContactProfile POSTs the email payload and returns the new snapshot', async () => {
    const fetchImpl = mockFetchSequence([me, { body: { success: true, profile } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.updateContactProfile({
      email: 'alice@example.com',
      subject: 'Re: demo',
      body: 'Thanks, talk soon!',
      direction: 'inbound',
      name: 'Alice',
    })
    expect(fetchImpl.mock.calls[1][0]).toBe(
      'https://api.test/api/v1/user/7/plugins/synamail/profiles/alice%40example.com/update',
    )
    const body = lastCallBody(fetchImpl) as Record<string, unknown>
    expect(body.direction).toBe('inbound')
    expect(body.subject).toBe('Re: demo')
    expect(body.name).toBe('Alice')
    expect(r).toEqual(profile)
  })

  it('maps a 404 (plugin not installed) to PROFILING_UNAVAILABLE', async () => {
    const fetchImpl = mockFetchSequence([me, { status: 404, body: 'No route found' }])
    const c = buildClient(fetchImpl as unknown as typeof fetch, { maxRetries: 0 })
    await expect(c.getContactProfile('alice@example.com')).rejects.toMatchObject({
      code: 'PROFILING_UNAVAILABLE',
    })
  })

  it('deleteContactProfile issues a DELETE on the profile route', async () => {
    const fetchImpl = mockFetchSequence([me, { body: { success: true, deleted: true } }])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    await c.deleteContactProfile('alice@example.com')
    const [url, init] = fetchImpl.mock.calls[1] as [string, RequestInit]
    expect(url).toBe('https://api.test/api/v1/user/7/plugins/synamail/profiles/alice%40example.com')
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

describe('RealSynaplanClient — streaming', () => {
  it('streams chunks and returns the accumulated text', async () => {
    const fetchImpl = vi.fn()
    fetchImpl.mockResolvedValueOnce(jsonResponse({ success: true, chat: { id: 7 } }))
    fetchImpl.mockResolvedValueOnce(
      sseResponse([
        sseFrame({ status: 'data', chunk: 'Hola ' }),
        sseFrame({ status: 'data', chunk: 'mundo' }),
        sseFrame({ status: 'complete', messageId: 1 }),
      ]),
    )
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const seen: string[] = []
    const r = await c.translate({ text: 'Hello world', targetLanguage: 'es' }, (t) => seen.push(t))

    expect(r.translation).toBe('Hola mundo')
    expect(seen).toEqual(['Hola ', 'Hola mundo'])
    const [url, init] = fetchImpl.mock.calls[1] as [string, RequestInit]
    expect(url).toContain('/api/v1/messages/stream')
    expect(url).toContain('chatId=7')
    expect((init.headers as Headers).get('X-API-Key')).toBe('sk_test')
  })

  it('falls back to the blocking endpoint when the stream request fails', async () => {
    const fetchImpl = vi.fn()
    fetchImpl.mockResolvedValueOnce(jsonResponse({ success: true, chat: { id: 9 } }))
    fetchImpl.mockResolvedValueOnce(new Response('nope', { status: 404 }))
    fetchImpl.mockResolvedValueOnce(
      jsonResponse({ success: true, outgoingMessage: { text: 'fallback text' } }),
    )
    const c = buildClient(fetchImpl as unknown as typeof fetch, { maxRetries: 0 })
    const r = await c.translate({ text: 'x', targetLanguage: 'de' }, () => {})

    expect(r.translation).toBe('fallback text')
    expect(fetchImpl).toHaveBeenCalledTimes(3)
    expect(fetchImpl.mock.calls[2][0] as string).toContain('/api/v1/messages/send')
  })

  it('does not stream (no chat created) when no onChunk is given', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'blocking' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.translate({ text: 'x', targetLanguage: 'fr' })
    expect(r.translation).toBe('blocking')
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    expect(fetchImpl.mock.calls[0][0] as string).toContain('/api/v1/messages/send')
  })
})

describe('cleanEmailText', () => {
  it('strips angle-bracket autolinks and long tracking URLs', () => {
    const dirty = `Read more <https://nl.nytimes.com/q/${'A'.repeat(80)}> now`
    const clean = cleanEmailText(dirty)
    expect(clean).not.toContain('http')
    expect(clean).not.toContain('<')
    expect(clean).toContain('Read more')
    expect(clean).toContain('now')
  })

  it('does not eat text that follows a stripped autolink', () => {
    const clean = cleanEmailText(`<https://track.example.com/${'x'.repeat(50)}>KEEPME`)
    expect(clean).toContain('KEEPME')
    expect(clean).not.toContain('track.example.com')
  })

  it('keeps short, human-meaningful URLs', () => {
    expect(cleanEmailText('see https://syn.io')).toContain('https://syn.io')
  })

  it('collapses runs of blank lines', () => {
    expect(cleanEmailText('a\n\n\n\nb')).toBe('a\n\nb')
  })
})
