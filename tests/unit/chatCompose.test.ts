import { describe, expect, it, vi } from 'vitest'
import { RealSynaplanClient, errorMessage } from '@shared/synaplan-client'

describe('errorMessage', () => {
  it('renders ApiError objects (not "[object Object]")', () => {
    expect(
      errorMessage({ status: 500, code: 'AI_FAILED', message: 'Model must be specified' }),
    ).toBe('Model must be specified')
  })
  it('handles Error, string and falls back to JSON', () => {
    expect(errorMessage(new Error('boom'))).toBe('boom')
    expect(errorMessage('plain')).toBe('plain')
    expect(errorMessage({ a: 1 })).toBe('{"a":1}')
  })
})

interface MockResponseStep {
  status?: number
  body?: unknown
}

function mockFetchSequence(steps: MockResponseStep[]): ReturnType<typeof vi.fn> {
  const fn = vi.fn()
  for (const step of steps) {
    const status = step.status ?? 200
    const body = typeof step.body === 'string' ? step.body : JSON.stringify(step.body ?? {})
    fn.mockResolvedValueOnce(
      new Response(body, { status, headers: { 'Content-Type': 'application/json' } }),
    )
  }
  return fn
}

function buildClient(fetchImpl: typeof fetch): RealSynaplanClient {
  return new RealSynaplanClient({
    baseUrl: 'https://api.test',
    apiKey: 'sk_test',
    fetchImpl,
    maxRetries: 0,
  })
}

describe('RealSynaplanClient.chat', () => {
  it('creates a chat then sends the simple-chat prompt when no chatId is given', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, chat: { id: 7 } } },
      { body: { success: true, outgoingMessage: { text: 'hello there' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.chat({ conversationId: 'home', question: 'hi' })
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/chats')
    expect(fetchImpl.mock.calls[1][0]).toBe('https://api.test/api/v1/messages/send')
    const sendBody = JSON.parse((fetchImpl.mock.calls[1][1] as RequestInit).body as string)
    expect(sendBody.trackId).toBe(7)
    expect(sendBody.message).toMatch(/Synaplan/i)
    expect(sendBody.message).toContain('hi')
    expect(r).toEqual({ chatId: 7, answer: 'hello there' })
  })

  it('reuses an existing chatId without creating a new chat', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, outgoingMessage: { text: 'again' } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.chat({ conversationId: 'home', question: 'q', chatId: 99 })
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/messages/send')
    expect(r.chatId).toBe(99)
  })
})

describe('RealSynaplanClient.getModelConfig', () => {
  it('joins defaults ids with the catalog to resolve names', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, defaults: { CHAT: 53, TEXT2PIC: 21, VECTORIZE: 3 } } },
      {
        body: {
          success: true,
          models: {
            CHAT: [{ id: 53, name: 'Llama 3.3 70B', service: 'Groq' }],
            TEXT2PIC: [{ id: 21, name: 'FLUX.1', service: 'Replicate' }],
            VECTORIZE: [{ id: 3, name: 'nomic-embed-text', service: 'Ollama' }],
          },
        },
      },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.getModelConfig()
    expect(fetchImpl.mock.calls[0][0]).toBe('https://api.test/api/v1/config/models/defaults')
    expect(fetchImpl.mock.calls[1][0]).toBe('https://api.test/api/v1/config/models')
    expect(r.chat).toEqual({ id: 53, name: 'Llama 3.3 70B', service: 'Groq' })
    expect(r.imageGen?.name).toBe('FLUX.1')
    expect(r.vectorize?.name).toBe('nomic-embed-text')
  })

  it('returns null for capabilities with no configured model', async () => {
    const fetchImpl = mockFetchSequence([
      { body: { success: true, defaults: { CHAT: 53, TEXT2PIC: null, VECTORIZE: 3 } } },
      { body: { success: true, models: { CHAT: [{ id: 53, name: 'Chat', service: 'X' }] } } },
    ])
    const c = buildClient(fetchImpl as unknown as typeof fetch)
    const r = await c.getModelConfig()
    expect(r.chat?.name).toBe('Chat')
    expect(r.imageGen).toBeNull()
    // id present but missing from catalog → falls back to a "#id" label.
    expect(r.vectorize).toEqual({ id: 3, name: '#3' })
  })
})
