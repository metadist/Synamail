import { describe, expect, it, vi } from 'vitest'
import {
  MockSynaplanClient,
  RealSynaplanClient,
  createSynaplanClient,
  isApiError,
} from '@shared/synaplan-client'

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

  it('hashes conversationId deterministically for chat ids', async () => {
    const c = new MockSynaplanClient(0)
    const a = await c.ask({ conversationId: 'abc', question: '?' })
    const b = await c.ask({ conversationId: 'abc', question: 'again?' })
    expect(a.chatId).toBe(b.chatId)
  })
})

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

describe('RealSynaplanClient', () => {
  it('injects the X-API-Key header on every request', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(new Response(JSON.stringify({ ok: true }), { status: 200 }))
    const c = new RealSynaplanClient({
      baseUrl: 'https://api.test',
      apiKey: 'sk_test',
      fetchImpl: fetchImpl as unknown as typeof fetch,
    })
    await c.ping()
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    const headers = (fetchImpl.mock.calls[0][1] as RequestInit).headers as Headers
    expect(headers.get('X-API-Key')).toBe('sk_test')
  })

  it('throws an ApiError with status 401 on auth failure', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(new Response('nope', { status: 401 }))
    const c = new RealSynaplanClient({
      baseUrl: 'https://api.test',
      apiKey: 'sk_bad',
      fetchImpl: fetchImpl as unknown as typeof fetch,
    })
    await expect(c.ping()).rejects.toMatchObject({ status: 401 })
  })

  it('retries 5xx with backoff up to maxRetries', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(new Response('boom', { status: 503 }))
      .mockResolvedValueOnce(new Response('boom', { status: 503 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ ok: true }), { status: 200 }))
    const c = new RealSynaplanClient({
      baseUrl: 'https://api.test',
      apiKey: 'sk_x',
      fetchImpl: fetchImpl as unknown as typeof fetch,
      maxRetries: 2,
    })
    const r = await c.ping()
    expect(r).toEqual({ ok: true })
    expect(fetchImpl).toHaveBeenCalledTimes(3)
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
