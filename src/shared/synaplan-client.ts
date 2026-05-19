/**
 * Synaplan API client.
 *
 * Sprint 2 ships the interface, the mock implementation, and a thin real
 * implementation that hits the live API. Sprint 3 replaces the mock as the
 * default and adds Zod validation generated from the OpenAPI spec.
 *
 * Selection happens via the `SYNAMAIL_USE_MOCK_CLIENT` build-time flag
 * (`vite.config.ts`) and a runtime override exposed through Settings.
 *
 * 401 handling: callers wrap calls; on `ApiError.status === 401` the
 * useAuth composable clears roaming settings and bounces to SignIn.
 */

import type {
  ApiError,
  ChatTurnInput,
  ChatTurnResult,
  ClassifyInput,
  ClassifyResult,
  DraftReplyInput,
  DraftReplyResult,
  FileUploadInput,
  FileUploadResult,
  RagGroup,
  RagSearchHit,
  RagSearchInput,
  SummariseInput,
  SummariseResult,
  TranslateInput,
  TranslateResult,
} from './types'

export interface SynaplanClient {
  ping(): Promise<{ ok: boolean; email?: string }>
  summarise(input: SummariseInput): Promise<SummariseResult>
  translate(input: TranslateInput): Promise<TranslateResult>
  draftReply(input: DraftReplyInput): Promise<DraftReplyResult>
  classify(input: ClassifyInput): Promise<ClassifyResult>
  ask(input: ChatTurnInput): Promise<ChatTurnResult>
  ragSearch(input: RagSearchInput): Promise<RagSearchHit[]>
  ragGroups(): Promise<RagGroup[]>
  ragCreateGroup(name: string): Promise<RagGroup>
  fileUpload(input: FileUploadInput): Promise<FileUploadResult>
  fileProcess(fileId: number, level: 'extract' | 'vectorize' | 'analyze'): Promise<void>
  revokeApiKey(keyId: number): Promise<void>
}

export interface ClientOptions {
  baseUrl: string
  apiKey: string
  /** Internal — for tests. */
  fetchImpl?: typeof fetch
  /** Max retries for transient 5xx. Default: 2. */
  maxRetries?: number
}

// ---------------------------------------------------------------------------
// Real client — hits web.synaplan.com (or a self-hosted Synaplan instance).
// ---------------------------------------------------------------------------

export class RealSynaplanClient implements SynaplanClient {
  private readonly baseUrl: string
  private readonly apiKey: string
  private readonly fetchImpl: typeof fetch
  private readonly maxRetries: number

  constructor(opts: ClientOptions) {
    this.baseUrl = opts.baseUrl.replace(/\/$/, '')
    this.apiKey = opts.apiKey
    this.fetchImpl = opts.fetchImpl ?? fetch.bind(globalThis)
    this.maxRetries = opts.maxRetries ?? 2
  }

  private async request<T>(path: string, init: RequestInit = {}): Promise<T> {
    const url = `${this.baseUrl}${path}`
    const headers = new Headers(init.headers)
    headers.set('X-API-Key', this.apiKey)
    if (!headers.has('Content-Type') && init.body) {
      headers.set('Content-Type', 'application/json')
    }
    headers.set('Accept', 'application/json')

    let lastErr: unknown
    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      try {
        const res = await this.fetchImpl(url, { ...init, headers })
        if (res.status === 401) {
          throw apiError(401, 'AUTH_FAILED', 'API key rejected')
        }
        if (res.status >= 500 && attempt < this.maxRetries) {
          await sleep(backoff(attempt))
          continue
        }
        if (!res.ok) {
          const text = await safeText(res)
          throw apiError(res.status, 'HTTP_ERROR', text || res.statusText)
        }
        if (res.status === 204) return undefined as T
        return (await res.json()) as T
      } catch (err) {
        lastErr = err
        if (isApiError(err) && err.status === 401) throw err
        if (attempt >= this.maxRetries) throw err
        await sleep(backoff(attempt))
      }
    }
    throw lastErr
  }

  ping(): Promise<{ ok: boolean; email?: string }> {
    return this.request('/api/v1/profile')
  }

  summarise(input: SummariseInput): Promise<SummariseResult> {
    return this.request('/api/v1/messages/send', {
      method: 'POST',
      body: JSON.stringify({ action: 'summarise', input }),
    })
  }

  translate(input: TranslateInput): Promise<TranslateResult> {
    return this.request('/api/v1/messages/send', {
      method: 'POST',
      body: JSON.stringify({ action: 'translate', input }),
    })
  }

  draftReply(input: DraftReplyInput): Promise<DraftReplyResult> {
    return this.request('/api/v1/messages/send', {
      method: 'POST',
      body: JSON.stringify({ action: 'reply', input }),
    })
  }

  classify(input: ClassifyInput): Promise<ClassifyResult> {
    return this.request('/api/v1/messages/send', {
      method: 'POST',
      body: JSON.stringify({ action: 'classify', input }),
    })
  }

  ask(input: ChatTurnInput): Promise<ChatTurnResult> {
    return this.request('/api/v1/chats', {
      method: 'POST',
      body: JSON.stringify(input),
    })
  }

  ragSearch(input: RagSearchInput): Promise<RagSearchHit[]> {
    return this.request('/api/v1/rag/search', {
      method: 'POST',
      body: JSON.stringify(input),
    })
  }

  ragGroups(): Promise<RagGroup[]> {
    return this.request('/api/v1/files/groups')
  }

  ragCreateGroup(name: string): Promise<RagGroup> {
    return this.request('/api/v1/files/groups', {
      method: 'POST',
      body: JSON.stringify({ name }),
    })
  }

  fileUpload(input: FileUploadInput): Promise<FileUploadResult> {
    return this.request('/api/v1/files/upload', {
      method: 'POST',
      body: JSON.stringify(input),
    })
  }

  fileProcess(fileId: number, level: 'extract' | 'vectorize' | 'analyze'): Promise<void> {
    return this.request(`/api/v1/files/${fileId}/process`, {
      method: 'POST',
      body: JSON.stringify({ level }),
    })
  }

  revokeApiKey(keyId: number): Promise<void> {
    return this.request(`/api/v1/apikeys/${keyId}`, { method: 'DELETE' })
  }
}

// ---------------------------------------------------------------------------
// Mock client — used in Sprint 2 and as the offline-dev default.
// Returns canned, deterministic shapes after a short simulated delay so
// the UI exercises loading + error states properly.
// ---------------------------------------------------------------------------

export class MockSynaplanClient implements SynaplanClient {
  private readonly delay: number

  constructor(delay = 250) {
    this.delay = delay
  }

  private async wait<T>(value: T): Promise<T> {
    await sleep(this.delay)
    return value
  }

  ping() {
    return this.wait({ ok: true, email: 'demo@synaplan.test' })
  }

  summarise(input: SummariseInput): Promise<SummariseResult> {
    const subject = input.subject || '(no subject)'
    return this.wait({
      summary: `(mock) Summary of "${subject}"`,
      bullets: [
        `Subject: ${subject}`,
        `From: ${input.from ?? 'unknown'}`,
        `Body length: ${input.body.length} chars`,
        'This is a mock summary — Sprint 3 swaps in the real Synaplan call.',
      ],
      language: 'en',
    })
  }

  translate(input: TranslateInput): Promise<TranslateResult> {
    return this.wait({
      translation: `(mock ${input.targetLanguage}) ${input.text}`,
      detectedLanguage: 'en',
    })
  }

  draftReply(input: DraftReplyInput): Promise<DraftReplyResult> {
    return this.wait({
      htmlBody:
        `<p>(Mock ${input.tone} reply in ${input.language})</p>` +
        `<p>Re: ${input.subject}</p>` +
        `<p>Sprint 3 replaces this with the real Synaplan response.</p>`,
    })
  }

  classify(input: ClassifyInput): Promise<ClassifyResult> {
    return this.wait({
      category: input.subject.toLowerCase().includes('invoice') ? 'billing' : 'general',
      confidence: 0.82,
      reasoning: '(mock) keyword heuristic',
    })
  }

  ask(input: ChatTurnInput): Promise<ChatTurnResult> {
    return this.wait({
      chatId: hashToChatId(input.conversationId),
      answer: `(mock answer) You asked: ${input.question}`,
    })
  }

  ragSearch(input: RagSearchInput): Promise<RagSearchHit[]> {
    return this.wait([
      {
        fileId: 1,
        filename: 'mock-quarterly-report.pdf',
        snippet: `(mock) "${input.query}" found in section 3.2 …`,
        score: 0.84,
        group: input.groups?.[0],
      },
    ])
  }

  ragGroups(): Promise<RagGroup[]> {
    return this.wait([
      { id: 'default', name: 'default' },
      { id: 'work-notes', name: 'work-notes' },
    ])
  }

  ragCreateGroup(name: string): Promise<RagGroup> {
    return this.wait({ id: name, name })
  }

  fileUpload(_input: FileUploadInput): Promise<FileUploadResult> {
    return this.wait({ fileId: Math.floor(Math.random() * 10_000) })
  }

  fileProcess(_fileId: number, _level: 'extract' | 'vectorize' | 'analyze'): Promise<void> {
    return this.wait(undefined)
  }

  revokeApiKey(_keyId: number): Promise<void> {
    return this.wait(undefined)
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

export interface CreateClientOptions {
  baseUrl: string
  apiKey: string
  useMock?: boolean
}

export function createSynaplanClient(opts: CreateClientOptions): SynaplanClient {
  if (opts.useMock || !opts.apiKey || opts.apiKey.startsWith('mock-key-')) {
    return new MockSynaplanClient()
  }
  return new RealSynaplanClient({ baseUrl: opts.baseUrl, apiKey: opts.apiKey })
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function backoff(attempt: number): number {
  return Math.min(2000, 200 * 2 ** attempt)
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms))
}

async function safeText(res: Response): Promise<string> {
  try {
    return await res.text()
  } catch {
    return ''
  }
}

function apiError(status: number, code: string, message: string): ApiError {
  return { status, code, message }
}

export function isApiError(err: unknown): err is ApiError {
  return (
    typeof err === 'object' &&
    err !== null &&
    typeof (err as ApiError).status === 'number' &&
    typeof (err as ApiError).code === 'string'
  )
}

function hashToChatId(s: string): number {
  let h = 0
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) | 0
  }
  return Math.abs(h) || 1
}
