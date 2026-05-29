/**
 * Synaplan API client.
 *
 * The interface (`SynaplanClient`) is shaped around the taskpane's
 * read-mode / compose-mode features. Two implementations:
 *
 *   - `RealSynaplanClient` — hits a live Synaplan instance.
 *   - `MockSynaplanClient` — returns canned data for offline dev / Vitest.
 *
 * Selection happens via `createSynaplanClient({ useMock })` and the auto-
 * detection of mock keys (prefix `mock-key-`). Sprint 3 swaps the default
 * to the real client.
 *
 * All AI features (summarise / translate / draftReply / classify / ask)
 * route through Synaplan's single `POST /api/v1/messages/send` chat
 * endpoint with a per-action system prompt prepended to the email body.
 * Synaplan's `outgoingMessage` response field carries the AI's reply
 * text; its exact internal shape is `type: object` (undocumented) in the
 * OpenAPI spec, so `extractAiText()` tries a handful of common field
 * names. See the comment on that helper for the live-smoke checklist.
 *
 * 401 handling: callers wrap calls; on `ApiError.status === 401` the
 * useAuth composable clears roaming settings and bounces to SignIn.
 */

import {
  ask as askPrompt,
  classify as classifyPrompt,
  newMail as newMailPrompt,
  reply as replyPrompt,
  simpleChat as simpleChatPrompt,
  summarise as summarisePrompt,
  translate as translatePrompt,
} from './prompts'
import type {
  ApiError,
  ChatTurnInput,
  ChatTurnResult,
  ClassifyInput,
  ClassifyResult,
  ComposeNewInput,
  ComposeNewResult,
  CreateSpamRuleInput,
  CreateSpamRuleResult,
  DraftReplyInput,
  DraftReplyResult,
  FileUploadInput,
  FileUploadResult,
  ModelConfig,
  RagGroup,
  RagSearchHit,
  RagSearchInput,
  SenderHistoryInput,
  SenderHistoryResult,
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
  /**
   * General-purpose chat, not tied to a specific email. Mirrors `ask` for
   * chat-id persistence (caller stores the returned `chatId` in roaming
   * under a stable key such as `home`), but uses the simple-chat system
   * prompt instead of the email-grounded one.
   */
  chat(input: ChatTurnInput): Promise<ChatTurnResult>
  /** Draft a brand-new email (subject + HTML body) from a description. */
  composeNew(input: ComposeNewInput): Promise<ComposeNewResult>
  ragSearch(input: RagSearchInput): Promise<RagSearchHit[]>
  ragGroups(): Promise<RagGroup[]>
  /**
   * Returns a synthetic RagGroup with the requested name. Synaplan creates
   * groups implicitly when an upload uses a new `group_key`, so this call
   * is a client-side reservation — the group materialises server-side as
   * soon as `fileUpload({ groupId: name })` runs.
   */
  ragCreateGroup(name: string): Promise<RagGroup>
  fileUpload(input: FileUploadInput): Promise<FileUploadResult>
  revokeApiKey(keyId: number): Promise<void>

  /**
   * The user's currently-configured models for Chat, Image generation
   * (TEXT2PIC) and Vectorization (VECTORIZE). Resolved by joining
   * `GET /config/models/defaults` (ids) with `GET /config/models` (catalog).
   */
  getModelConfig(): Promise<ModelConfig>

  /**
   * List recent messages from a single sender. The data source depends on
   * the implementation:
   *   - Real client (Sprint 4): asks Synaplan's saved-email index AND,
   *     when Outlook is reachable, fans out to EWS FindItem for live
   *     mailbox results.
   *   - Mock client: returns canned data shaped like a real inbox so the
   *     UI can be designed end-to-end without a backend round-trip.
   */
  senderHistory(input: SenderHistoryInput): Promise<SenderHistoryResult>

  /**
   * Create a "block sender" inbox rule. The real client routes through
   * Outlook (EWS CreateInboxRules / Graph messageRules) when available so
   * the rule actually filters future mail; the mock returns a stable id
   * for dev iteration.
   */
  createSpamRule(input: CreateSpamRuleInput): Promise<CreateSpamRuleResult>
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
// Real client — hits a live Synaplan instance (`web.synaplan.com` by
// default, overridable for self-hosted instances).
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
    headers.set('Accept', 'application/json')
    // Only set JSON content-type when the caller hasn't supplied a body
    // shape of its own (e.g. multipart for /files/upload).
    if (!headers.has('Content-Type') && init.body && typeof init.body === 'string') {
      headers.set('Content-Type', 'application/json')
    }

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

  // -------------------------------------------------------------------------
  // Ping — uses GET /auth/me, which is purpose-built for "who am I" and
  // returns a compact, stable `user.email` (vs /profile which returns the
  // larger billing-oriented record).
  // -------------------------------------------------------------------------

  async ping(): Promise<{ ok: boolean; email?: string }> {
    const res = await this.request<AuthMeResponse>('/api/v1/auth/me')
    return { ok: !!res?.success, email: res?.user?.email }
  }

  // -------------------------------------------------------------------------
  // AI actions — every one of these is a single POST /messages/send call
  // whose `message` field is the system prompt prepended to the email body.
  // The AI's reply comes back in `outgoingMessage` (undocumented shape;
  // extracted via extractAiText below).
  // -------------------------------------------------------------------------

  async summarise(input: SummariseInput): Promise<SummariseResult> {
    const lang = pickLanguage(input)
    const message = composeMessage(summarisePrompt(lang), buildEmailBlock(input))
    const text = await this.sendChat(message)
    return {
      summary: text,
      bullets: parseMarkdownBullets(text),
      language: lang,
    }
  }

  async translate(input: TranslateInput): Promise<TranslateResult> {
    const message = composeMessage(translatePrompt(input.targetLanguage), input.text)
    const text = await this.sendChat(message)
    return {
      translation: text,
      // Detection isn't returned by the server in this call shape; we record
      // the target so downstream UI has something deterministic to show.
      detectedLanguage: 'auto',
    }
  }

  async draftReply(input: DraftReplyInput): Promise<DraftReplyResult> {
    const message = composeMessage(
      replyPrompt(input.tone, input.language),
      buildEmailBlock({
        subject: input.subject,
        body: input.body,
      }),
    )
    const text = await this.sendChat(message)
    return { htmlBody: text }
  }

  async classify(input: ClassifyInput): Promise<ClassifyResult> {
    const categories = ['billing', 'support', 'internal', 'personal', 'spam', 'general']
    const message = composeMessage(classifyPrompt(categories), buildEmailBlock(input))
    const text = await this.sendChat(message)
    return parseClassifyResponse(text)
  }

  async ask(input: ChatTurnInput): Promise<ChatTurnResult> {
    // 1. Ensure a chat exists for this conversation. If the caller didn't
    //    supply chatId from roaming, create a fresh chat now and return its
    //    id alongside the answer so the caller can persist it.
    let chatId = input.chatId
    if (!chatId) {
      const created = await this.request<CreateChatResponse>('/api/v1/chats', {
        method: 'POST',
        body: JSON.stringify({ title: `Outlook: ${input.conversationId.slice(0, 40)}` }),
      })
      if (!created?.success || !created.chat?.id) {
        throw apiError(0, 'CHAT_CREATE_FAILED', 'Could not create chat')
      }
      chatId = created.chat.id
    }
    // 2. Send the question into that chat via /messages/send with trackId.
    const prefix = input.emailContext
      ? `${askPrompt()}\n\n[email context]\n${input.emailContext}\n\n[question]\n`
      : `${askPrompt()}\n\n`
    const answer = await this.sendChat(`${prefix}${input.question}`, chatId)
    return { chatId, answer }
  }

  async chat(input: ChatTurnInput): Promise<ChatTurnResult> {
    let chatId = input.chatId
    if (!chatId) {
      const created = await this.request<CreateChatResponse>('/api/v1/chats', {
        method: 'POST',
        body: JSON.stringify({ title: 'Synamail chat' }),
      })
      if (!created?.success || !created.chat?.id) {
        throw apiError(0, 'CHAT_CREATE_FAILED', 'Could not create chat')
      }
      chatId = created.chat.id
    }
    const answer = await this.sendChat(`${simpleChatPrompt()}\n\n${input.question}`, chatId)
    return { chatId, answer }
  }

  async composeNew(input: ComposeNewInput): Promise<ComposeNewResult> {
    const message = composeMessage(newMailPrompt(input.language ?? 'en'), input.description)
    const text = await this.sendChat(message)
    return parseComposeResponse(text, input.description)
  }

  // -------------------------------------------------------------------------
  // RAG
  // -------------------------------------------------------------------------

  async ragSearch(input: RagSearchInput): Promise<RagSearchHit[]> {
    // Synaplan's /rag/search takes one group_key; multi-group search would
    // need to be fanned out client-side. V1 just uses the first group when
    // multiple are provided (matches how the contact-KB feature uses it).
    const body: Record<string, unknown> = { query: input.query }
    if (input.limit !== undefined) body.limit = input.limit
    if (input.threshold !== undefined) body.min_score = input.threshold
    if (input.groups && input.groups.length > 0) body.group_key = input.groups[0]

    const res = await this.request<RagSearchResponse>('/api/v1/rag/search', {
      method: 'POST',
      body: JSON.stringify(body),
    })
    return (res?.results ?? []).map((hit) => ({
      fileId: hit.file_id ?? 0,
      filename: hit.file_name ?? '',
      snippet: hit.text ?? '',
      score: hit.score ?? 0,
      group: hit.group_key ?? undefined,
    }))
  }

  async ragGroups(): Promise<RagGroup[]> {
    const res = await this.request<RagGroupsResponse>('/api/v1/files/groups')
    // The endpoint returns either `{groups: [...]}` or a bare array depending
    // on the deployment; normalise both.
    const list: RawRagGroup[] = Array.isArray(res) ? res : (res?.groups ?? [])
    return list.map((g) => ({
      id: g.key ?? g.id ?? g.name ?? '',
      name: g.name ?? g.key ?? '',
      description: g.description ?? undefined,
    }))
  }

  async ragCreateGroup(name: string): Promise<RagGroup> {
    // No backend endpoint exists for explicit group creation; groups are
    // materialised on first upload with a new `group_key`. We return a
    // synthetic descriptor so the picker can list the name immediately.
    return { id: name, name }
  }

  async fileUpload(input: FileUploadInput): Promise<FileUploadResult> {
    const blob = base64ToBlob(input.contentBase64, input.mimeType)
    const form = new FormData()
    form.append('files[]', blob, input.filename)
    if (input.groupId) form.append('group_key', input.groupId)
    form.append('process_level', input.processLevel ?? 'extract')

    const res = await this.request<FileUploadResponse>('/api/v1/files/upload', {
      method: 'POST',
      body: form,
    })
    // The response shape varies (`files[0].id` vs `fileIds[0]` vs `file_id`);
    // pick the first id we can find so callers get a stable `FileUploadResult`.
    const fileId = pickFirstFileId(res)
    return { fileId }
  }

  async revokeApiKey(keyId: number): Promise<void> {
    await this.request(`/api/v1/apikeys/${keyId}`, { method: 'DELETE' })
  }

  async getModelConfig(): Promise<ModelConfig> {
    const [defaults, catalog] = await Promise.all([
      this.request<ConfigDefaultsResponse>('/api/v1/config/models/defaults'),
      this.request<ConfigModelsResponse>('/api/v1/config/models'),
    ])
    const resolve = (capability: 'CHAT' | 'TEXT2PIC' | 'VECTORIZE') => {
      const id = defaults?.defaults?.[capability]
      if (typeof id !== 'number') return null
      const model = (catalog?.models?.[capability] ?? []).find((m) => m.id === id)
      if (!model) return { id, name: `#${id}` }
      return { id, name: model.name ?? `#${id}`, service: model.service ?? undefined }
    }
    return {
      chat: resolve('CHAT'),
      imageGen: resolve('TEXT2PIC'),
      vectorize: resolve('VECTORIZE'),
    }
  }

  // -------------------------------------------------------------------------
  // Sender-centric features (Sprint 4 — placeholder stubs)
  //
  // These hit two surfaces in the final implementation:
  //   1. Synaplan's saved-email index, for context the user already chose
  //      to remember (Save to knowledge base feeds it).
  //   2. The user's live mailbox via EWS / Graph, for messages that never
  //      reached Synaplan. The Outlook side belongs in a useOutlookMailbox
  //      composable to keep this client pure-Synaplan; for now we throw a
  //      typed error so the mock-key dev loop hits MockSynaplanClient and
  //      the live path doesn't silently return empty.
  // -------------------------------------------------------------------------

  async senderHistory(_input: SenderHistoryInput): Promise<SenderHistoryResult> {
    throw apiError(
      501,
      'NOT_IMPLEMENTED',
      'senderHistory lands in Sprint 4 once the EWS bridge composable ships.',
    )
  }

  async createSpamRule(_input: CreateSpamRuleInput): Promise<CreateSpamRuleResult> {
    throw apiError(
      501,
      'NOT_IMPLEMENTED',
      'createSpamRule lands in Sprint 4 once the EWS bridge composable ships.',
    )
  }

  // -------------------------------------------------------------------------
  // Internal: AI chat round-trip
  // -------------------------------------------------------------------------

  private async sendChat(message: string, trackId?: number): Promise<string> {
    const body: Record<string, unknown> = { message }
    if (trackId !== undefined) body.trackId = trackId
    const res = await this.request<MessagesSendResponse>('/api/v1/messages/send', {
      method: 'POST',
      body: JSON.stringify(body),
    })
    if (!res?.success) {
      throw apiError(0, 'AI_FAILED', 'Synaplan reported a non-success response')
    }
    return extractAiText(res.outgoingMessage)
  }
}

// ---------------------------------------------------------------------------
// Mock client — used for offline-dev / Vitest. Returns canned, deterministic
// shapes after a short simulated delay so the UI exercises loading + error
// states properly.
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
      chatId: input.chatId ?? hashToChatId(input.conversationId),
      answer: `(mock answer) You asked: ${input.question}`,
    })
  }

  chat(input: ChatTurnInput): Promise<ChatTurnResult> {
    return this.wait({
      chatId: input.chatId ?? hashToChatId(input.conversationId),
      answer: `(mock chat) Re "${input.question}": this is a canned reply — sign in to a real Synaplan instance for live answers.`,
    })
  }

  composeNew(input: ComposeNewInput): Promise<ComposeNewResult> {
    const subject = input.description.split(/\r?\n/)[0]?.slice(0, 60) || 'New message'
    return this.wait({
      subject: `(mock) ${subject}`,
      htmlBody:
        `<p>(Mock draft for: ${subject})</p>` +
        `<p>Sign in to a real Synaplan instance to generate a live draft from your description.</p>`,
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

  revokeApiKey(_keyId: number): Promise<void> {
    return this.wait(undefined)
  }

  getModelConfig(): Promise<ModelConfig> {
    return this.wait({
      chat: { id: 53, name: 'Llama 3.3 70B', service: 'Groq' },
      imageGen: { id: 21, name: 'FLUX.1 schnell', service: 'Replicate' },
      vectorize: { id: 3, name: 'nomic-embed-text', service: 'Ollama' },
    })
  }

  senderHistory(input: SenderHistoryInput): Promise<SenderHistoryResult> {
    const limit = Math.max(1, Math.min(input.limit ?? 25, 50))
    const subjects = [
      'Re: Q3 review prep',
      'Lunch on Thursday?',
      'Re: vendor onboarding docs',
      'FYI: pricing changes effective May',
      'Quick question about the brief',
      'Re: photoshoot timeline',
      'Calendar invite — sync next week',
      'Updated proposal attached',
    ]
    const baseDate = Date.now()
    // Synthesise a plausible-looking history so the UI exercises pagination,
    // unread styling, and the multi-message summary action.
    const items = Array.from({ length: limit }, (_, i) => {
      const dayOffset = (i + 1) * 3
      return {
        date: new Date(baseDate - dayOffset * 86_400_000).toISOString(),
        subject: subjects[i % subjects.length],
        snippet: `(mock) Latest from ${input.email} — message #${i + 1} preview text…`,
        unread: i < 2,
        messageId: `mock-msg-${i + 1}`,
      }
    })
    return this.wait({
      email: input.email,
      total: items.length,
      items,
      fromOutlook: false,
    })
  }

  createSpamRule(input: CreateSpamRuleInput): Promise<CreateSpamRuleResult> {
    return this.wait({
      ruleId: `mock-rule-${Math.random().toString(36).slice(2, 8)}`,
      movedCount: input.alsoCleanExisting ? Math.floor(Math.random() * 6) : 0,
      serverSide: false,
    })
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

/**
 * Human-readable message for any thrown value. Crucially handles the plain
 * `ApiError` objects this client throws (which are NOT `Error` instances, so
 * `String(err)` would render the useless "[object Object]").
 */
export function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message
  if (isApiError(err)) return err.message || err.code || `Request failed (${err.status})`
  if (typeof err === 'string') return err
  try {
    return JSON.stringify(err)
  } catch {
    return String(err)
  }
}

function hashToChatId(s: string): number {
  let h = 0
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) | 0
  }
  return Math.abs(h) || 1
}

function composeMessage(systemPrompt: string, userBlock: string): string {
  return `${systemPrompt}\n\n${userBlock}`
}

function buildEmailBlock(input: {
  subject?: string
  body: string
  from?: string
  to?: string[]
}): string {
  const parts: string[] = []
  if (input.subject) parts.push(`Subject: ${input.subject}`)
  if (input.from) parts.push(`From: ${input.from}`)
  if (input.to && input.to.length) parts.push(`To: ${input.to.join(', ')}`)
  parts.push('', input.body)
  return parts.join('\n')
}

function pickLanguage(input: SummariseInput | ClassifyInput): string {
  // Summarise / classify don't carry a target language in their input shape
  // today; default to English. Sprint 3.x adds a user override hooked off
  // Settings.language and Office.context.displayLanguage.
  void input
  return 'en'
}

function parseMarkdownBullets(text: string): string[] {
  // The summarise prompt asks for `- ` bullets. Extract the leaf text from
  // any line that starts with `-`, `*`, or `•` (with leading whitespace),
  // stripping the marker. Falls back to splitting on newlines.
  const lines = text.split(/\r?\n/)
  const bullets = lines
    .map((l) => l.match(/^\s*(?:[-*•]|\d+[.)])\s+(.+)$/)?.[1] ?? null)
    .filter((b): b is string => !!b)
    .map((b) => b.trim())
  if (bullets.length > 0) return bullets
  return lines.map((l) => l.trim()).filter(Boolean)
}

function parseClassifyResponse(text: string): ClassifyResult {
  // The classify prompt asks for JSON: { category, confidence, reasoning }.
  // Tolerate fenced output (```json ... ```), surrounding prose, and the
  // occasional missing field.
  const cleaned = text
    .replace(/```(?:json)?\s*/gi, '')
    .replace(/```/g, '')
    .trim()
  const start = cleaned.indexOf('{')
  const end = cleaned.lastIndexOf('}')
  if (start === -1 || end === -1 || end <= start) {
    return { category: 'unknown', confidence: 0, reasoning: 'parse_failed: no JSON object found' }
  }
  try {
    const parsed = JSON.parse(cleaned.slice(start, end + 1)) as Partial<ClassifyResult>
    return {
      category: typeof parsed.category === 'string' ? parsed.category : 'unknown',
      confidence: typeof parsed.confidence === 'number' ? parsed.confidence : 0,
      reasoning: typeof parsed.reasoning === 'string' ? parsed.reasoning : '',
    }
  } catch (err) {
    return {
      category: 'unknown',
      confidence: 0,
      reasoning: `parse_failed: ${err instanceof Error ? err.message : String(err)}`,
    }
  }
}

function parseComposeResponse(text: string, description: string): ComposeNewResult {
  // The newMail prompt asks for JSON { subject, htmlBody }. Tolerate fenced
  // output and surrounding prose; fall back to treating the whole reply as
  // the body with a subject derived from the user's description.
  const cleaned = text
    .replace(/```(?:json)?\s*/gi, '')
    .replace(/```/g, '')
    .trim()
  const start = cleaned.indexOf('{')
  const end = cleaned.lastIndexOf('}')
  if (start !== -1 && end > start) {
    try {
      const parsed = JSON.parse(cleaned.slice(start, end + 1)) as Partial<ComposeNewResult>
      const subject = typeof parsed.subject === 'string' ? parsed.subject.trim() : ''
      const htmlBody = typeof parsed.htmlBody === 'string' ? parsed.htmlBody.trim() : ''
      if (subject || htmlBody) {
        return {
          subject: subject || deriveSubject(description),
          htmlBody: htmlBody || `<p>${escapeHtml(text)}</p>`,
        }
      }
    } catch {
      // Fall through to the plain-text fallback below.
    }
  }
  return {
    subject: deriveSubject(description),
    htmlBody: `<p>${escapeHtml(text || description)}</p>`,
  }
}

function deriveSubject(description: string): string {
  return description.split(/\r?\n/)[0]?.slice(0, 80).trim() || 'New message'
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * Extract the AI reply text from `outgoingMessage`. The OpenAPI spec only
 * declares this as `type: object`, so we try the most common field names
 * Synaplan uses in adjacent endpoints (`text`, `content`, `body`,
 * `message`, `answer`). Once we have a real API key we should do a one-
 * shot live smoke and pin the exact field name here.
 *
 * If none of those match but the object is itself a string, return it.
 * If nothing fits, return JSON-stringified shape so the UI still shows
 * SOMETHING the user can copy to a bug report rather than a silent
 * empty string.
 */
function extractAiText(msg: unknown): string {
  if (typeof msg === 'string') return msg
  if (msg && typeof msg === 'object') {
    const obj = msg as Record<string, unknown>
    for (const key of ['text', 'content', 'body', 'message', 'answer', 'reply']) {
      const v = obj[key]
      if (typeof v === 'string' && v.length > 0) return v
    }
  }
  return typeof msg === 'undefined' ? '' : JSON.stringify(msg)
}

function base64ToBlob(b64: string, mime: string): Blob {
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return new Blob([bytes], { type: mime })
}

interface FileUploadResponse {
  files?: { id?: number; file_id?: number }[]
  fileIds?: number[]
  file_id?: number
}

function pickFirstFileId(res: FileUploadResponse | undefined | null): number {
  if (!res) return 0
  if (Array.isArray(res.files) && res.files.length > 0) {
    return res.files[0].id ?? res.files[0].file_id ?? 0
  }
  if (Array.isArray(res.fileIds) && res.fileIds.length > 0) return res.fileIds[0]
  if (typeof res.file_id === 'number') return res.file_id
  return 0
}

// ---------------------------------------------------------------------------
// Wire-format type shims for the response shapes we accept
// ---------------------------------------------------------------------------

interface AuthMeResponse {
  success?: boolean
  user?: { email?: string; id?: number; level?: string; isAdmin?: boolean }
}

interface MessagesSendResponse {
  success?: boolean
  incomingMessage?: unknown
  // `type: object` in the OpenAPI; extracted by extractAiText().
  outgoingMessage?: unknown
}

interface CreateChatResponse {
  success?: boolean
  chat?: { id?: number; title?: string; createdAt?: string; updatedAt?: string }
}

interface RagSearchResponse {
  success?: boolean
  results?: {
    id?: number
    text?: string
    score?: number
    file_id?: number
    file_name?: string
    group_key?: string | null
  }[]
  query?: string
  total_results?: number
}

interface RawRagGroup {
  id?: string
  key?: string
  name?: string
  description?: string
}

interface RagGroupsResponse {
  groups?: RawRagGroup[]
}

interface ConfigDefaultsResponse {
  success?: boolean
  defaults?: Record<string, number | null>
}

interface RawCatalogModel {
  id: number
  name?: string
  service?: string
}

interface ConfigModelsResponse {
  success?: boolean
  models?: Record<string, RawCatalogModel[]>
}
