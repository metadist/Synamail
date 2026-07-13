/**
 * Synaplan API client.
 *
 * The interface (`SynaplanClient`) is shaped around the taskpane's
 * read-mode / compose-mode features. There is a single implementation,
 * `RealSynaplanClient`, which hits whatever Synaplan instance the user
 * signed in to — the local bridge (`https://localhost:5174`) or a remote
 * server (`https://web.synaplan.com`, or a self-hosted host). There is no
 * mock/offline mode.
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
  compose as composePrompt,
  meetingProposals as meetingProposalsPrompt,
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
  ComposeDraftInput,
  ComposeDraftResult,
  ContactProfileData,
  DraftReplyInput,
  DraftReplyResult,
  FileUploadInput,
  FileUploadResult,
  MeetingExtractInput,
  MeetingProposal,
  ModelConfig,
  ProfileEmailInput,
  RagGroup,
  RagSearchHit,
  RagSearchInput,
  SummariseInput,
  SummariseResult,
  TranslateInput,
  TranslateResult,
} from './types'

/**
 * Called with the cumulative text so far each time the server streams another
 * chunk. Passing it to a streaming-capable method switches that call from a
 * single blocking round-trip to a live Server-Sent-Events stream.
 */
export type StreamHandler = (textSoFar: string) => void

export interface SynaplanClient {
  ping(): Promise<{ ok: boolean; email?: string }>
  summarise(input: SummariseInput, onChunk?: StreamHandler): Promise<SummariseResult>
  translate(input: TranslateInput, onChunk?: StreamHandler): Promise<TranslateResult>
  draftReply(input: DraftReplyInput): Promise<DraftReplyResult>
  /**
   * Compose-mode "Draft from prompt": turn a short user intent into a full
   * email body. Returns HTML suitable for `body.setAsync` in the open draft.
   */
  composeDraft(input: ComposeDraftInput): Promise<ComposeDraftResult>
  classify(input: ClassifyInput): Promise<ClassifyResult>
  ask(input: ChatTurnInput, onChunk?: StreamHandler): Promise<ChatTurnResult>
  /**
   * General-purpose chat, not tied to a specific email. Mirrors `ask` for
   * chat-id persistence (caller stores the returned `chatId` in roaming
   * under a stable key such as `home`), but uses the simple-chat system
   * prompt instead of the email-grounded one.
   */
  chat(input: ChatTurnInput, onChunk?: StreamHandler): Promise<ChatTurnResult>
  /**
   * Extract proposed meeting / call times from an email body. Returns a
   * (possibly empty) list of candidate slots the caller can turn into Outlook
   * appointments. Pure AI extraction via `POST /messages/send`.
   */
  extractMeetingTimes(input: MeetingExtractInput): Promise<MeetingProposal[]>
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

  // ---------------------------------------------------------------------
  // Contact AI Profiling — served by the `synamail` Synaplan plugin
  // (docs/CONTACT_PROFILING.md). When the plugin isn't installed on the
  // user's Synaplan instance these calls throw an ApiError with code
  // `PROFILING_UNAVAILABLE`, which the UI renders as an install hint.
  // ---------------------------------------------------------------------

  /** The stored rolling profile of a contact, or null when none exists yet. */
  getContactProfile(email: string): Promise<ContactProfileData | null>
  /** Roll one email into the contact's profile and return the updated snapshot. */
  updateContactProfile(input: ProfileEmailInput): Promise<ContactProfileData>
  /** Delete the contact's profile entirely (privacy / GDPR). */
  deleteContactProfile(email: string): Promise<void>
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
  /** Synaplan user id, resolved lazily from /auth/me for plugin routes. */
  private userId: number | null = null

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

  async summarise(input: SummariseInput, onChunk?: StreamHandler): Promise<SummariseResult> {
    const lang = input.language ?? 'en'
    const message = composeMessage(summarisePrompt(lang), buildEmailBlock(input))
    const text = await this.runChat(message, { onChunk, chatTitle: 'Outlook: summarise' })
    return {
      summary: text,
      bullets: parseMarkdownBullets(text),
      language: lang,
    }
  }

  async translate(input: TranslateInput, onChunk?: StreamHandler): Promise<TranslateResult> {
    const message = composeMessage(
      translatePrompt(input.targetLanguage),
      cleanEmailText(input.text),
    )
    const text = await this.runChat(message, { onChunk, chatTitle: 'Outlook: translate' })
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

  async composeDraft(input: ComposeDraftInput): Promise<ComposeDraftResult> {
    const parts = ['[intent]', input.intent.trim()]
    if (input.referenceBody) {
      parts.push('', '[replying to]', cleanEmailText(input.referenceBody))
    }
    const message = composeMessage(composePrompt(input.tone, input.language), parts.join('\n'))
    const text = await this.sendChat(message)
    return { htmlBody: text }
  }

  async classify(input: ClassifyInput): Promise<ClassifyResult> {
    const categories = ['billing', 'support', 'internal', 'personal', 'spam', 'general']
    const message = composeMessage(classifyPrompt(categories), buildEmailBlock(input))
    const text = await this.sendChat(message)
    return parseClassifyResponse(text)
  }

  async ask(input: ChatTurnInput, onChunk?: StreamHandler): Promise<ChatTurnResult> {
    // 1. Ensure a chat exists for this conversation. If the caller didn't
    //    supply chatId from roaming, create a fresh chat now and return its
    //    id alongside the answer so the caller can persist it.
    const chatId =
      input.chatId ?? (await this.ensureChat(`Outlook: ${input.conversationId.slice(0, 40)}`))
    // 2. Send the question into that chat. Email context is cleaned so tracking
    //    URLs / autolink soup don't pollute the grounding (or the reply).
    const prefix = input.emailContext
      ? `${askPrompt()}\n\n[email context]\n${cleanEmailText(input.emailContext)}\n\n[question]\n`
      : `${askPrompt()}\n\n`
    const answer = await this.runChat(`${prefix}${input.question}`, {
      onChunk,
      chatId,
      fileIds: input.fileIds,
    })
    return { chatId, answer }
  }

  async chat(input: ChatTurnInput, onChunk?: StreamHandler): Promise<ChatTurnResult> {
    const chatId = input.chatId ?? (await this.ensureChat('Synamail chat'))
    const answer = await this.runChat(`${simpleChatPrompt()}\n\n${input.question}`, {
      onChunk,
      chatId,
      fileIds: input.fileIds,
    })
    return { chatId, answer }
  }

  async extractMeetingTimes(input: MeetingExtractInput): Promise<MeetingProposal[]> {
    const message = composeMessage(
      meetingProposalsPrompt(input.nowIso, input.timezone),
      buildEmailBlock({ subject: input.subject, body: input.body, from: input.from }),
    )
    const text = await this.sendChat(message)
    return parseMeetingProposals(text)
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
    // Live API (verified 2026-06-04) returns hits as
    // `{ chunk_id, message_id, text, score, start_line, end_line }` — NOT the
    // `{ file_id, file_name, group_key }` the OpenAPI annotation advertises.
    // Map from the real fields; `message_id` is the source-message id, the hit
    // carries no filename, and the group is the one we searched in.
    const searchedGroup = input.groups && input.groups.length > 0 ? input.groups[0] : undefined
    return (res?.results ?? []).map((hit) => ({
      fileId: hit.message_id ?? hit.file_id ?? 0,
      filename: hit.file_name ?? '',
      snippet: hit.text ?? '',
      score: hit.score ?? 0,
      group: hit.group_key ?? searchedGroup,
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
    // Synaplan returns HTTP 200/206 with `success:false` + an `errors` array
    // when a file is rejected (e.g. unsupported type). That's NOT a successful
    // save, so surface the real reason instead of a bogus "file #0".
    if (!fileId) {
      throw apiError(0, 'UPLOAD_FAILED', uploadErrorMessage(res))
    }
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
  // Contact AI Profiling — `synamail` plugin endpoints
  // -------------------------------------------------------------------------

  async getContactProfile(email: string): Promise<ContactProfileData | null> {
    const res = await this.pluginRequest<ProfileResponse>(
      `/profiles/${encodeURIComponent(email.trim().toLowerCase())}`,
    )
    return res?.profile ?? null
  }

  async updateContactProfile(input: ProfileEmailInput): Promise<ContactProfileData> {
    const res = await this.pluginRequest<ProfileResponse>(
      `/profiles/${encodeURIComponent(input.email.trim().toLowerCase())}/update`,
      {
        method: 'POST',
        body: JSON.stringify({
          subject: input.subject ?? '',
          body: cleanEmailText(input.body),
          date: input.date ?? '',
          direction: input.direction,
          name: input.name ?? '',
        }),
      },
    )
    if (!res?.success || !res.profile) {
      throw apiError(0, 'AI_FAILED', 'Profile update failed')
    }
    return res.profile
  }

  async deleteContactProfile(email: string): Promise<void> {
    await this.pluginRequest(`/profiles/${encodeURIComponent(email.trim().toLowerCase())}`, {
      method: 'DELETE',
    })
  }

  /**
   * Request against the synamail plugin's user-scoped route prefix. Resolves
   * (and caches) the numeric user id from /auth/me first, and maps a 404 —
   * the plugin not being installed on this Synaplan instance — to the
   * `PROFILING_UNAVAILABLE` error code the UI knows how to explain.
   */
  private async pluginRequest<T>(path: string, init: RequestInit = {}): Promise<T> {
    if (this.userId === null) {
      const me = await this.request<AuthMeResponse>('/api/v1/auth/me')
      const id = me?.user?.id
      if (typeof id !== 'number') {
        throw apiError(0, 'AUTH_FAILED', 'Could not resolve the Synaplan user id')
      }
      this.userId = id
    }
    try {
      return await this.request<T>(`/api/v1/user/${this.userId}/plugins/synamail${path}`, init)
    } catch (err) {
      if (isApiError(err) && (err.status === 404 || err.status === 403)) {
        throw apiError(
          err.status,
          'PROFILING_UNAVAILABLE',
          'The Synamail plugin is not installed on this Synaplan instance.',
        )
      }
      throw err
    }
  }

  // -------------------------------------------------------------------------
  // Internal: AI chat round-trip
  // -------------------------------------------------------------------------

  private async sendChat(message: string, trackId?: number, fileIds?: number[]): Promise<string> {
    const body: Record<string, unknown> = { message }
    if (trackId !== undefined) body.trackId = trackId
    if (fileIds && fileIds.length > 0) body.fileIds = fileIds
    const res = await this.request<MessagesSendResponse>('/api/v1/messages/send', {
      method: 'POST',
      body: JSON.stringify(body),
    })
    if (!res?.success) {
      throw apiError(0, 'AI_FAILED', 'Synaplan reported a non-success response')
    }
    return extractAiText(res.outgoingMessage)
  }

  /** Create a chat and return its id (used when the caller has none yet). */
  private async ensureChat(title: string): Promise<number> {
    const created = await this.request<CreateChatResponse>('/api/v1/chats', {
      method: 'POST',
      body: JSON.stringify({ title }),
    })
    if (!created?.success || !created.chat?.id) {
      throw apiError(0, 'CHAT_CREATE_FAILED', 'Could not create chat')
    }
    return created.chat.id
  }

  /**
   * Run one AI turn. With `onChunk` it streams via SSE (the answer renders as it
   * arrives); without it, it falls back to the single blocking round-trip. The
   * stream endpoint requires a `chatId`, so a throwaway chat is created on
   * demand for the stateless actions (translate/summarise) that don't carry one.
   */
  private async runChat(
    message: string,
    opts: { onChunk?: StreamHandler; chatId?: number; chatTitle?: string; fileIds?: number[] },
  ): Promise<string> {
    if (!opts.onChunk) {
      return this.sendChat(message, opts.chatId, opts.fileIds)
    }
    const chatId = opts.chatId ?? (await this.ensureChat(opts.chatTitle ?? 'Synamail chat'))
    try {
      return await this.streamChat(message, opts.onChunk, { chatId, fileIds: opts.fileIds })
    } catch (err) {
      if (isApiError(err) && err.status === 401) throw err
      // Streaming unavailable on this deployment (or it failed): fall back to
      // the blocking endpoint so the action still completes. The caller sets
      // the final text from the return value, so the UI is correct either way.
      return this.sendChat(message, chatId, opts.fileIds)
    }
  }

  /**
   * Stream an AI reply over Server-Sent Events from `GET /messages/stream`.
   * The server emits newline-delimited `data: {json}` frames whose `status`
   * field tags the event (`data` carries an incremental `chunk`; `complete`
   * ends the turn; `error` aborts it). Chunks are deltas, so we accumulate
   * them and hand the running total to `onChunk` after each one.
   */
  private async streamChat(
    message: string,
    onChunk: StreamHandler,
    opts: { chatId: number; fileIds?: number[] },
  ): Promise<string> {
    const params = new URLSearchParams({ message, chatId: String(opts.chatId) })
    if (opts.fileIds && opts.fileIds.length > 0) params.set('fileIds', opts.fileIds.join(','))

    const headers = new Headers()
    headers.set('X-API-Key', this.apiKey)
    headers.set('Accept', 'text/event-stream')

    const res = await this.fetchImpl(`${this.baseUrl}/api/v1/messages/stream?${params}`, {
      method: 'GET',
      headers,
    })
    if (res.status === 401) throw apiError(401, 'AUTH_FAILED', 'API key rejected')
    if (!res.ok || !res.body) {
      const text = await safeText(res)
      throw apiError(res.status, 'HTTP_ERROR', text || res.statusText)
    }

    const reader = res.body.getReader()
    const decoder = new TextDecoder()
    let buffer = ''
    let full = ''
    let streamError: string | null = null

    const handleFrame = (frame: string): void => {
      // A frame is one or more `data:` lines; join their payloads and parse.
      const data = frame
        .split('\n')
        .filter((l) => l.startsWith('data:'))
        .map((l) => l.slice(5).trim())
        .join('\n')
      if (!data) return
      let evt: SseEvent
      try {
        evt = JSON.parse(data) as SseEvent
      } catch {
        return
      }
      if (evt.status === 'data' && typeof evt.chunk === 'string') {
        full += evt.chunk
        onChunk(full)
      } else if (evt.status === 'error') {
        streamError = evt.error || evt.message || 'Streaming failed'
      }
    }

    for (;;) {
      const { value, done } = await reader.read()
      if (done) break
      buffer += decoder.decode(value, { stream: true })
      let sep: number
      while ((sep = buffer.indexOf('\n\n')) !== -1) {
        handleFrame(buffer.slice(0, sep))
        buffer = buffer.slice(sep + 2)
      }
    }
    if (buffer.trim()) handleFrame(buffer)

    if (streamError) throw apiError(0, 'AI_FAILED', streamError)
    return full
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

export interface CreateClientOptions {
  baseUrl: string
  apiKey: string
}

export function createSynaplanClient(opts: CreateClientOptions): SynaplanClient {
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
  parts.push('', cleanEmailText(input.body))
  return parts.join('\n')
}

/**
 * Strip the noise that wrecks AI output on real-world (esp. newsletter) email:
 * tracking URLs and angle-bracket autolinks (`<https://…>`), and runs of blank
 * lines. Short, human-meaningful URLs are left intact. Keeping this junk out of
 * the prompt is what stops translations/summaries from echoing link soup back.
 */
export function cleanEmailText(s: string): string {
  if (!s) return ''
  return (
    s
      .replace(/\r\n/g, '\n')
      // Long tracking URLs, with or without surrounding < > (autolink) brackets.
      // The char class excludes < > and whitespace so it stops at the autolink's
      // closing bracket and never eats adjacent text that follows it.
      .replace(/<?\b(?:https?:\/\/|mailto:)[^\s<>]{30,}>?/gi, '')
      // Leftover empty autolink brackets.
      .replace(/<\s*>/g, '')
      // Trailing spaces and 3+ blank lines collapse to a single blank line.
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim()
  )
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

function parseMeetingProposals(text: string): MeetingProposal[] {
  // The meetingProposals prompt asks for a JSON array. Tolerate fenced output
  // and surrounding prose; return [] on anything unparseable rather than throw
  // (no proposals is a normal, expected outcome).
  const cleaned = text
    .replace(/```(?:json)?\s*/gi, '')
    .replace(/```/g, '')
    .trim()
  const start = cleaned.indexOf('[')
  const end = cleaned.lastIndexOf(']')
  if (start === -1 || end === -1 || end <= start) return []
  let parsed: unknown
  try {
    parsed = JSON.parse(cleaned.slice(start, end + 1))
  } catch {
    return []
  }
  if (!Array.isArray(parsed)) return []
  const out: MeetingProposal[] = []
  for (const raw of parsed) {
    if (!raw || typeof raw !== 'object') continue
    const o = raw as Record<string, unknown>
    const startIso = typeof o.start === 'string' ? o.start : ''
    if (!startIso) continue
    const endIso = typeof o.end === 'string' && o.end ? o.end : addMinutesIso(startIso, 30)
    out.push({
      title: typeof o.title === 'string' && o.title ? o.title : 'Meeting',
      startIso,
      endIso,
      location: typeof o.location === 'string' && o.location ? o.location : undefined,
    })
  }
  return out
}

/** Add minutes to a local ISO datetime, returning a local ISO datetime. */
function addMinutesIso(iso: string, minutes: number): string {
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return iso
  d.setMinutes(d.getMinutes() + minutes)
  // Re-emit local wall-clock without the trailing Z/offset.
  const pad = (n: number): string => String(n).padStart(2, '0')
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
    `T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
  )
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
  success?: boolean
  files?: { id?: number; file_id?: number; success?: boolean; error?: string }[]
  fileIds?: number[]
  file_id?: number
  errors?: { filename?: string; error?: string }[]
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

/** Best human-readable reason a file upload didn't yield a stored file id. */
function uploadErrorMessage(res: FileUploadResponse | undefined | null): string {
  const fromErrors = res?.errors?.find((e) => e.error)?.error
  if (fromErrors) return fromErrors
  const fromFile = res?.files?.find((f) => f.error)?.error
  if (fromFile) return fromFile
  return 'Upload failed — the server did not store the file.'
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

/**
 * One decoded `data:` frame from `GET /messages/stream`. The server tags every
 * frame with `status`; we act on `data` (incremental `chunk`) and `error`, and
 * ignore the rest (`status`, `thinking`, `complete`, `file`, `perf`, …).
 */
interface SseEvent {
  status?: string
  chunk?: string
  error?: string
  message?: string
}

interface CreateChatResponse {
  success?: boolean
  chat?: { id?: number; title?: string; createdAt?: string; updatedAt?: string }
}

interface RagSearchResponse {
  success?: boolean
  results?: {
    // Real fields returned by the live API (verified 2026-06-04).
    chunk_id?: string
    message_id?: number
    text?: string
    score?: number
    start_line?: number | null
    end_line?: number | null
    // Legacy/spec-advertised fields kept for forward-compat — not currently
    // emitted by the server but tolerated if a future version adds them.
    id?: number
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

/** Wire shape of the synamail plugin's profile endpoints. */
interface ProfileResponse {
  success?: boolean
  profile?: ContactProfileData | null
}
