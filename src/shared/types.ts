/**
 * Shared types used by both the taskpane and the dialog.
 *
 * Kept narrow on purpose — full Zod-validated schemas land in Sprint 3 when
 * we generate them from Synaplan's OpenAPI spec (see docs/FEATURES.md §6).
 */

export interface SignInPayload {
  state: string
  apiKey: string
  keyId: number
  email: string
  baseUrl: string
}

export interface RoamingSettings {
  apiKey: string
  keyId: number
  email: string
  baseUrl: string
  /** Per-Outlook-conversation chat id mapping for the Ask feature. */
  chats?: Record<string, number>
  /** Last-used RAG group id, to pre-select in the group picker. */
  lastRagGroupId?: string
  /**
   * Per-user UI language override (taskpane chrome). `'auto'` follows the
   * Outlook display language; any other value must be a shipped UI locale
   * (see `SUPPORTED_LOCALES` in `src/i18n.ts`). This is NOT the translate
   * target language — that lives in each view's local `targetLang`.
   */
  language?: 'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'pt'
}

export interface SummariseInput {
  subject: string
  body: string
  from?: string
  to?: string[]
}

export interface SummariseResult {
  summary: string
  bullets: string[]
  language: string
}

export interface DraftReplyInput {
  subject: string
  body: string
  threadContext?: string[]
  tone: 'formal' | 'concise' | 'friendly'
  language: string
}

export interface DraftReplyResult {
  htmlBody: string
}

export interface TranslateInput {
  text: string
  targetLanguage: string
}

export interface TranslateResult {
  translation: string
  detectedLanguage: string
}

export interface ClassifyInput {
  subject: string
  body: string
  from?: string
}

export interface ClassifyResult {
  category: string
  confidence: number
  reasoning: string
}

export interface ChatTurnInput {
  conversationId: string
  question: string
  /** Optional context: original email body. */
  emailContext?: string
  /**
   * Existing chat id for this conversation (from roaming.chats[conversationId]).
   * When set, the client reuses the chat via `trackId`; when undefined, the
   * client creates a new chat first and the returned `chatId` should be
   * persisted by the caller for subsequent turns.
   */
  chatId?: number
}

export interface ChatTurnResult {
  chatId: number
  answer: string
}

export interface ComposeNewInput {
  /** Free-form description of the email the user wants written. */
  description: string
  /** Target language code; defaults to English. */
  language?: string
}

export interface ComposeNewResult {
  subject: string
  /** HTML body content (no doctype/html/head wrapper). */
  htmlBody: string
}

export interface RagSearchInput {
  query: string
  threshold?: number
  limit?: number
  groups?: string[]
}

export interface RagSearchHit {
  fileId: number
  filename: string
  snippet: string
  score: number
  group?: string
}

export interface FileUploadInput {
  filename: string
  contentBase64: string
  mimeType: string
  /** RAG group id to add the file to (becomes `group_key` in the upload). */
  groupId?: string
  /** Free-form metadata; we tag email-sourced files with from/to addresses. */
  metadata?: Record<string, string>
  /**
   * How aggressively Synaplan processes the file after upload. Combines what
   * used to be a separate POST /files/{id}/process call into the same
   * request. Defaults to `extract` (safe minimum); use `vectorize` for RAG
   * ingestion and `full` for full extraction + vectorisation + analysis.
   */
  processLevel?: 'store' | 'extract' | 'vectorize' | 'full'
}

export interface FileUploadResult {
  fileId: number
}

export interface RagGroup {
  id: string
  name: string
  description?: string
}

export interface ModelChoice {
  id: number
  name: string
  /** Provider/service name (e.g. "Groq", "OpenAI"). */
  service?: string
}

/**
 * The user's currently-configured Synaplan models for the three capabilities
 * Synamail surfaces. `null` means no model is set for that capability.
 */
export interface ModelConfig {
  chat: ModelChoice | null
  imageGen: ModelChoice | null
  vectorize: ModelChoice | null
}

export interface SenderHistoryInput {
  email: string
  /** Hard cap on results (default 25). */
  limit?: number
  /** Optional folder hint; defaults to all mailbox folders. */
  folder?: 'inbox' | 'all' | 'archive'
}

export interface SenderHistoryItem {
  /** ISO 8601 received date. */
  date: string
  subject: string
  /** Short plain-text preview (~120 chars). */
  snippet: string
  /** Whether the message has been read. Drives the bold/regular treatment. */
  unread: boolean
  /** Stable id usable to deep-link back to the message (EWS id when available). */
  messageId?: string
}

export interface SenderHistoryResult {
  email: string
  total: number
  items: SenderHistoryItem[]
  /** True when the source was Outlook itself (EWS / REST); false when the
   *  client returned canned mock data — drives the "(mock)" UI badge. */
  fromOutlook: boolean
}

export interface CreateSpamRuleInput {
  senderEmail: string
  /** Also move existing messages from this sender to Junk. */
  alsoCleanExisting?: boolean
  /** Display name for the rule (defaults to "Synamail: block <sender>"). */
  displayName?: string
}

export interface CreateSpamRuleResult {
  ruleId: string
  /** Number of existing messages moved to Junk (0 when alsoCleanExisting is false). */
  movedCount: number
  /** When true the rule was created server-side via EWS / Graph; false when
   *  the mock client stubbed it for dev. */
  serverSide: boolean
}

export interface ApiError {
  status: number
  code: string
  message: string
}
