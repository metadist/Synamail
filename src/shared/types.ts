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
  /** Per-user language override. */
  language?: 'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'
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
}

export interface ChatTurnResult {
  chatId: number
  answer: string
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
  /** RAG group id to add the file to. */
  groupId?: string
  /** Free-form metadata; we tag email-sourced files with from/to addresses. */
  metadata?: Record<string, string>
}

export interface FileUploadResult {
  fileId: number
}

export interface RagGroup {
  id: string
  name: string
  description?: string
}

export interface ApiError {
  status: number
  code: string
  message: string
}
