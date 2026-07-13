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

/**
 * The writing styles offered by the Email-writing box. `concise` (knapp),
 * `detailed` (ausführlich) and `formal` (formell) are the three surfaced
 * buttons; `friendly` is retained for the legacy tone picker.
 */
export type EmailTone = 'formal' | 'concise' | 'friendly' | 'detailed'

export interface SummariseInput {
  subject: string
  body: string
  from?: string
  to?: string[]
  /** Target language code for the summary (e.g. "de"). Defaults to English. */
  language?: string
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
  tone: EmailTone
  language: string
}

export interface DraftReplyResult {
  htmlBody: string
}

export interface ComposeDraftInput {
  /** One-line intent, e.g. "invite Alice to lunch on Friday and ask her to confirm". */
  intent: string
  tone: EmailTone
  language: string
  /** Body of the message being replied to/forwarded, when composing a reply. */
  referenceBody?: string
}

export interface ComposeDraftResult {
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
  /**
   * Synaplan file ids to attach to this turn (already uploaded via
   * `fileUpload`). Used to give the AI vision over email images/screenshots —
   * each file's extracted text/vision description is injected as context by
   * `POST /messages/send`.
   */
  fileIds?: number[]
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

export interface MeetingExtractInput {
  subject: string
  body: string
  from?: string
  /** Current local time as ISO 8601, so the AI can resolve relative dates. */
  nowIso: string
  /** IANA timezone of the user (e.g. "Europe/Berlin") for disambiguation. */
  timezone: string
}

export interface MeetingProposal {
  title: string
  /** ISO 8601 local datetime, no offset (e.g. "2026-06-03T15:00:00"). */
  startIso: string
  /** ISO 8601 local datetime, no offset. */
  endIso: string
  location?: string
}

// ---------------------------------------------------------------------------
// Contact AI Profiling (docs/CONTACT_PROFILING.md). The rolling profile is
// computed and stored server-side by the `synamail` Synaplan plugin; the
// add-in only feeds emails in and renders the snapshot.
// ---------------------------------------------------------------------------

/**
 * The rolling profile of one mailing partner, as returned by the synamail
 * plugin (`GET/POST /api/v1/user/{userId}/plugins/synamail/profiles/...`).
 */
export interface ContactProfileData {
  /** Lower-cased contact email — the profile key. */
  email: string
  name?: string
  /** Organisation derived from the email domain (null for freemail). */
  org?: string | null
  /** The rolling narrative — who this person is and where things stand. */
  summary: string
  /** Current tone of the relationship, e.g. "friendly but distanced". */
  tone?: string | null
  /** Short, stable facts worth remembering. */
  facts: string[]
  /** Unresolved commitments/questions, each prefixed "me:" or "them:". */
  openLoops: string[]
  /** How many emails have been rolled into this profile. */
  emailCount: number
  firstSeen?: string | null
  lastInbound?: string
  lastOutbound?: string
  /** ISO timestamp of the last profile update ("as of" line in the UI). */
  updatedAt?: string
}

/** One email to roll into a contact's profile. */
export interface ProfileEmailInput {
  /** The contact the profile belongs to. */
  email: string
  subject?: string
  body: string
  /** ISO 8601 date of the email. */
  date?: string
  /** inbound = the contact wrote to me; outbound = I wrote to the contact. */
  direction: 'inbound' | 'outbound'
  /** Display name of the contact, if known. */
  name?: string
}

export interface ApiError {
  status: number
  code: string
  message: string
}
