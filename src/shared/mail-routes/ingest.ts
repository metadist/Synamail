/**
 * Pure planning for the knowledge-base ingestion routes (rules 2 & 3).
 *
 * Decides WHAT to ingest from a matched email and HOW (translate-before-add,
 * tagging, filenames) — without doing any I/O. The runtime composable executes
 * the plan: translating via the Synaplan client and uploading via `fileUpload`.
 * Keeping the decision logic pure makes the "is this worth ingesting?" and the
 * "store original + translation" rules unit-testable.
 *
 * Language policy (the user's requirement): the KB has a working language;
 * off-language content is translated into it before being added, BUT the
 * original is also stored so searches can reveal the source text in its
 * original language. We therefore plan up to two body artifacts: the original
 * and (when translation is needed) a working-language companion linked by
 * metadata.
 */

import type { IngestRouteBase } from './types'

/** Document attachment extensions worth adding to a knowledge base. */
export const INGESTABLE_EXTENSIONS = [
  'pdf',
  'doc',
  'docx',
  'xls',
  'xlsx',
  'ppt',
  'pptx',
  'txt',
  'csv',
  'md',
  'rtf',
  'odt',
  'odp',
  'ods',
] as const

/** True for attachments we want to capture (by content type or extension). */
export function isIngestableAttachment(name: string, contentType?: string): boolean {
  const ct = (contentType ?? '').toLowerCase()
  if (
    ct.includes('pdf') ||
    ct.includes('word') ||
    ct.includes('excel') ||
    ct.includes('spreadsheet') ||
    ct.includes('presentation') ||
    ct.includes('powerpoint') ||
    ct === 'text/plain' ||
    ct === 'text/csv' ||
    ct === 'text/markdown' ||
    ct.includes('opendocument')
  ) {
    return true
  }
  const ext = (name.split('.').pop() ?? '').toLowerCase()
  return (INGESTABLE_EXTENSIONS as readonly string[]).includes(ext)
}

export interface IngestAttachmentMeta {
  name: string
  contentType?: string
  /** Skip tiny inline images / signature logos at the runtime layer. */
  size?: number
  isInline?: boolean
}

export interface IngestSource {
  /** Dedup key for "seen" tracking; absent on hosts that don't expose it. */
  internetMessageId?: string
  subject: string
  from?: string
  bodyText: string
  attachments: IngestAttachmentMeta[]
}

export type IngestArtifactRole = 'body-original' | 'body-translation' | 'attachment'

export interface IngestArtifactPlan {
  role: IngestArtifactRole
  /** Suggested filename for the uploaded artifact. */
  filename: string
  /** For `attachment` artifacts: which source attachment to fetch. */
  attachmentName?: string
}

export interface IngestPlan {
  groupId: string
  tag: string
  workingLanguage: string
  /** Whether the body text should be translated into `workingLanguage`. */
  translateBody: boolean
  /** Ordered artifacts to create. Empty ⇒ nothing qualified; skip the mail. */
  artifacts: IngestArtifactPlan[]
  /** Metadata stamped on every uploaded artifact. */
  metadata: Record<string, string>
}

function sanitizeBase(subject: string): string {
  const cleaned = subject.replace(/[^a-zA-Z0-9-_]+/g, '_').replace(/^_+|_+$/g, '')
  return cleaned.slice(0, 80) || 'message'
}

function normalizeLang(code: string): string {
  return code.trim().toLowerCase().split(/[-_]/)[0] ?? ''
}

export interface PlanIngestOptions {
  /**
   * Detected source language code, when already known. When it equals the
   * route's working language, no translation companion is planned. When
   * omitted, translation is planned (the translate step is a safe no-op if the
   * AI finds the text already in the working language).
   */
  detectedLanguage?: string
}

/**
 * Build the ingestion plan for one matched email under one ingest route.
 * Returns a plan with an empty `artifacts` list when nothing qualifies (body
 * too short and no ingestable attachments) — callers should skip the mail.
 */
export function planIngest(
  source: IngestSource,
  route: IngestRouteBase,
  opts: PlanIngestOptions = {},
): IngestPlan {
  const base = sanitizeBase(source.subject)
  const workingLanguage = route.workingLanguage
  const bodyQualifies = source.bodyText.trim().length >= route.minBodyChars

  const sameLang =
    opts.detectedLanguage !== undefined &&
    normalizeLang(opts.detectedLanguage) === normalizeLang(workingLanguage)
  const translateBody = bodyQualifies && !sameLang

  const artifacts: IngestArtifactPlan[] = []
  if (bodyQualifies) {
    artifacts.push({ role: 'body-original', filename: `${base}.txt` })
    if (translateBody) {
      artifacts.push({
        role: 'body-translation',
        filename: `${base}.${normalizeLang(workingLanguage)}.txt`,
      })
    }
  }

  if (route.includeAttachments) {
    for (const att of source.attachments) {
      if (att.isInline) continue
      if (!isIngestableAttachment(att.name, att.contentType)) continue
      artifacts.push({ role: 'attachment', filename: att.name, attachmentName: att.name })
    }
  }

  const metadata: Record<string, string> = {
    source: 'synamail-route',
    routeKind: route.kind,
    tag: route.tag,
    workingLanguage,
  }
  if (source.from) metadata.from = source.from
  if (source.subject) metadata.subject = source.subject
  if (source.internetMessageId) metadata.messageId = source.internetMessageId

  return {
    groupId: route.groupId,
    tag: route.tag,
    workingLanguage,
    translateBody,
    artifacts,
    metadata,
  }
}
