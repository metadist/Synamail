import { describe, expect, it } from 'vitest'
import { isIngestableAttachment, planIngest, type IngestSource } from '@shared/mail-routes/ingest'
import type { ProjectIngestRoute } from '@shared/mail-routes/types'

const route: ProjectIngestRoute = {
  id: 'p1',
  kind: 'projectIngest',
  name: 'Project Helios',
  enabled: true,
  senders: ['@partner.com'],
  groupId: 'project:helios',
  tag: 'helios',
  workingLanguage: 'en',
  minBodyChars: 200,
  includeAttachments: true,
}

function source(overrides: Partial<IngestSource> = {}): IngestSource {
  return {
    internetMessageId: '<abc@partner.com>',
    subject: 'Spec update',
    from: 'lead@partner.com',
    bodyText: 'x'.repeat(300),
    attachments: [],
    ...overrides,
  }
}

describe('isIngestableAttachment', () => {
  it('accepts documents by extension', () => {
    expect(isIngestableAttachment('spec.pdf')).toBe(true)
    expect(isIngestableAttachment('budget.xlsx')).toBe(true)
    expect(isIngestableAttachment('notes.md')).toBe(true)
  })

  it('accepts by content type when the name lacks an extension', () => {
    expect(isIngestableAttachment('blob', 'application/pdf')).toBe(true)
  })

  it('rejects images and unknown types', () => {
    expect(isIngestableAttachment('logo.png', 'image/png')).toBe(false)
    expect(isIngestableAttachment('thing.bin')).toBe(false)
  })
})

describe('planIngest', () => {
  it('plans an original + translation body artifact when language is unknown', () => {
    const plan = planIngest(source(), route)
    expect(plan.translateBody).toBe(true)
    expect(plan.artifacts.map((a) => a.role)).toEqual(['body-original', 'body-translation'])
    expect(plan.artifacts[1].filename).toMatch(/\.en\.txt$/)
    expect(plan.groupId).toBe('project:helios')
    expect(plan.metadata.tag).toBe('helios')
    expect(plan.metadata.messageId).toBe('<abc@partner.com>')
  })

  it('skips translation when the detected language matches the working language', () => {
    const plan = planIngest(source(), route, { detectedLanguage: 'en-US' })
    expect(plan.translateBody).toBe(false)
    expect(plan.artifacts.map((a) => a.role)).toEqual(['body-original'])
  })

  it('drops the body when it is below the minimum length', () => {
    const plan = planIngest(source({ bodyText: 'too short' }), route)
    expect(plan.artifacts.filter((a) => a.role.startsWith('body'))).toHaveLength(0)
  })

  it('includes ingestable, non-inline attachments and skips others', () => {
    const plan = planIngest(
      source({
        bodyText: 'short',
        attachments: [
          { name: 'spec.pdf', contentType: 'application/pdf' },
          { name: 'logo.png', contentType: 'image/png', isInline: true },
          { name: 'random.bin' },
        ],
      }),
      route,
    )
    expect(plan.artifacts.map((a) => a.attachmentName)).toEqual(['spec.pdf'])
  })

  it('honours includeAttachments=false', () => {
    const plan = planIngest(source({ bodyText: 'short', attachments: [{ name: 'spec.pdf' }] }), {
      ...route,
      includeAttachments: false,
    })
    expect(plan.artifacts).toHaveLength(0)
  })
})
