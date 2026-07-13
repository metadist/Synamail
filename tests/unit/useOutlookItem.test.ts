import { describe, expect, it, vi } from 'vitest'
import {
  buildEmailText,
  countKnowledgeBaseAttachments,
  getDocumentAttachments,
  getKnowledgeBaseAttachments,
  getReadItemAsFile,
  setComposeBody,
} from '@/taskpane/composables/useOutlookItem'
import type { OutlookItemSnapshot } from '@/taskpane/composables/useOutlookItem'

function setOfficeItem(item: unknown): void {
  const office = (globalThis as unknown as { Office: { context: { mailbox: { item: unknown } } } })
    .Office
  office.context.mailbox.item = item
}

function snapshot(overrides: Partial<OutlookItemSnapshot> = {}): OutlookItemSnapshot {
  return {
    mode: 'read',
    subject: 'Q3 plan',
    from: 'alice@example.test',
    to: ['team@example.test'],
    cc: [],
    bodyText: 'Hello world',
    attachments: [],
    ...overrides,
  }
}

describe('getReadItemAsFile', () => {
  it('exports the message as a .txt with a header block + body (never .eml)', () => {
    const file = getReadItemAsFile(snapshot({ subject: 'Q3 plan' }))
    expect(file.filename).toBe('Q3_plan.txt')
    expect(file.mimeType).toBe('text/plain')
    const text = atob(file.contentBase64)
    expect(text).toContain('Subject: Q3 plan')
    expect(text).toContain('From: alice@example.test')
    expect(text).toContain('To: team@example.test')
    expect(text).toContain('Hello world')
  })

  it('uses body-only and message.txt when there are no headers', () => {
    const file = getReadItemAsFile(
      snapshot({ subject: '', from: undefined, to: [], cc: [], bodyText: 'hello' }),
    )
    expect(file.filename).toBe('message.txt')
    expect(file.mimeType).toBe('text/plain')
    expect(atob(file.contentBase64)).toBe('hello')
  })

  it('sanitises the subject when building the filename', () => {
    const file = getReadItemAsFile(snapshot({ subject: 'Hello / World? *!' }))
    expect(file.filename).toBe('Hello_World.txt')
  })
})

describe('buildEmailText', () => {
  it('puts headers before a blank line and the body', () => {
    const text = buildEmailText(snapshot({ subject: 'Hi', from: 'a@b.test', to: [], cc: [] }))
    expect(text).toBe('Subject: Hi\nFrom: a@b.test\n\nHello world')
  })

  it('cleans noise (tracking URLs, quoted tail) out of the body before export', () => {
    const noisy = [
      'Please review the plan.',
      '',
      '-----Original Message-----',
      'From: bob@example.test',
      'lots of quoted history',
    ].join('\n')
    const text = buildEmailText(
      snapshot({ subject: 'Plan', from: 'a@b.test', to: [], bodyText: noisy }),
    )
    expect(text).toContain('Please review the plan.')
    expect(text).not.toContain('Original Message')
    expect(text).not.toContain('quoted history')
  })
})

interface TestAttachment {
  id: string
  name: string
  contentType?: string
  size?: number
  isInline?: boolean
  attachmentType?: string
}

function attachment(over: TestAttachment): Office.AttachmentDetails {
  return {
    contentType: '',
    size: 100_000,
    isInline: false,
    attachmentType: 'file',
    ...over,
  } as unknown as Office.AttachmentDetails
}

/** Stub an Outlook item that hands back base64 bytes for any attachment id. */
function setItemWithAttachmentBytes(): void {
  setOfficeItem({
    getAttachmentContentAsync: (
      id: string,
      cb: (r: { status: string; value: { format: string; content: string } }) => void,
    ) => cb({ status: 'succeeded', value: { format: 'base64', content: `b64:${id}` } }),
  })
}

describe('getDocumentAttachments', () => {
  it('fetches supported documents, preserves the extension, guesses the MIME', async () => {
    setItemWithAttachmentBytes()
    const snap = snapshot({
      attachments: [
        attachment({ id: 'a', name: 'Report Q3.pdf' }),
        attachment({ id: 'b', name: 'notes.docx' }),
        attachment({ id: 'c', name: 'photo.png', contentType: 'image/png' }),
        attachment({ id: 'd', name: 'archive.zip' }),
      ],
    })
    const files = await getDocumentAttachments(snap)
    expect(files.map((f) => f.filename)).toEqual(['Report_Q3.pdf', 'notes.docx'])
    expect(files[0].mimeType).toBe('application/pdf')
    expect(files[1].mimeType).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    expect(files[0].contentBase64).toBe('b64:a')
  })

  it('returns [] when the host cannot hand back attachment bytes', async () => {
    setOfficeItem({}) // no getAttachmentContentAsync
    const snap = snapshot({ attachments: [attachment({ id: 'a', name: 'x.pdf' })] })
    expect(await getDocumentAttachments(snap)).toEqual([])
  })
})

describe('getKnowledgeBaseAttachments', () => {
  it('includes real images + documents, skips tiny inline logos and cloud/oversized', async () => {
    setItemWithAttachmentBytes()
    const snap = snapshot({
      attachments: [
        attachment({ id: 'img', name: 'screenshot.png', contentType: 'image/png', size: 40_000 }),
        attachment({
          id: 'logo',
          name: 'logo.png',
          contentType: 'image/png',
          size: 2_000,
          isInline: true,
        }),
        attachment({ id: 'doc', name: 'contract.pdf' }),
        attachment({ id: 'big', name: 'huge.pdf', size: 40 * 1024 * 1024 }),
        attachment({ id: 'cloud', name: 'shared.pptx', attachmentType: 'cloud' }),
      ],
    })
    const files = await getKnowledgeBaseAttachments(snap)
    expect(files.map((f) => f.filename).sort()).toEqual(['contract.pdf', 'screenshot_png.png'])
  })

  it('countKnowledgeBaseAttachments matches without fetching bytes', () => {
    const snap = snapshot({
      attachments: [
        attachment({ id: 'img', name: 'a.png', contentType: 'image/png', size: 40_000 }),
        attachment({ id: 'doc', name: 'b.pdf' }),
        attachment({
          id: 'logo',
          name: 'c.png',
          contentType: 'image/png',
          size: 1_000,
          isInline: true,
        }),
      ],
    })
    expect(countKnowledgeBaseAttachments(snap)).toBe(2)
  })
})

describe('setComposeBody', () => {
  it('writes HTML to the compose body via setAsync and resolves true', async () => {
    const setAsync = vi.fn((_data: string, _opts: unknown, cb: (r: { status: string }) => void) =>
      cb({ status: 'succeeded' }),
    )
    setOfficeItem({ body: { setAsync } })
    const ok = await setComposeBody('<p>Hello</p>')
    expect(ok).toBe(true)
    expect(setAsync).toHaveBeenCalledTimes(1)
    expect(setAsync.mock.calls[0][0]).toBe('<p>Hello</p>')
    expect(setAsync.mock.calls[0][1]).toEqual({ coercionType: 'html' })
  })

  it('resolves false when there is no editable compose item (read mode)', async () => {
    setOfficeItem(undefined)
    expect(await setComposeBody('<p>x</p>')).toBe(false)
  })

  it('resolves false when the host reports a failed write', async () => {
    setOfficeItem({
      body: {
        setAsync: (_d: string, _o: unknown, cb: (r: { status: string }) => void) =>
          cb({ status: 'failed' }),
      },
    })
    expect(await setComposeBody('<p>x</p>')).toBe(false)
  })
})
