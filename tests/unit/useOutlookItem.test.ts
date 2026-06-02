import { describe, expect, it } from 'vitest'
import { buildEmailText, getReadItemAsFile } from '@/taskpane/composables/useOutlookItem'
import type { OutlookItemSnapshot } from '@/taskpane/composables/useOutlookItem'

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
})
