import { describe, expect, it } from 'vitest'
import { getReadItemAsFile } from '@/taskpane/composables/useOutlookItem'
import type { OutlookItemSnapshot } from '@/taskpane/composables/useOutlookItem'

function snapshot(overrides: Partial<OutlookItemSnapshot> = {}): OutlookItemSnapshot {
  return {
    mode: 'read',
    subject: 'Q3 plan',
    to: ['team@example.test'],
    cc: [],
    bodyText: 'Hello world',
    attachments: [],
    ...overrides,
  }
}

describe('getReadItemAsFile', () => {
  it('uses getAsFileAsync when the host exposes it (Mailbox 1.8+)', async () => {
    ;(globalThis as unknown as { Office: { context: { mailbox: { item: unknown } } } }).Office =
      Object.assign({}, (globalThis as unknown as { Office: object }).Office, {
        context: {
          ...((globalThis as unknown as { Office: { context: object } }).Office.context as object),
          mailbox: {
            item: {
              getAsFileAsync: (cb: (r: { status: string; value: string }) => void) =>
                cb({ status: 'succeeded', value: 'BASE64EMLDATA' }),
            },
          },
        },
      })

    const file = await getReadItemAsFile(snapshot({ subject: 'Q3 plan!' }))
    expect(file.mimeType).toBe('message/rfc822')
    expect(file.filename).toBe('Q3_plan.eml')
    expect(file.contentBase64).toBe('BASE64EMLDATA')
  })

  it('falls back to plain-text export when getAsFileAsync is missing', async () => {
    // The default setup.ts stub has no getAsFileAsync — use it as-is.
    const file = await getReadItemAsFile(snapshot({ subject: '', bodyText: 'hello' }))
    expect(file.filename).toBe('message.txt')
    expect(file.mimeType).toBe('text/plain')
    expect(atob(file.contentBase64)).toBe('hello')
  })

  it('surfaces getAsFileAsync errors as rejections', async () => {
    ;(globalThis as unknown as { Office: { context: { mailbox: { item: unknown } } } }).Office =
      Object.assign({}, (globalThis as unknown as { Office: object }).Office, {
        context: {
          ...((globalThis as unknown as { Office: { context: object } }).Office.context as object),
          mailbox: {
            item: {
              getAsFileAsync: (cb: (r: { status: string; error?: { message: string } }) => void) =>
                cb({ status: 'failed', error: { message: 'permission denied' } }),
            },
          },
        },
      })

    await expect(getReadItemAsFile(snapshot())).rejects.toThrow(/permission denied/)
  })

  it('sanitises the subject when building the filename', async () => {
    const file = await getReadItemAsFile(snapshot({ subject: 'Hello / World? *!' }))
    expect(file.filename).toBe('Hello_World.txt')
  })
})
