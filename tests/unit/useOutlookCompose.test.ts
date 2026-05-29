import { describe, expect, it } from 'vitest'
import { canDisplayNewMessage, displayNewMessage } from '@/taskpane/composables/useOutlookCompose'

type OfficeStub = { context: { mailbox: Record<string, unknown> } }

function officeStub(): OfficeStub {
  return (globalThis as unknown as { Office: OfficeStub }).Office
}

describe('useOutlookCompose', () => {
  it('reports unsupported when the host lacks displayNewMessageFormAsync', () => {
    expect(canDisplayNewMessage()).toBe(false)
  })

  it('rejects displayNewMessage when the host is unsupported', async () => {
    await expect(displayNewMessage({ subject: 's', htmlBody: '<p>b</p>' })).rejects.toThrow(
      /cannot open a new message/i,
    )
  })

  it('opens a new message when the host supports it', async () => {
    let received: { subject?: string; htmlBody?: string } | null = null
    officeStub().context.mailbox.displayNewMessageFormAsync = (
      params: { subject?: string; htmlBody?: string },
      cb?: (r: { status: string }) => void,
    ) => {
      received = params
      cb?.({ status: 'succeeded' })
    }
    expect(canDisplayNewMessage()).toBe(true)
    await displayNewMessage({ subject: 'Hi', htmlBody: '<p>Body</p>' })
    expect(received).toEqual({ toRecipients: undefined, subject: 'Hi', htmlBody: '<p>Body</p>' })
  })
})
