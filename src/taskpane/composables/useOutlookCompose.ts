/**
 * Helpers for opening Outlook compose surfaces from the taskpane.
 *
 * `displayNewMessage` opens a brand-new compose window prefilled with an
 * AI-drafted subject + body. Uses `Office.context.mailbox.displayNewMessageFormAsync`
 * (Mailbox 1.0 — no extra manifest permission required).
 */

export interface NewMessageOptions {
  subject: string
  /** HTML body content (no doctype/html/head wrapper). */
  htmlBody: string
  toRecipients?: string[]
}

interface NewMessageHost {
  displayNewMessageFormAsync?: (
    parameters: { toRecipients?: string[]; subject?: string; htmlBody?: string },
    callback?: (result: Office.AsyncResult<void>) => void,
  ) => void
}

/** True when the host can open a new compose window. */
export function canDisplayNewMessage(): boolean {
  const mailbox = (typeof Office !== 'undefined' ? Office.context?.mailbox : undefined) as
    | NewMessageHost
    | undefined
  return typeof mailbox?.displayNewMessageFormAsync === 'function'
}

export function displayNewMessage(opts: NewMessageOptions): Promise<void> {
  const mailbox = (typeof Office !== 'undefined' ? Office.context?.mailbox : undefined) as
    | NewMessageHost
    | undefined
  if (!mailbox || typeof mailbox.displayNewMessageFormAsync !== 'function') {
    return Promise.reject(
      new Error('This Outlook host cannot open a new message window from an add-in.'),
    )
  }
  return new Promise<void>((resolve, reject) => {
    mailbox.displayNewMessageFormAsync!(
      {
        toRecipients: opts.toRecipients,
        subject: opts.subject,
        htmlBody: opts.htmlBody,
      },
      (r) => {
        if (!r || r.status === Office.AsyncResultStatus.Succeeded) resolve()
        else reject(new Error(r.error?.message ?? 'Could not open a new message window.'))
      },
    )
  })
}
