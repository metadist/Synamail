/**
 * Injectable Office.js shim for the L2 E2E harness.
 *
 * The real taskpane renders inside an Office runtime (OWA iframe / desktop
 * WebView2). For automated UI tests we run the *real* Vue taskpane in plain
 * Playwright Chromium and feed it a faithful, deterministic `Office` global
 * via `page.addInitScript` — installed BEFORE any app script runs, exactly as
 * the host would. We also pre-seed the roaming-settings store with a valid
 * API key so the app boots straight into the signed-in Home view.
 *
 * Every Office mutation the views trigger (reply form, compose body writes,
 * new-message / appointment forms) is recorded on `window.__officeCalls` so a
 * spec can assert the add-in asked Outlook to do the right thing.
 *
 * See docs/E2E_OUTLOOK_AUTOMATION.md for the layered automation strategy and
 * how this maps onto the real Outlook clients.
 */

import type { Page } from '@playwright/test'

export interface MailItemSeed {
  mode: 'read' | 'compose'
  subject: string
  /** Sender address (read mode). Empty for compose. */
  from: string
  to: string[]
  cc: string[]
  conversationId: string
  /** Read-mode body text, or the "currently selected" text in compose mode. */
  bodyText: string
}

export interface SeedSettings {
  apiKey: string
  keyId: number
  email: string
  baseUrl: string
}

export interface OfficeSeed {
  item: MailItemSeed | null
  settings: SeedSettings
}

/** A read-mode message about an invoice, with a known sender. */
export const READ_ITEM: MailItemSeed = {
  mode: 'read',
  subject: 'May invoice #4821 — please confirm',
  from: 'alice@contoso.com',
  to: ['me@example.com'],
  cc: [],
  conversationId: 'conv-e2e-read-1',
  bodyText:
    'Hi, could you confirm the May invoice #4821 is approved for payment by Friday? ' +
    'We can also meet next Tuesday at 3pm to review the figures. Thanks, Alice.',
}

/** A compose-mode draft whose "selection" is a rough sentence to transform. */
export const COMPOSE_ITEM: MailItemSeed = {
  mode: 'compose',
  subject: 'Re: May invoice',
  from: '',
  to: ['alice@contoso.com'],
  cc: [],
  conversationId: 'conv-e2e-compose-1',
  bodyText: 'thanks i will look into the invoice and get back asap',
}

/**
 * Install the Office shim + seeded auth into a page. Call BEFORE `page.goto`.
 * The function passed to `addInitScript` is serialised and runs in the browser
 * with `seed` as its only argument, so it must be fully self-contained.
 */
export async function installOfficeShim(page: Page, seed: OfficeSeed): Promise<void> {
  await page.addInitScript((s: OfficeSeed) => {
    type AsyncCb<T> = (r: { status: string; value?: T; error?: { message: string } }) => void
    const win = window as unknown as {
      Office?: unknown
      __officeCalls?: { name: string; arg?: unknown }[]
    }
    const calls: { name: string; arg?: unknown }[] = []
    win.__officeCalls = calls
    const record = (name: string, arg?: unknown): void => {
      calls.push({ name, arg })
    }
    const ok = <T>(cb: AsyncCb<T>, value?: T): void => cb({ status: 'succeeded', value })

    const SETTINGS_KEY = 'synamail.settings'
    // Seed the signed-in identity (roaming settings are localStorage-backed in
    // this shim, mirroring the app's own non-Office fallback).
    try {
      const existing = window.localStorage.getItem(SETTINGS_KEY)
      if (!existing) {
        window.localStorage.setItem(SETTINGS_KEY, JSON.stringify({ ...s.settings, language: 'en' }))
      }
    } catch {
      /* localStorage may be unavailable — the in-memory roaming map still works */
    }

    const toEmailObj = (e: string): { emailAddress: string; displayName: string } => ({
      emailAddress: e,
      displayName: e,
    })

    const buildItem = (seedItem: MailItemSeed): Record<string, unknown> => {
      const common = {
        itemType: 'message',
        subject: seedItem.subject,
        conversationId: seedItem.conversationId,
        attachments: [] as unknown[],
      }
      if (seedItem.mode === 'read') {
        return {
          ...common,
          from: toEmailObj(seedItem.from),
          to: seedItem.to.map(toEmailObj),
          cc: seedItem.cc.map(toEmailObj),
          body: {
            getAsync: (_coercion: unknown, cb: AsyncCb<string>) => ok(cb, seedItem.bodyText),
          },
          displayReplyForm: (arg: unknown) => record('displayReplyForm', arg),
          getAsFileAsync: (cb: AsyncCb<string>) =>
            ok(
              cb,
              btoa(
                unescape(
                  encodeURIComponent(
                    `Subject: ${seedItem.subject}\r\nFrom: ${seedItem.from}\r\n\r\n${seedItem.bodyText}`,
                  ),
                ),
              ),
            ),
        }
      }
      // compose
      return {
        ...common,
        subject: { getAsync: (cb: AsyncCb<string>) => ok(cb, seedItem.subject) },
        to: { getAsync: (cb: AsyncCb<unknown[]>) => ok(cb, seedItem.to.map(toEmailObj)) },
        cc: { getAsync: (cb: AsyncCb<unknown[]>) => ok(cb, seedItem.cc.map(toEmailObj)) },
        body: {
          setAsync: (html: string, _opts: unknown, cb: AsyncCb<void>) => {
            record('body.setAsync', html)
            ok(cb)
          },
          setSelectedDataAsync: (text: string, _opts: unknown, cb: AsyncCb<void>) => {
            record('body.setSelectedDataAsync', text)
            ok(cb)
          },
          getAsync: (_coercion: unknown, cb: AsyncCb<string>) => ok(cb, seedItem.bodyText),
        },
      }
    }

    const roaming = {
      get(key: string): unknown {
        try {
          const raw = window.localStorage.getItem(key)
          return raw ? JSON.parse(raw) : undefined
        } catch {
          return undefined
        }
      },
      set(key: string, value: unknown): void {
        try {
          window.localStorage.setItem(key, JSON.stringify(value))
        } catch {
          /* no-op */
        }
      },
      remove(key: string): void {
        try {
          window.localStorage.removeItem(key)
        } catch {
          /* no-op */
        }
      },
      saveAsync(cb?: AsyncCb<void>): void {
        if (cb) ok(cb)
      },
    }

    const office = {
      onReady: (cb?: (info: { host: string; platform: string }) => void) => {
        if (cb) cb({ host: 'Outlook', platform: 'OfficeOnline' })
        return Promise.resolve({ host: 'Outlook', platform: 'OfficeOnline' })
      },
      EventType: {
        ItemChanged: 'olkItemSelectedChanged',
        DialogMessageReceived: 'dialogMessageReceived',
        DialogEventReceived: 'dialogEventReceived',
      },
      MailboxEnums: {
        ItemType: { Message: 'message', Appointment: 'appointment' },
        AttachmentContentFormat: { Base64: 'base64', Url: 'url', Eml: 'eml' },
      },
      CoercionType: { Text: 'text', Html: 'html' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      context: {
        displayLanguage: 'en-US',
        mailbox: {
          item: s.item ? buildItem(s.item) : undefined,
          userProfile: { emailAddress: s.settings.email, displayName: 'E2E User' },
          addHandlerAsync: (_type: unknown, _handler: unknown, cb?: AsyncCb<void>) => {
            if (cb) ok(cb)
          },
          removeHandlerAsync: (_type: unknown, cb?: AsyncCb<void>) => {
            if (cb) ok(cb)
          },
          displayNewMessageFormAsync: (params: unknown, cb?: AsyncCb<void>) => {
            record('displayNewMessageFormAsync', params)
            if (cb) ok(cb)
          },
          displayNewAppointmentForm: (params: unknown) =>
            record('displayNewAppointmentForm', params),
        },
        roamingSettings: roaming,
        ui: {
          displayDialogAsync: () => record('displayDialogAsync'),
          messageParent: (msg: unknown) => record('messageParent', msg),
        },
      },
    }

    win.Office = office
  }, seed)
}
