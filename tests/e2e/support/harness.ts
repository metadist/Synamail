/**
 * L2 E2E harness entry point.
 *
 * `bootTaskpane(page, item)` brings the real taskpane up signed-in, with a
 * seeded Outlook item and a deterministic Synaplan backend, then returns once
 * the signed-in shell is visible. Specs then drive the real UI.
 *
 * Mode is controlled by env:
 *   - default                → mocked Synaplan (deterministic, CI-safe)
 *   - SYNAPLAN_E2E_LIVE=1     → real Synaplan; requires SYNAPLAN_BASE_URL and
 *                               SYNAPLAN_API_KEY (mint via the dashboard / API)
 *
 * See docs/E2E_OUTLOOK_AUTOMATION.md.
 */

import { expect, type Locator, type Page } from '@playwright/test'
import { installOfficeShim, type MailItemSeed, type SeedSettings } from './office-shim'
import { mockSynaplan } from './synaplan-mock'

export { READ_ITEM } from './office-shim'
export type { MailItemSeed } from './office-shim'

export const LIVE = process.env.SYNAPLAN_E2E_LIVE === '1'

const TASKPANE_PATH = '/src/taskpane/taskpane.html'

function settings(): SeedSettings {
  if (LIVE) {
    const baseUrl = process.env.SYNAPLAN_BASE_URL
    const apiKey = process.env.SYNAPLAN_API_KEY
    if (!baseUrl || !apiKey) {
      throw new Error(
        'SYNAPLAN_E2E_LIVE=1 requires SYNAPLAN_BASE_URL and SYNAPLAN_API_KEY to be set.',
      )
    }
    return { baseUrl, apiKey, keyId: 0, email: 'live-user@synaplan' }
  }
  return {
    baseUrl: 'https://synaplan.test',
    apiKey: 'e2e-test-key',
    keyId: 1,
    email: 'admin@synaplan.com',
  }
}

/**
 * Boot the taskpane with `item` selected. Installs the Office shim + (unless
 * live) the Synaplan mock, navigates, and waits for the signed-in shell.
 */
export async function bootTaskpane(page: Page, item: MailItemSeed | null): Promise<void> {
  // The taskpane HTML loads the real Office.js from the Microsoft CDN, which
  // would overwrite our injected shim (and provide an empty, item-less context
  // outside a real Outlook host). Block it so the shim is the Office runtime.
  await page.route('https://appsforoffice.microsoft.com/**', (route) =>
    route.fulfill({
      status: 200,
      contentType: 'application/javascript',
      body: '/* office.js blocked for the E2E Office shim */',
    }),
  )
  await installOfficeShim(page, { item, settings: settings() })
  if (!LIVE) await mockSynaplan(page)
  await page.goto(TASKPANE_PATH)
  // Signed-in shell renders the brand button + home nav icon.
  await expect(page.getByRole('button', { name: 'Home' }).first()).toBeVisible()
}

/** Calls the add-in made into the Office host (reply form, body writes, …). */
export async function officeCalls(page: Page): Promise<{ name: string; arg?: unknown }[]> {
  return page.evaluate(
    () =>
      (window as unknown as { __officeCalls?: { name: string; arg?: unknown }[] }).__officeCalls ??
      [],
  )
}

/** Home function card by its visible title (e.g. "Summarize email"). */
export function homeCard(page: Page, title: string): Locator {
  return page.locator('.syn-card', { has: page.getByRole('heading', { name: title }) })
}

/**
 * Wait until Home shows the read-mail affordances (summarise language buttons
 * and knowledge-base actions). Replaces the old "Email actions" accordion.
 */
export async function expectHomeReadReady(page: Page): Promise<void> {
  await expect(homeCard(page, 'Summarize email').getByRole('button').first()).toBeVisible()
  await expect(homeCard(page, 'Add to your sources').getByRole('button').first()).toBeVisible()
}
