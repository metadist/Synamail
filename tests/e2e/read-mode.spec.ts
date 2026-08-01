import { test, expect } from '@playwright/test'
import {
  bootTaskpane,
  expectHomeReadReady,
  homeCard,
  officeCalls,
  READ_ITEM,
  LIVE,
} from './support/harness'

/**
 * @ci — Home read-mail flows against the condensed taskpane (write / summarise /
 * knowledge / chat). Set SYNAPLAN_E2E_LIVE=1 to run the same UI path against a
 * real instance (content assertions relax to "a result appeared").
 */
test.describe('@ci read-mode', () => {
  test.beforeEach(async ({ page }) => {
    await bootTaskpane(page, READ_ITEM)
    await expectHomeReadReady(page)
  })

  test('summarise opens a new email with the summary', async ({ page }) => {
    const card = homeCard(page, 'Summarize email')
    await card.getByRole('button', { name: 'English' }).click()
    await expect(card.getByText(/Summary opened/i)).toBeVisible()
    await expect
      .poll(async () => (await officeCalls(page)).some((c) => c.name === 'displayNewMessageForm'))
      .toBe(true)
    if (!LIVE) {
      const calls = await officeCalls(page)
      const call = calls.find((c) => c.name === 'displayNewMessageForm')
      const html = String((call?.arg as { htmlBody?: string } | undefined)?.htmlBody ?? '')
      expect(html).toContain('Alice asks to confirm May invoice #4821')
    }
  })

  test('write reply opens an Outlook reply form', async ({ page }) => {
    const card = homeCard(page, 'Write email or reply')
    await card.getByPlaceholder('Topic of your mail').fill('Confirm the invoice')
    await card.getByRole('button', { name: 'Concise' }).click()
    await expect(card.getByText(/Opened in a new email window/i)).toBeVisible()
    await expect
      .poll(async () => (await officeCalls(page)).some((c) => c.name === 'displayReplyForm'))
      .toBe(true)
  })

  test('add to sources saves via the New dialog', async ({ page }) => {
    const card = homeCard(page, 'Add to your sources')
    await card.getByRole('button', { name: 'New' }).click()
    const dialog = page.getByRole('dialog')
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: 'Save', exact: true }).click()
    if (!LIVE) await expect(card.getByText(/Saved to/i)).toBeVisible()
  })

  test('home chat answers with the open email as context', async ({ page }) => {
    const card = homeCard(page, 'Chat')
    await card.getByPlaceholder('Ask synaplan').fill('Do I need to pay this?')
    await card.getByRole('button', { name: 'Send' }).click()
    await expect(card.getByText('Do I need to pay this?')).toBeVisible()
    if (!LIVE) {
      await expect(card.getByText(/invoice #4821/i)).toBeVisible()
    } else {
      // Live: any non-empty AI bubble is enough.
      await expect(card.locator('.chat__bubble--ai').last()).not.toHaveText('')
    }
  })
})
