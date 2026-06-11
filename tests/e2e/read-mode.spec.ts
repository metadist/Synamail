import { test, expect } from '@playwright/test'
import {
  bootTaskpane,
  openReadView,
  runEmailAction,
  officeCalls,
  READ_ITEM,
  LIVE,
} from './support/harness'

/**
 * @ci — Read-mode flows driven through the real taskpane against a mocked
 * Synaplan (deterministic). Set SYNAPLAN_E2E_LIVE=1 to run the same UI path
 * against a real instance (content assertions relax to "a result appeared").
 */
test.describe('@ci read-mode', () => {
  test.beforeEach(async ({ page }) => {
    await bootTaskpane(page, READ_ITEM)
    await openReadView(page)
  })

  test('summarise renders bullet points', async ({ page }) => {
    await runEmailAction(page, 'Summarize')
    const result = page.locator('.ea__result')
    await expect(result).toBeVisible()
    if (!LIVE) await expect(result).toContainText('Alice asks to confirm May invoice #4821')
  })

  test('translate renders a translation', async ({ page }) => {
    await runEmailAction(page, 'Translate')
    const result = page.locator('.ea__result')
    await expect(result).toBeVisible()
    if (!LIVE) await expect(result).toContainText('Hallo')
  })

  test('classify renders a category + confidence', async ({ page }) => {
    await runEmailAction(page, 'Classify')
    const result = page.locator('.ea__result')
    await expect(result).toBeVisible()
    if (!LIVE) await expect(result).toContainText('support (91%)')
  })

  test('draft reply opens an Outlook reply form', async ({ page }) => {
    await runEmailAction(page, 'Draft reply')
    await expect
      .poll(async () => (await officeCalls(page)).some((c) => c.name === 'displayReplyForm'))
      .toBe(true)
  })

  test('save to knowledge base confirms with a file id', async ({ page }) => {
    await runEmailAction(page, 'Save to knowledge base')
    const dialog = page.getByRole('dialog')
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: 'Save', exact: true }).click()
    if (!LIVE) await expect(page.getByText(/Saved as file/i)).toBeVisible()
  })

  test('ask threads a grounded answer', async ({ page }) => {
    await page.getByPlaceholder('Ask about this email…').fill('Do I need to pay this?')
    await page.getByRole('button', { name: 'Ask', exact: true }).click()
    const turn = page.locator('.ea__turn').last()
    await expect(turn).toBeVisible()
    await expect(turn).toContainText('Do I need to pay this?')
    if (!LIVE) await expect(turn).toContainText('invoice #4821')
  })

  test('find meeting times proposes a slot', async ({ page }) => {
    await runEmailAction(page, 'Find meeting times')
    const slots = page.locator('.ea__slots')
    if (!LIVE) {
      await expect(slots).toBeVisible()
      await expect(slots).toContainText('Invoice review')
    }
  })
})
