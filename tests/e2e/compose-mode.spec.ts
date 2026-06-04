import { test, expect } from '@playwright/test'
import { bootTaskpane, openComposeView, officeCalls, COMPOSE_ITEM, LIVE } from './support/harness'

/**
 * @ci — Compose-mode flows. Each assertion confirms the add-in asked Outlook
 * to write the right thing into the draft (body.setAsync / setSelectedDataAsync),
 * recorded by the Office shim.
 */
test.describe('@ci compose-mode', () => {
  test.beforeEach(async ({ page }) => {
    await bootTaskpane(page, COMPOSE_ITEM)
    await openComposeView(page)
  })

  test('draft from prompt writes the body', async ({ page }) => {
    await page.getByPlaceholder('Describe what you want to write…').fill('Approve the invoice')
    await page.getByRole('button', { name: 'Draft from prompt' }).click()
    await expect
      .poll(
        async () => (await officeCalls(page)).find((c) => c.name === 'body.setAsync')?.arg ?? '',
      )
      .not.toBe('')
    if (!LIVE) {
      const calls = await officeCalls(page)
      const body = calls.find((c) => c.name === 'body.setAsync')?.arg as string
      expect(body).toContain('approved')
    }
  })

  test('improve selection replaces the selected text', async ({ page }) => {
    await page.getByRole('button', { name: 'Improve selection' }).click()
    await expect
      .poll(async () =>
        (await officeCalls(page)).some((c) => c.name === 'body.setSelectedDataAsync'),
      )
      .toBe(true)
    if (!LIVE) {
      const calls = await officeCalls(page)
      const sel = calls.find((c) => c.name === 'body.setSelectedDataAsync')?.arg as string
      expect(sel).toContain('Payment is due Friday')
    }
  })

  test('insert from knowledge base puts a snippet into the draft', async ({ page }) => {
    await page.getByPlaceholder('Search your knowledge base…').fill('invoice')
    await page.getByPlaceholder('Search your knowledge base…').press('Enter')
    const hits = page.locator('.compose__hits li')
    await expect(hits.first()).toBeVisible()
    await hits.first().getByRole('button').click()
    await expect
      .poll(async () =>
        (await officeCalls(page)).some((c) => c.name === 'body.setSelectedDataAsync'),
      )
      .toBe(true)
    if (!LIVE) {
      const calls = await officeCalls(page)
      const sel = calls.find((c) => c.name === 'body.setSelectedDataAsync')?.arg as string
      expect(sel).toContain('May invoice #4821')
    }
  })
})
