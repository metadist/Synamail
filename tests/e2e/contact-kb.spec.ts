import { test, expect } from '@playwright/test'
import { bootTaskpane, openReadView, READ_ITEM, LIVE } from './support/harness'

/**
 * @ci — Contact knowledge base: search this sender's saved emails, save the
 * current email into the contact group, and ask a grounded question. Reaches
 * the view via Read → "Contact knowledge base" (the real user path).
 */
test.describe('@ci contact-kb', () => {
  test.beforeEach(async ({ page }) => {
    await bootTaskpane(page, READ_ITEM)
    await openReadView(page)
    await page.getByRole('button', { name: 'Contact knowledge base' }).click()
    await expect(page.getByRole('heading', { name: 'Saved emails' })).toBeVisible()
  })

  test('auto-search lists this contact’s saved emails', async ({ page }) => {
    const hits = page.locator('.ckb__hit')
    if (!LIVE) {
      await expect(hits.first()).toBeVisible()
      await expect(hits.first()).toContainText('May invoice #4821')
      // Live search hits carry no filename → the view shows the fallback label.
      await expect(hits.first()).toContainText('Saved email')
    }
  })

  test('save current email reports success', async ({ page }) => {
    await page.getByRole('button', { name: 'Save current email here' }).click()
    if (!LIVE) await expect(page.getByText(/Saved to/i)).toBeVisible()
  })

  test('ask about this contact returns an answer', async ({ page }) => {
    await page.getByPlaceholder(/Ask about/).fill('What was the last invoice?')
    await page.getByRole('button', { name: 'Ask', exact: true }).click()
    const turn = page.locator('.ckb__turn').last()
    await expect(turn).toBeVisible()
    await expect(turn).toContainText('What was the last invoice?')
  })
})
