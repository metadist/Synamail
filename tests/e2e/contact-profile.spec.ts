import { test, expect } from '@playwright/test'
import { bootTaskpane, openReadView, READ_ITEM, LIVE } from './support/harness'

/**
 * @ci — Contact AI Profiling: search this sender's saved emails, save the
 * current email into the contact group, and ask a grounded question. Reaches
 * the view via Read → "Contact AI Profiling" (the real user path).
 */
test.describe('@ci contact-profile', () => {
  test.beforeEach(async ({ page }) => {
    await bootTaskpane(page, READ_ITEM)
    await openReadView(page)
    await page.getByRole('button', { name: 'Contact AI Profiling' }).click()
    await expect(page.getByRole('heading', { name: 'Saved emails' })).toBeVisible()
  })

  test('search lists this contact’s saved emails', async ({ page }) => {
    // Searching requires a query term (no empty-query auto-search on open).
    await page.getByPlaceholder(/Search this contact/).fill('invoice')
    await page.getByRole('button', { name: 'Search', exact: true }).click()
    const hits = page.locator('.cp__hit')
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
    const turn = page.locator('.cp__turn').last()
    await expect(turn).toBeVisible()
    await expect(turn).toContainText('What was the last invoice?')
  })

  test('rolling profile: empty state → update from email → delete', async ({ page }) => {
    test.skip(LIVE, 'needs the synamail plugin installed on the live instance')
    const card = page.locator('.syn-card', {
      has: page.getByRole('heading', { name: 'Rolling profile' }),
    })
    // No profile yet — empty state with the update affordance.
    await expect(card.getByText(/No profile yet/)).toBeVisible()
    // Roll the open email in.
    await card.getByRole('button', { name: 'Update profile from this email' }).click()
    await expect(card.locator('.cp__as-of')).toContainText('1 email(s) profiled')
    await expect(card.getByText('Open loops')).toBeVisible()
    await expect(card.getByText(/confirm the May invoice/)).toBeVisible()
    // Delete brings the empty state back.
    page.on('dialog', (dialog) => dialog.accept())
    await card.getByRole('button', { name: 'Delete profile' }).click()
    await expect(card.getByText(/No profile yet/)).toBeVisible()
  })
})
