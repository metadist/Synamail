import { test, expect } from '@playwright/test'
import { bootTaskpane, homeCard, READ_ITEM, LIVE } from './support/harness'

/**
 * @ci — Contact AI Profiling. Home shows a Profiling card for the open mail.
 * "Create user profile" rolls that mail into the synamail plugin; "Full profile"
 * opens search, save, ask, and delete.
 */
test.describe('@ci contact-profile', () => {
  test.beforeEach(async ({ page }) => {
    await bootTaskpane(page, READ_ITEM)
  })

  test('creates a profile from the open email', async ({ page }) => {
    test.skip(LIVE, 'needs the synamail plugin installed on the live instance')
    const card = homeCard(page, 'Profiling')
    await expect(card.getByText(/No profile yet/)).toBeVisible()
    await card.getByRole('button', { name: 'Create user profile' }).click()
    await expect(card.locator('.pp__contact')).toContainText('1 email(s) profiled')
    await expect(card.locator('.pp__summary')).toContainText('Alice handles invoicing')
  })

  test.describe('full profile', () => {
    test.beforeEach(async ({ page }) => {
      await homeCard(page, 'Profiling').getByRole('button', { name: 'Full profile' }).click()
      await expect(page.getByRole('heading', { name: 'Contact AI Profiling' })).toBeVisible()
      await expect(page.getByRole('heading', { name: 'Saved emails' })).toBeVisible()
    })

    test('search lists this contact’s saved emails', async ({ page }) => {
      await page.getByPlaceholder(/Search this contact/).fill('invoice')
      await page.getByRole('button', { name: 'Search', exact: true }).click()
      const hits = page.locator('.cp__hit')
      if (!LIVE) {
        await expect(hits.first()).toBeVisible()
        await expect(hits.first()).toContainText('May invoice #4821')
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

    test('rolling profile: empty state → create from email → delete', async ({ page }) => {
      test.skip(LIVE, 'needs the synamail plugin installed on the live instance')
      const card = page.locator('.syn-card', {
        has: page.getByRole('heading', { name: 'Rolling profile' }),
      })
      await expect(card.getByText(/No profile yet/)).toBeVisible()
      await card.getByRole('button', { name: 'Create user profile' }).click()
      await expect(card.locator('.cp__as-of')).toContainText('1 email(s) profiled')
      await expect(card.getByText('Open loops')).toBeVisible()
      await expect(card.getByText(/confirm the May invoice/)).toBeVisible()
      page.on('dialog', (dialog) => dialog.accept())
      await card.getByRole('button', { name: 'Delete profile' }).click()
      await expect(card.getByText(/No profile yet/)).toBeVisible()
    })
  })
})
