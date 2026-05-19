import { test, expect } from '@playwright/test'

/**
 * @ci — runs against the Vite dev server (Office stubbed by tests/setup.ts
 * equivalent in the HTML page). Sprint 2.9 adds the live OWA flow.
 */
test.describe('@ci smoke', () => {
  test('taskpane page loads', async ({ page }) => {
    await page.goto('/src/taskpane/taskpane.html')
    await expect(page.locator('#app')).toBeVisible()
  })
})
