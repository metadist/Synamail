import { chromium } from '@playwright/test'

/**
 * Warm the Vite dev server before the suite runs.
 *
 * The webServer is `vite dev`, which transforms modules on first request. The
 * first browser load of the taskpane compiles the whole graph (~100 modules +
 * Vue SFCs), which can exceed a single test's timeout on a cold/loaded CI box.
 * Loading it once here means every test then hits a warm cache.
 */
async function globalSetup(): Promise<void> {
  const baseURL = process.env.BASE_URL ?? 'https://localhost:3000'
  const browser = await chromium.launch()
  try {
    const page = await browser.newPage({ ignoreHTTPSErrors: true })
    await page.goto(`${baseURL}/src/taskpane/taskpane.html`, {
      waitUntil: 'networkidle',
      timeout: 90_000,
    })
    // Give Vite a beat to finish any trailing on-demand transforms.
    await page.waitForTimeout(1000)
  } finally {
    await browser.close()
  }
}

export default globalSetup
