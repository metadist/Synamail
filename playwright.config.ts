import { defineConfig, devices } from '@playwright/test'

/**
 * Playwright config for sideloaded add-in E2E.
 *
 * The CI subset (grep `@ci`) runs against the Vite dev server only and stubs
 * the Office runtime; the full suite (Sprint 2.9 onwards) targets a real
 * Outlook on the Web session.
 */
export default defineConfig({
  testDir: './tests/e2e',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  reporter: process.env.CI ? [['html'], ['github']] : 'list',
  use: {
    baseURL: process.env.BASE_URL ?? 'https://localhost:3000',
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
    ignoreHTTPSErrors: true,
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
})
