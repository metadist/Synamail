import { defineConfig, devices } from '@playwright/test'

/**
 * Playwright config for sideloaded add-in E2E.
 *
 * The CI subset (grep `@ci`) runs against the Vite dev server only and stubs
 * the Office runtime; the full suite (Sprint 2.9 onwards) targets a real
 * Outlook on the Web session.
 *
 * Locally, Vite serves HTTPS on :3000 (after `npx office-addin-dev-certs install`)
 * and `BASE_URL` defaults to https. In CI we set `BASE_URL=http://localhost:3000`
 * because the runner has no Office dev certs provisioned, so Vite gracefully
 * falls back to plain HTTP — which is fine for the smoke test (no Office runtime,
 * just a `page.goto`). See `.github/workflows/ci.yml` (E2E job).
 */
const BASE_URL = process.env.BASE_URL ?? 'https://localhost:3000'

export default defineConfig({
  testDir: './tests/e2e',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  reporter: process.env.CI ? [['html'], ['github']] : 'list',
  use: {
    baseURL: BASE_URL,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
    ignoreHTTPSErrors: true,
  },
  webServer: {
    command: 'npm run dev',
    url: `${BASE_URL}/src/taskpane/taskpane.html`,
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
    ignoreHTTPSErrors: true,
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
})
