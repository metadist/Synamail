import { test } from '@playwright/test'

/**
 * Contact AI Profiling E2E.
 *
 * The Home entry point (Email actions → "Contact AI Profiling") was removed
 * when Home was condensed to write / summarise / knowledge / chat. The
 * ContactProfile view still exists, but there is no user-reachable path from
 * the signed-in shell until Profiling is re-enabled on Home.
 *
 * Keep this file so the suite can be restored with the real navigation path
 * once that UI returns — do not drive a back-door router hook from CI.
 */
test.describe('@ci contact-profile', () => {
  test.skip(true, 'Profiling entry point temporarily disabled on Home')

  test('placeholder — re-enable when Profiling returns on Home', async () => {
    // Intentionally empty; the describe-level skip keeps @ci from timing out
    // waiting for the removed "Email actions" accordion.
  })
})
