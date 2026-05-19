/**
 * Auth relay dialog page.
 *
 * In v0.1 (offline / mock mode) this immediately posts a fake `apiKey` back
 * to the parent taskpane so the sign-in loop closes without a live Synaplan
 * round-trip.
 *
 * In Sprint 2.8 this file is replaced by the live bridge page served from
 * `synaplan/frontend/` (Vue route `/addin/connect`), which:
 *   1. Redirects to the existing Synaplan login if no cookie session.
 *   2. Calls `POST /api/v1/apikeys`.
 *   3. Posts the issued `apiKey` + `keyId` + `email` + `baseUrl` back via
 *      `Office.context.ui.messageParent`.
 *
 * The taskpane validates the `state` nonce. See useAuth.ts.
 */

const params = new URLSearchParams(window.location.search)
const state = params.get('state') ?? ''
const baseUrl = params.get('baseUrl') ?? 'https://web.synaplan.com'

Office.onReady(() => {
  try {
    const payload = {
      state,
      apiKey: 'mock-key-' + Math.random().toString(36).slice(2),
      keyId: 1,
      email: 'demo@synaplan.test',
      baseUrl,
    }
    Office.context.ui.messageParent(JSON.stringify(payload))
  } catch (err) {
    const e = document.getElementById('err')
    if (e) {
      e.hidden = false
      e.textContent = err instanceof Error ? err.message : String(err)
    }
  }
})
