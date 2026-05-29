/**
 * Auth relay dialog page — runs on the ADD-IN's own origin.
 *
 * This page exists so that `Office.context.ui.messageParent` is always called
 * from the same origin as the taskpane. Cross-origin `messageParent` (calling
 * it from the Synaplan bridge on a different origin) is silently dropped by
 * Outlook desktop after the login navigations, which left the panel stuck.
 *
 * Two modes:
 *   1. Return mode — the Synaplan bridge redirected back here with the
 *      sign-in payload in the URL fragment (`#payload=<base64 JSON>`). We
 *      decode it and hand it to the parent taskpane via messageParent.
 *   2. Mock mode (`?mock=1`, dev only) — immediately post a fake key so the
 *      offline dev loop closes without a real Synaplan round-trip.
 */

const params = new URLSearchParams(window.location.search)
const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ''))
const state = params.get('state') ?? ''
const baseUrl = params.get('baseUrl') ?? 'https://web.synaplan.com'

function showError(message: string): void {
  const e = document.getElementById('err')
  if (e) {
    e.hidden = false
    e.textContent = message
  }
}

function decodeBase64Json(b64: string): string {
  // Inverse of btoa(unescape(encodeURIComponent(json))) on the bridge side.
  return decodeURIComponent(escape(atob(b64)))
}

Office.onReady(() => {
  try {
    // Mode 1: returning from the Synaplan bridge with the payload.
    const returned = hashParams.get('payload') ?? params.get('payload')
    if (returned) {
      Office.context.ui.messageParent(decodeBase64Json(returned))
      return
    }

    // Mode 2: dev mock.
    if (params.get('mock') === '1') {
      const payload = {
        state,
        apiKey: 'mock-key-' + Math.random().toString(36).slice(2),
        keyId: 1,
        email: 'demo@synaplan.test',
        baseUrl,
      }
      Office.context.ui.messageParent(JSON.stringify(payload))
      return
    }

    showError('No sign-in payload was received from Synaplan. Please try signing in again.')
  } catch (err) {
    showError(err instanceof Error ? err.message : String(err))
  }
})
