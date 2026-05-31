/**
 * Auth relay dialog page — runs on the ADD-IN's own origin.
 *
 * This page exists so that `Office.context.ui.messageParent` is always called
 * from the same origin as the taskpane. Cross-origin `messageParent` (calling
 * it from the Synaplan bridge on a different origin) is silently dropped by
 * Outlook desktop after the login navigations, which left the panel stuck.
 *
 * The Synaplan bridge redirects back here with the sign-in payload in the URL
 * fragment (`#payload=<base64 JSON>`). We decode it and hand it to the parent
 * taskpane via messageParent. There is no mock/offline mode — sign-in is
 * always a real round-trip to the chosen Synaplan server.
 */

const params = new URLSearchParams(window.location.search)
const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ''))

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
    const returned = hashParams.get('payload') ?? params.get('payload')
    if (returned) {
      Office.context.ui.messageParent(decodeBase64Json(returned))
      return
    }

    showError('No sign-in payload was received from Synaplan. Please try signing in again.')
  } catch (err) {
    showError(err instanceof Error ? err.message : String(err))
  }
})
