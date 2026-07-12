/**
 * Taskpane entry point. Waits for Office.onReady, hydrates auth state, then
 * mounts the Vue app. Listens for the AUTH_INVALIDATED_EVENT to bounce the
 * user back to SignIn on 401.
 */

import { createApp } from 'vue'
import App from './App.vue'
import { i18n, setLocale, SUPPORTED_LOCALES, type Locale } from '@/i18n'
import { hydrateAuthState, isSignedIn } from './composables/useAuth'
import { loadSettings } from './composables/useRoamingSettings'
import { useOutlookItem } from './composables/useOutlookItem'
import { currentView, go } from './router'
import { AUTH_INVALIDATED_EVENT } from './composables/useSynaplanClient'
import '@/taskpane/styles/tokens.css'
import '@/taskpane/styles/app.css'

/**
 * If `Office.onReady` hasn't resolved this long after load, we stop waiting and
 * surface an actionable error instead of an endless spinner. Kept below the
 * 12s inline fallback in taskpane.html so this (more specific) message wins.
 *
 * The dominant reason `Office.onReady` never fires in Outlook on the web is
 * Safari's Intelligent Tracking Prevention: the add-in runs in a third-party
 * iframe whose storage/cookies WebKit fully partitions or blocks, which can
 * stall the Office.js runtime. Chrome/Edge don't do this, which is why the same
 * add-in opens there but appears "stuck" in Safari.
 */
const OFFICE_READY_TIMEOUT_MS = 9000

/**
 * Advice shown whenever boot fails inside Outlook. Safari's default
 * "Prevent cross-site tracking" is the single most common cause of an add-in
 * that loads in Chrome but not Safari (see docs/ARCHITECTURE.md test matrix).
 */
const SAFARI_TRACKING_HINT =
  'If you are using Safari, open Settings → Privacy and turn off ' +
  '"Prevent cross-site tracking", then close and reopen the add-in.'

/**
 * Replace the inline boot spinner (taskpane.html `#syn-boot`) with a visible,
 * actionable error. Without this a boot failure in Safari is invisible — the
 * pane just spins forever and reads as "the add-in won't open".
 *
 * We build the DOM with `textContent` (never interpolate an error string into
 * innerHTML) and drop the element's `id` so the 12s inline fallback timer in
 * taskpane.html bails instead of overwriting this message.
 */
function showBootError(message: string): void {
  const boot = document.getElementById('syn-boot')
  if (!boot) return

  boot.removeAttribute('id')
  ;(window as unknown as { __synamailHandled?: boolean }).__synamailHandled = true

  const wrap = document.createElement('div')
  wrap.className = 'syn-boot__error'
  wrap.setAttribute('role', 'alert')

  const msg = document.createElement('p')
  msg.style.margin = '0 0 0.5rem'
  msg.textContent = message
  wrap.appendChild(msg)

  const help = document.createElement('p')
  help.style.margin = '0'
  const link = document.createElement('a')
  link.href = 'https://web.synaplan.com/support'
  link.target = '_blank'
  link.rel = 'noopener'
  link.textContent = 'web.synaplan.com/support'
  help.append('Still stuck? Visit ', link, '.')
  wrap.appendChild(help)

  boot.replaceChildren(wrap)
}

function mount(): void {
  const app = createApp(App)
  app.use(i18n)
  app.mount('#app')
}

/**
 * Honour a saved UI-language preference once roaming settings are available.
 * `'auto'` (or no preference) leaves the Outlook-detected locale that i18n
 * was created with; an explicit shipped locale overrides it.
 */
function applyStoredLocale(): void {
  const pref = loadSettings()?.language
  if (pref && pref !== 'auto' && (SUPPORTED_LOCALES as readonly string[]).includes(pref)) {
    setLocale(pref as Locale)
  }
}

function routeAfterAuth(): void {
  // Chat-first: after sign-in we always land on Home, which works with or
  // without a selected message. Email actions live inside Home's accordion
  // when a read item is loaded.
  go(isSignedIn.value ? 'home' : 'sign-in')
}

function bootstrap(): void {
  hydrateAuthState()
  applyStoredLocale()
  routeAfterAuth()
  // Keep useOutlookItem warm for child components.
  void useOutlookItem()

  window.addEventListener(AUTH_INVALIDATED_EVENT, () => {
    hydrateAuthState()
    currentView.value = 'sign-in'
  })

  mount()
}

/**
 * Run `bootstrap()` but never let an uncaught error leave the pane spinning.
 * A thrown error here (Safari-only API quirk, blocked storage, etc.) would
 * otherwise be invisible; surface it and log it for DevTools instead.
 */
function safeBootstrap(): void {
  try {
    bootstrap()
  } catch (err) {
    console.error('[Synamail] Boot failed:', err)
    const detail = err instanceof Error ? err.message : String(err)
    showBootError(`Synamail hit a problem while starting. ${SAFARI_TRACKING_HINT} (${detail})`)
  }
}

if (typeof Office !== 'undefined') {
  let officeReady = false
  Office.onReady(() => {
    officeReady = true
    safeBootstrap()
  })
  // Watchdog for the case where Office.onReady never fires — typically Safari
  // ITP stalling the Office.js runtime in the third-party taskpane iframe.
  window.setTimeout(() => {
    if (officeReady) return
    if (!document.getElementById('syn-boot')) return
    console.error('[Synamail] Office.onReady did not fire within the boot window.')
    showBootError(`Synamail could not finish loading inside Outlook. ${SAFARI_TRACKING_HINT}`)
  }, OFFICE_READY_TIMEOUT_MS)
} else {
  // Plain-browser dev (vite dev outside Outlook) — just mount.
  safeBootstrap()
}
