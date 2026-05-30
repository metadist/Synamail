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
  // without a selected message. The email-specific Read/Compose views are
  // reachable from Home's contextual link when an item is loaded.
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

if (typeof Office !== 'undefined') {
  Office.onReady(bootstrap)
} else {
  // Plain-browser dev (vite dev outside Outlook) — just mount.
  bootstrap()
}
