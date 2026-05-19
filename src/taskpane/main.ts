/**
 * Taskpane entry point. Waits for Office.onReady, hydrates auth state, then
 * mounts the Vue app. Listens for the AUTH_INVALIDATED_EVENT to bounce the
 * user back to SignIn on 401.
 */

import { createApp } from 'vue'
import App from './App.vue'
import { i18n } from '@/i18n'
import { hydrateAuthState, isSignedIn } from './composables/useAuth'
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

function routeFromItem(): void {
  if (!isSignedIn.value) {
    go('sign-in')
    return
  }
  if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
    return
  }
  // Use the composable's snapshot detection to decide read vs compose. We
  // can't await composable code here, so do a quick local check on the body
  // surface to set the initial view immediately.
  const body = Office.context.mailbox.item.body as
    | { setAsync?: unknown; getAsync?: unknown }
    | undefined
  if (body && typeof (body as { setAsync?: unknown }).setAsync === 'function') {
    go('compose')
  } else if (body && typeof (body as { getAsync?: unknown }).getAsync === 'function') {
    go('read')
  }
}

function bootstrap(): void {
  hydrateAuthState()
  routeFromItem()
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
