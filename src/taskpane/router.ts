/**
 * Tiny reactive view store. We don't need vue-router for a taskpane — the
 * navigation state is a single string ref.
 */

import { ref } from 'vue'

export type ViewName =
  | 'sign-in'
  | 'home'
  | 'read'
  | 'compose'
  | 'settings'
  | 'rule-editor'
  | 'contact-kb'

export const currentView = ref<ViewName>('sign-in')
const history: ViewName[] = []

export function go(view: ViewName): void {
  if (currentView.value !== view) {
    history.push(currentView.value)
    currentView.value = view
  }
}

export function back(fallback: ViewName = 'settings'): void {
  const prev = history.pop()
  currentView.value = prev ?? fallback
}

export function resetHistory(): void {
  history.length = 0
}
