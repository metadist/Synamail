/**
 * Tiny reactive view store. We don't need vue-router for a taskpane — the
 * navigation state is a single string ref.
 */

import { ref } from 'vue'

export type ViewName = 'sign-in' | 'home' | 'read' | 'compose' | 'settings' | 'contact-profile'

export const currentView = ref<ViewName>('sign-in')
const history: ViewName[] = []

/**
 * The contact email the ContactProfile view is scoped to. Set via
 * `openContactProfile(email)` before navigating, so the view knows which
 * `contact:<email>` RAG group to search / save into. Null = no contact picked.
 */
export const selectedContactEmail = ref<string | null>(null)

export function openContactProfile(email: string): void {
  selectedContactEmail.value = email.trim().toLowerCase() || null
  go('contact-profile')
}

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
