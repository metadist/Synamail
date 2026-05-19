import { createI18n } from 'vue-i18n'
import en from '@/locales/en.json'
import de from '@/locales/de.json'

const fallbackLocale = 'en'

function detectLocale(): 'en' | 'de' {
  try {
    const display =
      typeof Office !== 'undefined' && Office.context?.displayLanguage
        ? Office.context.displayLanguage
        : (globalThis.navigator?.language ?? fallbackLocale)
    return display.toLowerCase().startsWith('de') ? 'de' : 'en'
  } catch {
    return fallbackLocale
  }
}

export const i18n = createI18n({
  legacy: false,
  locale: detectLocale(),
  fallbackLocale,
  messages: { en, de },
})

export function setLocale(loc: 'en' | 'de'): void {
  i18n.global.locale.value = loc
}
