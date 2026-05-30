import { createI18n } from 'vue-i18n'
import en from '@/locales/en.json'
import de from '@/locales/de.json'
import fr from '@/locales/fr.json'
import es from '@/locales/es.json'
import it from '@/locales/it.json'
import pt from '@/locales/pt.json'

// The UI locales Synamail ships. English is the fallback for any Outlook
// display language not in this list. Add a locale here AND a matching
// `src/locales/<code>.json` to expand coverage.
export const SUPPORTED_LOCALES = ['en', 'de', 'fr', 'es', 'it', 'pt'] as const
export type Locale = (typeof SUPPORTED_LOCALES)[number]

const fallbackLocale: Locale = 'en'

const messages: Record<Locale, typeof en> = { en, de, fr, es, it, pt }

/**
 * Decide the UI locale from the Outlook display language (or the browser
 * language outside Office), matching on the primary subtag only:
 *   `de-AT`, `de-CH` → `de`;  `pt-BR`, `pt-PT` → `pt`;  `fr-CA` → `fr`.
 * Anything we don't ship (e.g. `zh`, `ja`, `nl`) falls back to English.
 */
export function detectLocale(): Locale {
  try {
    const display =
      typeof Office !== 'undefined' && Office.context?.displayLanguage
        ? Office.context.displayLanguage
        : (globalThis.navigator?.language ?? fallbackLocale)
    const primary = display.toLowerCase().split('-')[0]
    return (SUPPORTED_LOCALES as readonly string[]).includes(primary)
      ? (primary as Locale)
      : fallbackLocale
  } catch {
    return fallbackLocale
  }
}

export const i18n = createI18n({
  legacy: false,
  locale: detectLocale(),
  fallbackLocale,
  messages,
})

export function setLocale(loc: Locale): void {
  i18n.global.locale.value = loc
}
