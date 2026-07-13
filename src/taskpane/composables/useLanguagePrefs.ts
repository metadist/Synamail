/**
 * Resolve the user's "standard" language for AI output (email drafts,
 * summaries). It follows the roaming `language` preference, falling back to
 * the Outlook display language when the preference is `auto` or unset.
 *
 * This is distinct from the UI chrome locale — it's the default target
 * language the writing/summarize boxes generate in.
 */

import { detectLocale, type Locale } from '@/i18n'
import { loadSettings } from './useRoamingSettings'

/** The configured standard language code (e.g. `de`), or the Outlook locale. */
export function standardLanguage(): Locale {
  const pref = loadSettings()?.language
  return pref && pref !== 'auto' ? pref : detectLocale()
}

/**
 * Up to three summary-language options for the Summarize box: the configured
 * standard language first, then German and English when they aren't already
 * the standard (deduplicated, capped at three).
 */
export function summaryLanguageOptions(): Locale[] {
  const out: Locale[] = [standardLanguage()]
  for (const l of ['de', 'en'] as Locale[]) {
    if (!out.includes(l)) out.push(l)
  }
  return out.slice(0, 3)
}
