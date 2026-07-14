/**
 * Resolve the user's "standard" language for AI output (email drafts,
 * summaries, translations). Precedence:
 *   1. the explicit `outputLanguage` (answer language) preference,
 *   2. otherwise the explicit UI `language` the user picked,
 *   3. otherwise the Outlook display language.
 *
 * Falling back to the UI language (step 2) keeps a user's selected language as
 * the primary Summarize button even when they haven't set the dedicated answer
 * language — only `auto`/unset on both drops through to auto-detect.
 */

import { detectLocale, type Locale } from '@/i18n'
import { loadSettings } from './useRoamingSettings'

/** The configured answer language code (e.g. `de`), or the Outlook locale. */
export function standardLanguage(): Locale {
  const s = loadSettings()
  for (const pref of [s?.outputLanguage, s?.language]) {
    if (pref && pref !== 'auto') return pref
  }
  return detectLocale()
}

/**
 * Up to three summary-language options for the Summarize box: the selected
 * answer language first, then English and German when they aren't already the
 * selected one (deduplicated, capped at three).
 */
export function summaryLanguageOptions(): Locale[] {
  const out: Locale[] = [standardLanguage()]
  for (const l of ['en', 'de'] as Locale[]) {
    if (!out.includes(l)) out.push(l)
  }
  return out.slice(0, 3)
}
