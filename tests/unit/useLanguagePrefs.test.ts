import { describe, expect, it } from 'vitest'
import { saveSettings } from '@/taskpane/composables/useRoamingSettings'
import { standardLanguage, summaryLanguageOptions } from '@/taskpane/composables/useLanguagePrefs'

// The Office stub in tests/setup.ts reports `displayLanguage: 'en-US'`, so
// `detectLocale()` resolves to `en` whenever no explicit preference is stored.

const auth = { apiKey: 'sk_x', keyId: 1, email: 'a@b.test', baseUrl: 'https://x' }

describe('standardLanguage', () => {
  it('falls back to the Outlook display language when no preference is stored', () => {
    expect(standardLanguage()).toBe('en')
  })

  it('falls back to the Outlook display language when set to auto', async () => {
    await saveSettings({ ...auth, outputLanguage: 'auto' })
    expect(standardLanguage()).toBe('en')
  })

  it('uses the saved answer language when set explicitly', async () => {
    await saveSettings({ ...auth, outputLanguage: 'fr' })
    expect(standardLanguage()).toBe('fr')
  })

  it('falls back to the selected UI language when no answer language is set', async () => {
    await saveSettings({ ...auth, language: 'de' })
    expect(standardLanguage()).toBe('de')
  })

  it('prefers the explicit answer language over the UI language', async () => {
    await saveSettings({ ...auth, language: 'de', outputLanguage: 'fr' })
    expect(standardLanguage()).toBe('fr')
  })
})

describe('summaryLanguageOptions', () => {
  it('lists the selected language first, then English and German', () => {
    // Default answer language resolves to English (Outlook display language).
    expect(summaryLanguageOptions()).toEqual(['en', 'de'])
  })

  it('puts German first when it is the selected answer language', async () => {
    await saveSettings({ ...auth, outputLanguage: 'de' })
    expect(summaryLanguageOptions()).toEqual(['de', 'en'])
  })

  it('leads with the selected language, then English then German', async () => {
    await saveSettings({ ...auth, outputLanguage: 'fr' })
    expect(summaryLanguageOptions()).toEqual(['fr', 'en', 'de'])
  })

  it('keeps the selected UI language first when no answer language is set', async () => {
    await saveSettings({ ...auth, language: 'de' })
    expect(summaryLanguageOptions()).toEqual(['de', 'en'])
  })
})
