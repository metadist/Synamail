import { describe, expect, it } from 'vitest'
import { createI18n } from 'vue-i18n'
import de from '@/locales/de.json'
import en from '@/locales/en.json'
import es from '@/locales/es.json'
import fr from '@/locales/fr.json'
import itLocale from '@/locales/it.json'
import pt from '@/locales/pt.json'

// Every shipped locale. Keep in sync with SUPPORTED_LOCALES in src/i18n.ts.
const locales: Record<string, Record<string, unknown>> = { en, de, fr, es, it: itLocale, pt }

/** All dotted leaf-key paths of a (possibly nested) messages object. */
function leafKeys(obj: Record<string, unknown>, prefix = ''): string[] {
  const keys: string[] = []
  for (const [k, v] of Object.entries(obj)) {
    const path = prefix ? `${prefix}.${k}` : k
    if (v && typeof v === 'object') keys.push(...leafKeys(v as Record<string, unknown>, path))
    else keys.push(path)
  }
  return keys
}

// vue-i18n compiles each message lazily on first t() access and THROWS a
// SyntaxError on invalid message syntax — most infamously a literal "@"
// (linked-message syntax) or unescaped "{"/"}". A malformed string only blows
// up when the component that renders it mounts, so it slips past a build: this
// is exactly what made the "Save to knowledge base" dialog fail to open
// (newGroupPlaceholder contained "alice@example.com"). Compiling every message
// here turns that runtime-only landmine into a fast, deterministic unit test.
describe('i18n messages compile for every locale', () => {
  for (const [loc, messages] of Object.entries(locales)) {
    it(`compiles every "${loc}" message without throwing`, () => {
      const i18n = createI18n({ legacy: false, locale: loc, messages: { [loc]: messages } })
      const t = i18n.global.t as (key: string) => string
      for (const key of leafKeys(messages)) {
        expect(() => t(key), `message key "${key}" in locale "${loc}" must compile`).not.toThrow()
      }
    })
  }

  it('exposes the same key set in all locales (no missing translations)', () => {
    const reference = leafKeys(en).sort()
    for (const [loc, messages] of Object.entries(locales)) {
      expect(leafKeys(messages).sort(), `locale "${loc}" key set`).toEqual(reference)
    }
  })
})
