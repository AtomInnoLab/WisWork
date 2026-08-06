import { describe, expect, it } from 'vitest'
import { LATEX_STRINGS, SUPPORTED_LATEX_LOCALES } from '../src/renderer/i18n/strings.js'

describe('LaTeX renderer strings', () => {
  it('provides every key for every supported locale', () => {
    const englishKeys = Object.keys(LATEX_STRINGS.en).sort()
    expect(SUPPORTED_LATEX_LOCALES.length).toBeGreaterThanOrEqual(16)
    for (const locale of SUPPORTED_LATEX_LOCALES) {
      expect(Object.keys(LATEX_STRINGS[locale]).sort()).toEqual(englishKeys)
      expect(Object.values(LATEX_STRINGS[locale]).every((value) => value.trim().length > 0)).toBe(
        true,
      )
    }
  })
})
