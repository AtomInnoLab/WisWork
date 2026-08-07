import { createContext, useContext, useMemo } from 'react'
import type { PropsWithChildren } from 'react'
import {
  LATEX_STRINGS,
  resolveLatexLocale,
  type LatexLocale,
  type LatexStringKey,
} from './strings.js'

interface LocaleValue {
  locale: LatexLocale
  t: (key: LatexStringKey) => string
}

const LocaleContext = createContext<LocaleValue>({
  locale: 'en',
  t: (key) => LATEX_STRINGS.en[key],
})

export function LatexLocaleProvider({ children }: PropsWithChildren) {
  const value = useMemo(() => {
    const locale = resolveLatexLocale(navigator.language)
    return { locale, t: (key: LatexStringKey) => LATEX_STRINGS[locale][key] }
  }, [])
  return <LocaleContext.Provider value={value}>{children}</LocaleContext.Provider>
}

export function useLatexLocale(): LocaleValue {
  return useContext(LocaleContext)
}
