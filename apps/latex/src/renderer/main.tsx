import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { App } from './App.js'
import { LatexLocaleProvider } from './i18n/locale.js'
import type { UiTheme } from '../shared/ipc.js'
import '@wiswork/ui/tokens.css'
import { EnhancedMutationConfirmation } from '@wiswork/ui'
import { normalizeLang } from '@wiswork/i18n'
import './styles.css'

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

const root = document.getElementById('root')
if (!root) throw new Error('Missing renderer root')

void (async () => {
  const confirmationLocale = normalizeLang(navigator.language)
  applyTheme(await window.latexApi.getTheme().catch(() => 'system' as const))
  window.latexApi.onThemeChanged(applyTheme)
  createRoot(root).render(
    <StrictMode>
      <LatexLocaleProvider>
        <App />
        <EnhancedMutationConfirmation api={window.codexRuntime} locale={confirmationLocale} />
      </LatexLocaleProvider>
    </StrictMode>,
  )
})()
