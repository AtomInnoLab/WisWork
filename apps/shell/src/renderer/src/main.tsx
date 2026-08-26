import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang } from '@wiswork/i18n'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import '@wiswork/ui/tokens.css'
import '@wiswork/ui/screentip.css'
import './home.css'
import './tabbar.css'
import { installScreenTips } from '@wiswork/ui'

installScreenTips()

// macOS shell window is created with vibrancy; a transparent body lets the
// editor views' translucent regions (e.g. slides thumbnail pane) show it
if (navigator.platform.toLowerCase().includes('mac')) document.body.classList.add('vib')

// resolve the persisted language, first-run flag, and theme before first paint
// so the UI never flashes (home showing briefly before the onboarding overlay)
void Promise.all([
  window.aiOffice.getLanguage(),
  window.aiOffice.getTheme(),
  // if the flag is unreadable, skip onboarding rather than block the home screen
  window.aiOffice.onboardingSeen().catch(() => true),
]).then(([lang, theme, onboardingSeen]) => {
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dataset.theme = theme
  window.aiOffice.onThemeChanged((nextTheme) => {
    document.documentElement.dataset.theme = nextTheme
  })
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        <AppFrame initialOnboardingSeen={onboardingSeen} />
      </LocaleProvider>
    </React.StrictMode>,
  )
})
