import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { App } from './App.js'
import { LatexLocaleProvider } from './i18n/locale.js'
import './styles.css'

const root = document.getElementById('root')
if (!root) throw new Error('Missing renderer root')

createRoot(root).render(
  <StrictMode>
    <LatexLocaleProvider>
      <App />
    </LatexLocaleProvider>
  </StrictMode>,
)
