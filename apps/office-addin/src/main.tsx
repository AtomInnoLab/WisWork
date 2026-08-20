import { createRoot } from 'react-dom/client'
import { App } from './App.js'
import {
  authorizationFromDialogStart,
  forwardBrowserOAuthCallback,
} from './auth/office-dialog-auth.js'
import { loadRuntimeConfig } from './config.js'
import './styles.css'

const root = document.getElementById('root')
if (!root) throw new Error('Missing Office Add-in root')

const runtime = loadRuntimeConfig(import.meta.env, { production: import.meta.env.PROD })
const callback = runtime.status === 'available' ? new URL(runtime.config.callbackUrl) : undefined
const dialogAuthorization =
  runtime.status === 'available'
    ? authorizationFromDialogStart(window.location.href, runtime.config.authorizationUrl)
    : undefined
const isOAuthCallback =
  callback !== undefined &&
  window.location.origin === callback.origin &&
  window.location.pathname === callback.pathname

if (dialogAuthorization) {
  window.location.replace(dialogAuthorization)
} else if (isOAuthCallback && runtime.status === 'available') {
  void Office.onReady()
    .then(() => {
      if (!forwardBrowserOAuthCallback(window.location.href, runtime.config.callbackUrl)) {
        createRoot(root).render(<App />)
      }
    })
    .catch(() => createRoot(root).render(<App />))
} else {
  createRoot(root).render(<App />)
}
