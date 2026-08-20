import { createRoot } from 'react-dom/client'
import { App } from './App.js'
import {
  authorizationFromDialogStart,
} from './auth/office-dialog-auth.js'
import { loadRuntimeConfig } from './config.js'
import { captureAndScrubOAuthCallback } from './auth/oauth-callback.js'
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
const capturedCallback =
  isOAuthCallback && runtime.status === 'available'
    ? captureAndScrubOAuthCallback(
        runtime.config.callbackUrl,
        window.location.href,
        (cleanUrl) => window.history.replaceState({}, '', cleanUrl),
      )
    : undefined

if (dialogAuthorization) {
  window.location.replace(dialogAuthorization)
} else if (isOAuthCallback && runtime.status === 'available') {
  createRoot(root).render(<App oauthCallback={capturedCallback} />)
} else {
  createRoot(root).render(<App />)
}
