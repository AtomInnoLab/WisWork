/// <reference types="vite/client" />

import type { LatexApi } from '../shared/ipc.js'

declare global {
  interface Window {
    latexApi: LatexApi
  }
}

export {}
