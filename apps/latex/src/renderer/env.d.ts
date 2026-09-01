/// <reference types="vite/client" />

import type { LatexApi } from '../shared/ipc.js'
import type { PcHostCodexApi } from '@wiswork/agent-runtime'

declare global {
  interface Window {
    latexApi: LatexApi
    codexRuntime: PcHostCodexApi
  }
}

export {}
