/// <reference types="vite/client" />

import type { LatexApi } from '../shared/ipc.js'
import type { CodexRuntimeApi, CodexToolApi } from '../../../shell/src/shared/codex-api.js'

declare global {
  interface Window {
    latexApi: LatexApi
    wisworkCodexTools: CodexToolApi
    wisworkCodexRuntime: CodexRuntimeApi
  }
}

export {}
