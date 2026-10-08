/// <reference types="vite/client" />

import type { DesktopApi } from '../shared/ipc'
import type { ProjectApi } from '@wiswork/project-store'
import type { PcHostCodexApi } from '@wiswork/agent-runtime'

declare global {
  interface Window {
    desktop: DesktopApi
    projectApi: ProjectApi
    codexRuntime: PcHostCodexApi
  }
}

export {}
