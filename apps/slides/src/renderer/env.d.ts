/// <reference types="vite/client" />
import type { SlidesApi } from '../shared/ipc'
import type { ProjectApi } from '@wiswork/project-store'
import type { PcHostCodexApi } from '@wiswork/agent-runtime'

declare global {
  interface Window {
    slidesApi: SlidesApi
    projectApi: ProjectApi
    codexRuntime: PcHostCodexApi
  }
}

export {}
