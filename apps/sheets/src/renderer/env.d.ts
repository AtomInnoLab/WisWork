declare module '*.md?raw' {
  const content: string
  export default content
}

import type { DesktopApi } from '../shared/desktop-api'
import type { ProjectApi } from '@wiswork/project-store'
import type { PcHostCodexApi } from '@wiswork/agent-runtime'

declare global {
  interface Window {
    readonly desktopApi: DesktopApi
    readonly projectApi: ProjectApi
    readonly codexRuntime: PcHostCodexApi
  }
}

export {}
