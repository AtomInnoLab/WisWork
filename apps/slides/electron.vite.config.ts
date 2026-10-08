import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const here = dirname(fileURLToPath(import.meta.url))

// Pin resolution to this repo's workspace sources (matches tsconfig paths;
// avoids bundling stale implementations when node_modules links point elsewhere)
const workspaceAlias = {
  // Subpath before the bare name: string aliases are prefix replacements
  '@wiswork/agent-core': resolve(here, '../../packages/agent-core/src/index.ts'),
  '@wiswork/presentation-ops': resolve(here, '../../packages/presentation-ops/src/index.ts'),
  '@wiswork/pptx-engine/table-grid': resolve(here, '../../packages/pptx-engine/src/table-grid.ts'),
  '@wiswork/pptx-engine/background-promote': resolve(
    here,
    '../../packages/pptx-engine/src/background-promote.ts',
  ),
  '@wiswork/pptx-render/preset-geometry': resolve(
    here,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@wiswork/pptx-render/coords': resolve(here, '../../packages/pptx-render/src/coords.ts'),
  '@wiswork/pptx-engine': resolve(here, '../../packages/pptx-engine/src/index.ts'),
  '@wiswork/pptx-render': resolve(here, '../../packages/pptx-render/src/index.ts'),
}

export default defineConfig({
  // Main process/preload must bundle @wiswork/* sources (they are pulled in as TS
  // source with extensionless relative imports; externalizing them under Node
  // yields ERR_MODULE_NOT_FOUND).
  main: {
    resolve: { alias: workspaceAlias },
    // Bundle opentype.js too (the packaged app ships only out/**, so external deps are unresolvable at runtime)
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@wiswork/auth',
          '@wiswork/pptx-engine',
          '@wiswork/pptx-render',
          '@wiswork/presentation-ops',
          '@wiswork/ai-search',
          '@wiswork/file-parse',
          '@wiswork/electron-utils',
          'opentype.js',
        ],
      }),
    ],
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
  },
  renderer: {
    define: {
      __WISWORK_SLIDES_ACCEPTANCE_E2E__: JSON.stringify(
        process.env.WISWORK_SLIDES_ACCEPTANCE_E2E === '1',
      ),
    },
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    server: {
      port: Number(process.env.SLIDES_DEV_PORT) || 5175,
      strictPort: Boolean(process.env.SLIDES_DEV_PORT),
    },
  },
})
