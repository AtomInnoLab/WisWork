import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

const here = dirname(fileURLToPath(import.meta.url))

export default defineConfig({
  // Always resolve to this repo's workspace sources (consistent with tsconfig paths)
  resolve: {
    alias: Object.entries({
      '@wiswork/pptx-engine/table-grid': resolve(here, '../pptx-engine/src/table-grid.ts'),
      '@wiswork/pptx-engine/background-promote': resolve(
        here,
        '../pptx-engine/src/background-promote.ts',
      ),
      '@wiswork/pptx-engine': resolve(here, '../pptx-engine/src/index.ts'),
    }).map(([name, replacement]) => ({
      // Keep bare-package aliases from capturing exported subpaths.
      find: new RegExp(`^${name}$`),
      replacement,
    })),
  },
  test: {
    include: ['tests/**/*.test.ts'],
  },
})
