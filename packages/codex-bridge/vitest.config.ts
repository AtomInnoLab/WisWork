import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: {
      '@wiswork/agent-core/internal': fileURLToPath(
        new URL('../agent-core/src/internal.ts', import.meta.url),
      ),
      '@wiswork/agent-core': fileURLToPath(new URL('../agent-core/src/index.ts', import.meta.url)),
      '@wiswork/agent-runtime': fileURLToPath(
        new URL('../agent-runtime/src/index.ts', import.meta.url),
      ),
      '@wiswork/presentation-verification': fileURLToPath(
        new URL('../presentation-verification/src/index.ts', import.meta.url),
      ),
    },
  },
  test: { include: ['tests/**/*.test.ts'], environment: 'node' },
})
