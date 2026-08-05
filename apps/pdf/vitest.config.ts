import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    setupFiles: ['../../tools/vitest-jsdom-setup.ts'],
    testTimeout: 20000,
  },
})
