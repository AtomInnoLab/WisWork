import { defineConfig } from '@playwright/test'

/** Browser-level Taskpane checks with a local Office.js stand-in. */
for (const key of ['NO_PROXY', 'no_proxy'])
  process.env[key] = [process.env[key], 'localhost', '127.0.0.1'].filter(Boolean).join(',')

export default defineConfig({
  testDir: '.',
  testMatch: 'office-taskpane.spec.ts',
  outputDir: '../test-results/office',
  timeout: 45_000,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  webServer: {
    command: 'npm run dev -w @wiswork/office-addin -- --host 127.0.0.1',
    url: 'https://localhost:3000/taskpane.html',
    ignoreHTTPSErrors: true,
    reuseExistingServer: false,
    timeout: 60_000,
  },
})
