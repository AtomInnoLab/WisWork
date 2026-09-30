import { expect, test } from '@playwright/test'

test.use({
  channel: 'chrome',
  ignoreHTTPSErrors: true,
  viewport: { width: 320, height: 900 },
})

test.beforeEach(async ({ page }) => {
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) =>
    route.fulfill({
      contentType: 'application/javascript',
      body: `
        const settings = new Map();
        window.Office = {
          onReady: async () => ({ host: 'PowerPoint' }),
          AsyncResultStatus: { Succeeded: 'succeeded' },
          CoercionType: { Text: 'text' },
          context: {
            host: 'PowerPoint',
            platform: 'PC',
            requirements: { isSetSupported: () => true },
            document: {
              url: 'https://example.test/deck.pptx',
              settings: {
                get: (key) => settings.get(key),
                set: (key, value) => settings.set(key, value),
                saveAsync: (callback) => callback({ status: 'succeeded' }),
              },
              getSelectedDataAsync: (_type, callback) => callback({ status: 'succeeded', value: '' }),
              setSelectedDataAsync: (_value, _options, callback) => callback({ status: 'succeeded' }),
            },
          },
        };
      `,
    }),
  )
  await page.goto('https://localhost:3000/taskpane.html')
})

test('renders the narrow PowerPoint pairing pane and accepts keyboard connection', async ({
  page,
}) => {
  await expect(page.getByRole('heading', { name: 'Connect to WisWork PC' })).toBeVisible()
  const connect = page.getByRole('button', { name: 'Connect to WisWork PC' })
  await expect(connect).toBeEnabled()
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(
    true,
  )
  await page.keyboard.press('Tab')
  await expect(connect).toBeFocused()
  await page.keyboard.press('Enter')
  await expect(page.getByRole('button', { name: /Looking for WisWork PC|Try again/ })).toBeVisible()
})
