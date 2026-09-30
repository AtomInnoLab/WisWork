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
        settings.set('wiswork.presentation.project.v1', 'project-1');
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
})

test('renders the narrow PowerPoint pairing pane and accepts keyboard connection', async ({
  page,
}) => {
  await page.goto('https://localhost:3000/taskpane.html')
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

test('reopens the paired project workbench with its saved project identity', async ({ page }) => {
  const project = {
    projectId: 'project-1',
    createdAt: '2026-09-28T00:00:00.000Z',
    title: '浏览器项目验收',
    status: 'compiled',
    latestRequestId: 'request-1',
    latestCompiledRequestId: 'request-1',
    slideCount: 1,
    slides: [{ id: 'page-1', title: '本机验证页' }],
    history: [{ requestId: 'request-1', sequence: 1, status: 'compiled', slideCount: 1 }],
    checks: {
      structure: 'passed',
      geometry: 'passed',
      render: 'not_run',
      sources: 'not_verified',
      roundTrip: 'not_run',
    },
  }
  await page.addInitScript((project) => {
    const NativeWebSocket = window.WebSocket
    let dropCurrentRelay: (() => void) | undefined
    let grantedCapabilities: string[] = []
    let resumeCount = 0
    class RelaySocket {
      readyState = 0
      onopen: (() => void) | null = null
      onmessage: ((event: { data: string }) => void) | null = null
      onclose: (() => void) | null = null
      onerror: (() => void) | null = null
      constructor() {
        dropCurrentRelay = () => this.close()
        setTimeout(() => {
          this.readyState = 1
          this.onopen?.()
        }, 0)
      }
      private receive(value: Record<string, unknown>) {
        queueMicrotask(() => this.onmessage?.({ data: JSON.stringify(value) }))
      }
      send(raw: string) {
        const frame = JSON.parse(raw) as Record<string, unknown>
        if (frame.type === 'office.create') {
          grantedCapabilities = frame.capabilities as string[]
          this.receive({
            version: 2,
            type: 'office.created',
            pairing_id: 'pairing1',
            verification_code: '123456',
            expires_in: 120,
          })
          this.receive({
            version: 2,
            type: 'office.approved',
            session_id: 'session1',
            capability: 'token1',
            expires_in: 1800,
            capabilities: frame.capabilities,
          })
        }
        if (frame.type === 'office.resume') {
          resumeCount++
          this.receive({
            version: 2,
            type: 'office.resumed',
            session_id: 'session1',
            expires_in: 1800,
            capabilities: grantedCapabilities,
          })
        }
        if (frame.type !== 'office.request') return
        const request = frame.body as { operation?: string }
        const body = JSON.stringify(
          request.operation === 'status' ? project : { error: 'not_found' },
        )
        const data = btoa(String.fromCharCode(...new TextEncoder().encode(body)))
        const common = { version: 2, session_id: frame.session_id, request_id: frame.request_id }
        this.receive({
          ...common,
          type: 'relay.start',
          status: 200,
          content_type: 'application/json',
        })
        this.receive({ ...common, type: 'relay.chunk', sequence: 0, data })
        this.receive({ ...common, type: 'relay.done' })
      }
      close() {
        this.readyState = 3
        this.onclose?.()
      }
    }
    function MockWebSocket(url: string, protocols?: string | string[]) {
      return url.includes('/office-relay') ? new RelaySocket() : new NativeWebSocket(url, protocols)
    }
    Object.defineProperty(window, 'WebSocket', { value: MockWebSocket })
    Object.defineProperty(window, '__wisworkRelayTest', {
      value: { drop: () => dropCurrentRelay?.(), resumes: () => resumeCount },
    })
  }, project)
  await page.goto('https://localhost:3000/taskpane.html')
  await page.getByRole('button', { name: 'Connect to WisWork PC' }).click()
  const workbench = page.getByRole('region', { name: '演示文稿项目' })
  await expect(workbench).toBeVisible()
  await expect(workbench.getByText('浏览器项目验收', { exact: true })).toBeVisible()
  await expect(workbench).toContainText('尚未完成视觉验证')
  await expect(page.getByRole('region', { name: '演示文稿制作阶段' })).toContainText('交付核验')
  await page.getByText('恢复记录 · 1 项').click()
  await expect(page.getByText(/项目生命周期已登记：浏览器项目验收/)).toBeVisible()
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(
    true,
  )
  const refresh = workbench.getByRole('button', { name: '刷新', exact: true })
  await refresh.focus()
  await page.keyboard.press('Enter')
  await expect(workbench.getByText('浏览器项目验收', { exact: true })).toBeVisible()
  await page.evaluate(() =>
    (window as unknown as { __wisworkRelayTest: { drop(): void } }).__wisworkRelayTest.drop(),
  )
  await expect
    .poll(() =>
      page.evaluate(() =>
        (
          window as unknown as { __wisworkRelayTest: { resumes(): number } }
        ).__wisworkRelayTest.resumes(),
      ),
    )
    .toBe(1)
  await expect(workbench).toContainText('浏览器项目验收')
  await page.reload()
  await page.getByRole('button', { name: 'Connect to WisWork PC' }).click()
  await expect(page.getByRole('region', { name: '演示文稿项目' })).toContainText('浏览器项目验收')
  await expect(page.getByText('恢复记录 · 1 项')).toBeVisible()
})
