/* global window, document */
import { spawn } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'
import electron from 'electron'
import { chromium } from '@playwright/test'
import WebSocket from 'ws'
import { inspectPcBusiness, releaseProductionFixture } from './ppt-agent-pc-business-smoke.mjs'
import { loadBenchmarkBundle } from './ppt-agent-benchmark-bundle.mjs'
import { inspectOfficeBuild } from './ppt-agent-release-preflight.mjs'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const children = []
const documentId = 'electron-real-relay-document'
const projectId = 'electron-real-relay-project'
const browserProjectId = 'electron-browser-relay-project'
const browserDocumentUrl = 'https://example.test/browser-relay-deck.pptx'
const browserDocumentId = JSON.stringify(['electron-browser-relay-document', browserDocumentUrl])
const concurrentDocuments = [1, 2, 3].map((index) => ({
  documentId: `electron-concurrent-document-${index}`,
  projectId: `electron-concurrent-project-${index}`,
  text: `Concurrent document ${index}`,
}))
const expectedSlideTexts = Array.from(
  { length: 8 },
  (_, index) => `Electron real Relay page ${index + 1}`,
)
const selectedBenchmark = process.argv.find((arg) => arg.startsWith('--benchmark='))
const benchmarkBatch = process.argv.includes('--benchmark-batch')
const p014Fallback = process.argv.includes('--p0-14-fallback')
if ([benchmarkBatch, Boolean(selectedBenchmark), p014Fallback].filter(Boolean).length > 1)
  throw new Error('choose one benchmark mode')
const fallbackUrls = [
  'https://93.184.216.34/2024GISTEMPMap-timeout.png',
  'https://93.184.216.34/2024GISTEMPMap_2K.png',
]
const fallbackImagePath = join(
  root,
  'docs/product/ppt-benchmark-materials/PPT-P0-14/images/nasa-2024-temperature-anomaly-2k.png',
)
const fallbackImage = p014Fallback ? await readFile(fallbackImagePath) : undefined
const fallbackImageSha = fallbackImage
  ? createHash('sha256').update(fallbackImage).digest('hex')
  : undefined
const batchCases = ['P0-05', 'P0-06', 'P0-07', 'P0-08', 'P0-15', 'P0-16', 'P0-18']
const builtTaskpane = process.argv.includes('--built-taskpane')
if (builtTaskpane) {
  const dist = join(root, 'apps/office-addin/dist')
  const manifest = await readFile(join(dist, 'manifest.xml'), 'utf8')
  const origin = /<AppDomain>([^<]+)<\/AppDomain>/.exec(manifest)?.[1]
  if (!origin) throw new Error('built Taskpane manifest origin missing')
  await inspectOfficeBuild(dist, origin)
}
const [benchmarkCase, selectedVariant] = selectedBenchmark
  ? selectedBenchmark.slice('--benchmark='.length).split(':')
  : ['P0-01']
const concurrentBenchmark = benchmarkCase === 'P0-17' && selectedVariant === 'all'
const sourceBackedCrash = benchmarkCase === 'P0-20'
const benchmarkVariant = concurrentBenchmark ? 'science' : selectedVariant
const researchCaseId = `PPT-${benchmarkCase}`
const {
  plan: researchPlan,
  deck: researchDeck,
  sourceAttachments: researchSources,
} = await loadBenchmarkBundle(root, researchCaseId, benchmarkVariant)
const derivedPageFixture =
  researchCaseId === 'PPT-P0-18'
    ? {
        requestId: 'P0-18-revised-p04',
        pageId: 'p04',
        slide: JSON.parse(
          await readFile(
            join(root, 'docs/product/ppt-benchmark-materials/PPT-P0-18/revised-page-deck.json'),
            'utf8',
          ),
        ).slides[0],
      }
    : undefined
const temp = await mkdtemp(join(tmpdir(), 'ppt-electron-real-relay-'))
const deck = {
  version: 1,
  id: projectId,
  title: 'Electron real Relay smoke',
  style: { fontFace: 'Arial', background: 'FFFFFF', textColor: '111111', accentColor: '3366FF' },
  assets: [],
  claims: [],
  slides: expectedSlideTexts.map((value, index) => ({
    id: `slide-${index + 1}`,
    title: `Page ${index + 1}`,
    elements: [{ id: 'title', kind: 'text', text: value, x: 1, y: 1, w: 8, h: 1 }],
  })),
}
const plan = {
  version: 1,
  projectId,
  title: deck.title,
  brief: {
    objective: 'Verify an eight-page editable deck',
    audience: 'Release test',
    language: 'en-US',
    minutes: 8,
    requiredContent: [],
    constraints: [],
  },
  sources: [],
  claims: [],
  style: deck.style,
  slides: deck.slides.map((slide) => ({
    id: slide.id,
    title: slide.title,
    purpose: 'Verify native text',
    claimIds: [],
    layout: 'content',
    requiredAssets: [],
    acceptanceCriteria: ['Native text remains editable'],
  })),
}
let smokeStage = 'setup'

async function inspectBrowserWorkbench(origin, pc) {
  const dev = spawn(
    builtTaskpane ? 'npx' : 'npm',
    builtTaskpane
      ? [
          '--no-install',
          'vite',
          'preview',
          '--config',
          'apps/office-addin/vite.config.ts',
          '--host',
          '127.0.0.1',
          '--port',
          '3000',
          '--strictPort',
        ]
      : ['run', 'dev', '-w', '@wiswork/office-addin', '--', '--host', '127.0.0.1'],
    {
      cwd: root,
      stdio: ['ignore', 'pipe', 'inherit'],
      detached: process.platform === 'linux',
    },
  )
  children.push(dev)
  await firstLine(dev, 'Office Taskpane', 60_000, (line) =>
    line.includes('https://localhost:3000/'),
  )
  const browser = await chromium.launch({ channel: 'chrome', headless: true })
  const sockets = new Map()
  const outbound = []
  const inbound = []
  try {
    const page = await browser.newPage({
      ignoreHTTPSErrors: true,
      viewport: { width: 320, height: 900 },
    })
    const dispatch = async (id, type, data) => {
      if (page.isClosed()) return
      await page
        .evaluate(
          ([socketId, eventType, eventData]) => {
            window.__relayDispatch(socketId, eventType, eventData)
          },
          [id, type, data],
        )
        .catch(() => {})
    }
    await page.exposeFunction('__relayOpen', (id) => {
      const socket = new WebSocket(origin.replace('http:', 'ws:') + '/office-relay', {
        headers: { Origin: 'https://office.8-216-134-194.sslip.io' },
      })
      sockets.set(id, socket)
      socket.on('open', () => void dispatch(id, 'open'))
      socket.on('message', (data) => {
        inbound.push(JSON.parse(data.toString()).type)
        void dispatch(id, 'message', data.toString())
      })
      socket.on('close', () => void dispatch(id, 'close'))
      socket.on('error', () => void dispatch(id, 'error'))
    })
    await page.exposeFunction('__relaySend', (id, data) => {
      outbound.push(JSON.parse(data).type)
      sockets.get(id)?.send(data)
    })
    await page.exposeFunction('__relayClose', (id) => sockets.get(id)?.close())
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) =>
      route.fulfill({
        contentType: 'application/javascript',
        body: `const settings = new Map([
          ['wiswork.presentation.document.v1', 'electron-browser-relay-document'],
          ['wiswork.presentation.project.v1', '${browserProjectId}']
        ]);
        window.__documentUrl = '${browserDocumentUrl}';
        window.Office = { onReady: async () => ({ host: 'PowerPoint' }),
          AsyncResultStatus: { Succeeded: 'succeeded' }, CoercionType: { Text: 'text' },
          FileType: { Compressed: 'compressed' },
          context: { host: 'PowerPoint', platform: 'PC', requirements: { isSetSupported: () => true },
            document: { get url() { return window.__documentUrl },
              settings: { get: key => settings.get(key), set: (key, value) => settings.set(key, value),
                saveAsync: callback => callback({ status: 'succeeded' }) },
              getFileAsync: (_type, _options, callback) => callback({ status: 'succeeded', value: {
                size: 5, sliceCount: 1,
                getSliceAsync: (index, done) => done({ status: 'succeeded', value: {
                  index, size: 5, data: [80, 75, 3, 4, 1]
                } }),
                closeAsync: done => done({ status: 'succeeded' })
              } }),
              getSelectedDataAsync: (_type, callback) => callback({ status: 'succeeded', value: '' }),
              setSelectedDataAsync: (_value, _options, callback) => callback({ status: 'succeeded' })
            } } };
        window.PowerPoint = { createPresentation: async base64 => {
          window.__createdPresentation = base64;
          window.__createdPresentationCount = (window.__createdPresentationCount || 0) + 1
        } };`,
      }),
    )
    if (builtTaskpane) {
      let failedVersionChecks = 0
      await page.route('https://localhost:3000/version.json', (route) => {
        if (failedVersionChecks++ === 0)
          return route.fulfill({
            status: 200,
            contentType: 'application/json',
            body: JSON.stringify({
              buildId: 'smoke_release',
              presentationMinPcProtocol: 1,
              presentationMinRelayProtocol: 2,
            }),
          })
        return route.continue()
      })
    }
    await page.addInitScript(() => {
      const sockets = new Map()
      window.__relayDispatch = (id, type, data) => {
        const socket = sockets.get(id)
        if (!socket) return
        if (type === 'open') socket.readyState = 1
        if (type === 'close') socket.readyState = 3
        socket[`on${type}`]?.(type === 'message' ? { data } : {})
      }
      window.WebSocket = class {
        constructor(url) {
          if (!url.includes('/office-relay')) throw Error('unexpected WebSocket endpoint')
          this.id = crypto.randomUUID()
          this.readyState = 0
          sockets.set(this.id, this)
          void window.__relayOpen(this.id)
        }
        send(data) {
          void window.__relaySend(this.id, data)
        }
        close() {
          void window.__relayClose(this.id)
        }
      }
    })
    const pageErrors = []
    page.on('pageerror', (error) => pageErrors.push(error.message.slice(0, 300)))
    const requestFailures = []
    page.on('requestfailed', (request) =>
      requestFailures.push(`${request.url()}: ${request.failure()?.errorText}`),
    )
    const failedResponses = []
    page.on('response', (received) => {
      if (received.status() >= 400) failedResponses.push(`${received.status()} ${received.url()}`)
    })
    const response = await page.goto('https://localhost:3000/taskpane.html')
    if (builtTaskpane) {
      await page.getByRole('heading', { name: 'Cannot verify WisWork version' }).waitFor()
      if (await page.getByRole('button', { name: 'Connect to WisWork PC' }).count())
        throw Error('unverified Taskpane version exposed PC connection')
      await page.getByRole('button', { name: 'Retry version check' }).click()
    }
    try {
      await page.getByRole('button', { name: 'Connect to WisWork PC' }).click()
    } catch (error) {
      const state = await page.evaluate(() => ({
        readyState: document.readyState,
        body: document.body?.innerText.slice(0, 500),
        html: document.documentElement?.outerHTML.slice(0, 800),
      }))
      throw new Error(
        `Taskpane connect button unavailable: ${JSON.stringify({ status: response?.status(), url: page.url(), ...state, pageErrors, requestFailures, failedResponses })}`,
        { cause: error },
      )
    }
    const codeText = await page
      .getByText(/Enter code [0-9]{6} in WisWork PC/)
      .textContent({ timeout: 15_000 })
    const code = codeText?.match(/Enter code ([0-9]{6})/)?.[1]
    if (!code) throw Error('browser pairing code missing')
    pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n')
    const workbench = page.getByRole('region', { name: '演示文稿项目' })
    await workbench
      .getByText('Browser to real PC project', { exact: true })
      .waitFor({ timeout: 20_000 })
    if (!(await workbench.textContent()).includes('尚未完成视觉验证'))
      throw Error('browser project verification state missing')
    await page.getByRole('button', { name: '创建副本后制作' }).click()
    await page.getByText(/副本已打开。请先另存为新文件/).waitFor()
    if ((await page.evaluate(() => window.__createdPresentation)) !== 'UEsDBAE=')
      throw Error('browser presentation copy did not receive the exported source')
    await page.evaluate(() => {
      window.__documentUrl = ''
    })
    await page.getByRole('button', { name: '创建副本后制作' }).click()
    await page.getByText(/请先保存当前文档，再创建副本/).waitFor()
    if ((await page.evaluate(() => window.__createdPresentationCount)) !== 1)
      throw Error('browser presentation copy opened an unsaved source')
    await page.evaluate((url) => {
      window.__documentUrl = url
    }, browserDocumentUrl)
    if (await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth))
      throw Error('browser project workbench overflows at 320px')
    const firstSocket = sockets.values().next().value
    if (!firstSocket || !inbound.includes('office.approved'))
      throw Error('browser relay approval not observed')
    firstSocket.terminate()
    const waitFor = async (check, label) => {
      for (let attempt = 0; attempt < 100; attempt++) {
        if (check()) return
        await new Promise((resolveWait) => setTimeout(resolveWait, 100))
      }
      throw Error(`browser relay ${label} timed out`)
    }
    await waitFor(() => outbound.includes('office.resume'), 'resume request')
    await waitFor(() => inbound.includes('office.resumed'), 'resume approval')
    const previousRequests = outbound.filter((type) => type === 'office.request').length
    const previousResponses = inbound.filter((type) => type === 'relay.done').length
    await workbench.getByRole('button', { name: '刷新', exact: true }).click()
    await waitFor(
      () => outbound.filter((type) => type === 'office.request').length > previousRequests,
      'project refresh',
    )
    await waitFor(
      () => inbound.filter((type) => type === 'relay.done').length > previousResponses,
      'project refresh response',
    )
    await workbench.getByText('Browser to real PC project', { exact: true }).waitFor()
    const createdBeforeReload = outbound.filter((type) => type === 'office.create').length
    await page.reload()
    try {
      await page.getByRole('button', { name: 'Connect to WisWork PC' }).click()
    } catch (error) {
      const state = await page.evaluate(() => ({
        readyState: document.readyState,
        body: document.body?.innerText.slice(0, 700),
      }))
      throw new Error(
        `Taskpane reconnect button unavailable: ${JSON.stringify({ url: page.url(), ...state, pageErrors, outboundTail: outbound.slice(-8), inboundTail: inbound.slice(-8) })}`,
        { cause: error },
      )
    }
    const reopenedCodeText = await page
      .getByText(/Enter code [0-9]{6} in WisWork PC/)
      .textContent({ timeout: 15_000 })
    const reopenedCode = reopenedCodeText?.match(/Enter code ([0-9]{6})/)?.[1]
    if (!reopenedCode) throw Error('reopened browser pairing code missing')
    pc.stdin.write(JSON.stringify({ type: 'claim', code: reopenedCode }) + '\n')
    await page
      .getByRole('region', { name: '演示文稿项目' })
      .getByText('Browser to real PC project', { exact: true })
      .waitFor({ timeout: 20_000 })
    if (outbound.filter((type) => type === 'office.create').length <= createdBeforeReload)
      throw Error('reopened browser did not create a new Relay pairing')
  } finally {
    for (const socket of sockets.values()) socket.close()
    await browser.close()
    if (process.platform === 'linux' && dev.pid) process.kill(-dev.pid, 'SIGTERM')
    else dev.kill('SIGTERM')
  }
}

function firstLine(child, label, timeoutMs, accept = () => true) {
  return new Promise((resolveLine, reject) => {
    const lines = createInterface({ input: child.stdout })
    const output = []
    const fail = (error) => {
      clearTimeout(timer)
      lines.close()
      reject(error)
    }
    const timer = setTimeout(
      () => fail(new Error(`${label} startup timed out: ${output.join(' | ')}`)),
      timeoutMs,
    )
    lines.on('line', (line) => {
      output.push(line)
      if (!accept(line)) return
      clearTimeout(timer)
      lines.close()
      resolveLine(line)
    })
    child.once('error', fail)
    child.once('exit', (code) =>
      fail(new Error(`${label} exited before ready: ${code}: ${output.join(' | ')}`)),
    )
  })
}

try {
  const userDataPath = join(temp, 'pc-data')
  await mkdir(userDataPath)
  const serviceBundle = join(temp, 'presentation-service.cjs')
  const relayBundle = join(temp, 'office-relay-client.cjs')
  const poolBundle = join(temp, 'office-relay-pool.cjs')
  const compilerBundle = join(temp, 'presentation-compiler.cjs')
  await build({
    entryPoints: {
      'presentation-service': join(root, 'apps/shell/src/main/presentation-service.ts'),
      'office-relay-client': join(root, 'apps/shell/src/main/office-relay-client.ts'),
      'office-relay-pool': join(root, 'apps/shell/src/main/office-relay-pool.ts'),
      'presentation-compiler': join(root, 'packages/pptx-engine/src/presentation-compiler.ts'),
    },
    outdir: temp,
    outExtension: { '.js': '.cjs' },
    bundle: true,
    platform: 'node',
    format: 'cjs',
    external: ['electron'],
  })
  const relay = spawn(
    'cargo',
    [
      'run',
      '--quiet',
      '--locked',
      '--manifest-path',
      'services/wiswork-relay/Cargo.toml',
      '--example',
      'local_business_smoke',
    ],
    { cwd: root, env: { ...process.env, CARGO_TARGET_DIR: '/tmp/wiswork-relay-target' } },
  )
  children.push(relay)
  const origin = await firstLine(relay, 'Rust Relay', 300_000)
  if (!/^http:\/\/127\.0\.0\.1:\d+$/.test(origin)) throw new Error('invalid loopback Relay origin')

  const driver = join(temp, 'driver.cjs')
  await writeFile(
    driver,
    `
const { app } = require('electron')
const { readFileSync } = require('node:fs')
const { createPresentationService } = require(${JSON.stringify(serviceBundle)})
const { createOfficeRelayClient } = require(${JSON.stringify(relayBundle)})
const { createOfficeRelayPool } = require(${JSON.stringify(poolBundle)})
const { compilePresentationDeck } = require(${JSON.stringify(compilerBundle)})
const documentId = ${JSON.stringify(documentId)}
const projectId = ${JSON.stringify(projectId)}
const concurrentDocuments = ${JSON.stringify(concurrentDocuments)}
app.whenReady().then(async () => {
  const crashProjectId = ${JSON.stringify(sourceBackedCrash ? researchDeck.id : `${projectId}-crash`)}
  const crashDocumentId = ${JSON.stringify(sourceBackedCrash ? `${researchCaseId}-local-document` : documentId)}
  const crashDeck = ${JSON.stringify(sourceBackedCrash ? researchDeck : { ...deck, id: `${projectId}-crash` })}
  const crashPageId = ${JSON.stringify(sourceBackedCrash ? researchDeck.slides[4].id : 'slide-2')}
  const completedBeforeCrash = ${sourceBackedCrash ? 4 : 1}
  ${p014Fallback ? 'let fallbackFetches = 0' : ''}
  const presentation = createPresentationService({
    userDataPath: ${JSON.stringify(userDataPath)},
    ${
      p014Fallback
        ? `fetchImage: async (url) => {
      if (url === ${JSON.stringify(fallbackUrls[0])}) throw Error('controlled_primary_failure')
      if (url !== ${JSON.stringify(fallbackUrls[1])}) throw Error('unexpected_image_url')
      if (++fallbackFetches !== 1) throw Error('fallback_image_fetched_again')
      return new Response(readFileSync(${JSON.stringify(fallbackImagePath)}), { headers: { 'content-type': 'image/png' } })
    },`
        : ''
    }
    compile: async (input, options) => {
      if (process.env.PPT_AGENT_SMOKE_STALL === '1' && input.id === crashProjectId && input.slides[0]?.id === crashPageId) {
        console.log('PRODUCTION_BLOCKED')
        await new Promise(() => {})
      }
      return compilePresentationDeck(input, options)
    },
  })
  const deck = ${JSON.stringify(deck)}
  if (process.env.PPT_AGENT_SMOKE_RESTARTED !== '1') {
    const compiled = JSON.parse(Buffer.from(await presentation({ operation: 'compile', documentId, requestId: 'run-1', deck }, new AbortController().signal)).toString('utf8'))
    if (compiled.status !== 'compiled') throw Error('Electron PC compile failed')
    const browserDeck = { ...deck, id: ${JSON.stringify(browserProjectId)}, title: 'Browser to real PC project', slides: [deck.slides[0]] }
    const browserCompiled = JSON.parse(Buffer.from(await presentation({ operation: 'compile', documentId: ${JSON.stringify(browserDocumentId)}, requestId: 'browser-run-1', deck: browserDeck }, new AbortController().signal)).toString('utf8'))
    if (browserCompiled.status !== 'compiled') throw Error('Electron PC browser project seed failed')
    const recoveryDeck = { ...deck, id: projectId + '-recovery' }
    const recoverySeed = JSON.parse(Buffer.from(await presentation({ operation: 'compile', documentId, requestId: 'run-recovery-seed', deck: recoveryDeck }, new AbortController().signal)).toString('utf8'))
    if (recoverySeed.status !== 'compiled') throw Error('Electron PC recovery project compile failed')
    for (const item of concurrentDocuments) {
      const onePageDeck = {
        ...deck,
        id: item.projectId,
        slides: [{ ...deck.slides[0], elements: [{ ...deck.slides[0].elements[0], text: item.text }] }],
      }
      const seeded = JSON.parse(Buffer.from(await presentation({ operation: 'compile', documentId: item.documentId, requestId: 'run-concurrent', deck: onePageDeck }, new AbortController().signal)).toString('utf8'))
      if (seeded.status !== 'compiled') throw Error('Electron PC concurrent project compile failed')
    }
  }
  let client
  client = createOfficeRelayPool({
    createClient: events => createOfficeRelayClient({
      endpoint: ${JSON.stringify(origin.replace('http:', 'ws:') + '/office-relay')},
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'local-business-smoke-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      presentationProxy: (body, signal) => presentation(body, signal),
      onPending: events.onPending,
      onPendingExpired: events.onPendingExpired,
      onStatus: events.onStatus,
    }),
    onPending: pending => { void client.approve(pending.pairingId).catch(error => { console.error(error); app.exit(1) }) },
  })
  const crashRequestId = 'production-crash-run'
  const crashCall = async (operation, extra = {}) => JSON.parse(Buffer.from(await presentation({
    operation, documentId: crashDocumentId, projectId: crashProjectId,
    ...(operation === 'save_plan' ? {} : { requestId: crashRequestId }), ...extra,
  }, new AbortController().signal)).toString('utf8'))
  const crashFailure = error => { console.error(error); app.exit(1) }
  async function startCrashJob() {
    if (!${sourceBackedCrash}) {
      const crashPlan = { ...${JSON.stringify(plan)}, projectId: crashProjectId }
      const saved = await crashCall('save_plan', { expectedRevision: 0, plan: crashPlan })
      if (saved.revision !== 1) throw Error('crash fixture plan save failed')
    }
    const begun = await crashCall('production_begin', { planRevision: 1, deck: crashDeck })
    if (begun.status !== 'pending' || begun.total !== 8) throw Error('crash fixture begin failed')
    const started = await crashCall('production_job_start')
    if (started.job?.state !== 'running') throw Error('crash fixture job start failed')
  }
  async function resumeCrashJob() {
    const interrupted = await crashCall('production_job_status')
    if (interrupted.job?.state !== 'interrupted' || interrupted.production?.compiledCount !== completedBeforeCrash ||
        interrupted.production?.pages?.slice(0, completedBeforeCrash).some(page => page.state !== 'compiled' || page.attempt !== 1) ||
        interrupted.production?.pages?.[completedBeforeCrash]?.state !== 'building' ||
        interrupted.production?.pages?.[completedBeforeCrash]?.attempt !== 1 ||
        interrupted.production?.pages?.slice(completedBeforeCrash + 1).some(page => page.state !== 'pending' || page.attempt !== 0))
      throw Error('crash fixture did not recover completed pages and one interrupted page')
    const resumed = await crashCall('production_job_resume')
    if (resumed.job?.state !== 'running') throw Error('crash fixture job resume failed')
    for (let attempt = 0; attempt < 200; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 100))
      const state = await crashCall('production_job_status')
      if (state.job?.state === 'completed') {
        if (state.production.compiledCount !== 8 ||
            state.production.pages.slice(0, completedBeforeCrash).some(page => page.attempt !== 1) ||
            state.production.pages[completedBeforeCrash].attempt !== 2 ||
            state.production.pages.slice(completedBeforeCrash + 1).some(page => page.attempt !== 1) ||
            state.production.pages.some(page => page.state !== 'compiled'))
          throw Error('crash fixture did not preserve completed page receipts')
        console.log('PRODUCTION_RECOVERED')
        return
      }
      if (state.job?.state === 'failed') throw Error('crash fixture resumed job failed')
    }
    throw Error('crash fixture job resume timed out')
  }
  let buffer = ''
  process.stdin.setEncoding('utf8')
  process.stdin.on('data', chunk => {
    buffer += chunk
    let end
    while ((end = buffer.indexOf('\\n')) >= 0) {
      const line = buffer.slice(0, end)
      buffer = buffer.slice(end + 1)
      const command = JSON.parse(line)
      if (command.type === 'claim') void client.claim(command.code).catch(error => { console.error(error); app.exit(1) })
      if (command.type === 'stop') { client.revoke(); app.quit() }
      if (command.type === 'start_crash_job') void startCrashJob().catch(crashFailure)
      if (command.type === 'resume_crash_job') void resumeCrashJob().catch(crashFailure)
    }
  })
  console.log('ELECTRON_READY')
}).catch(error => { console.error(error); app.exit(1) })
`,
  )
  const command = process.platform === 'linux' ? 'xvfb-run' : electron
  const args = process.platform === 'linux' ? ['-a', electron, '--no-sandbox', driver] : [driver]
  async function startPc(restarted, stall = false) {
    const pc = spawn(command, args, {
      cwd: root,
      stdio: ['pipe', 'pipe', 'inherit'],
      detached: process.platform === 'linux',
      env: {
        ...process.env,
        PPT_AGENT_SMOKE_RESTARTED: restarted ? '1' : '0',
        PPT_AGENT_SMOKE_STALL: stall ? '1' : '0',
      },
    })
    children.push(pc)
    await firstLine(pc, 'Electron PC', 30_000, (line) => line === 'ELECTRON_READY')
    return pc
  }
  async function stopPc(pc) {
    const exited = new Promise((resolveExit, reject) => {
      const timer = setTimeout(() => reject(new Error('Electron PC did not exit')), 10_000)
      pc.once('exit', (code) => {
        clearTimeout(timer)
        if (code === 0) resolveExit()
        else reject(new Error(`Electron PC exited ${code}`))
      })
    })
    pc.stdin.write(JSON.stringify({ type: 'stop' }) + '\n')
    await exited
  }
  const pc = await startPc(false)
  smokeStage = 'initial delivery'
  const initialProgress = []
  const result = await inspectPcBusiness(origin, documentId, projectId, {
    onCode: (code) => pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
    onProgress: (stage) => initialProgress.push(stage),
    timeoutMs: 15_000,
    uploadFixtures: true,
    compiledRequestId: 'run-1',
    expectedSlideTexts,
    manualObservation: 'create',
    productionFixture: { requestId: 'production-run-1', deck, plan, expectedSlideTexts },
  }).catch((error) => {
    throw new Error(
      `Electron PC initial smoke failed at ${initialProgress.join(', ')}: ${error.message}`,
    )
  })
  if (
    result.projectId !== projectId ||
    result.slideCount !== 8 ||
    !result.uploadChecked ||
    !result.imageChecked ||
    !result.textChecked ||
    !result.compiledDelivery?.pptxSha256 ||
    result.compiledDelivery.pdfBytes < 100 ||
    result.productionDelivery?.pageDigests.length !== 8 ||
    result.productionDelivery.pdfBytes < 100 ||
    result.manualObservation?.after?.shape?.text !== 'After edit'
  )
    throw new Error('Electron PC business response incomplete')
  smokeStage = 'browser Taskpane to real Relay and PC'
  await inspectBrowserWorkbench(origin, pc)
  const concurrentProgress = []
  smokeStage = concurrentBenchmark
    ? 'P0-17 three concurrent source-backed eight-page productions'
    : 'three concurrent documents'
  const concurrentBundles = concurrentBenchmark
    ? await Promise.all(
        ['science', 'legal', 'finance'].map((variant) =>
          loadBenchmarkBundle(root, 'PPT-P0-17', variant),
        ),
      )
    : []
  const concurrentStartedAt = Date.now()
  const concurrentSettled = await Promise.allSettled(
    concurrentDocuments.map((item, index) =>
      inspectPcBusiness(
        origin,
        item.documentId,
        concurrentBenchmark ? concurrentBundles[index].deck.id : item.projectId,
        {
          onCode: (code) => pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
          onProgress: (stage) => concurrentProgress.push(`${item.projectId}:${stage}`),
          timeoutMs: concurrentBenchmark ? 90_000 : 45_000,
          ...(concurrentBenchmark
            ? {
                sourceAttachments: concurrentBundles[index].sourceAttachments,
                productionFixture: {
                  requestId: `P0-17-parallel-${index + 1}`,
                  deck: concurrentBundles[index].deck,
                  plan: concurrentBundles[index].plan,
                  expectedSlideTexts: concurrentBundles[index].deck.slides.map(
                    (slide) => slide.title,
                  ),
                },
                createProduction: true,
              }
            : { compiledRequestId: 'run-concurrent', expectedSlideTexts: [item.text] }),
        },
      ),
    ),
  )
  if (concurrentSettled.some((entry) => entry.status === 'rejected'))
    throw new Error(
      `Electron PC concurrent document failure: ${concurrentSettled
        .filter((entry) => entry.status === 'rejected')
        .map((entry) => entry.reason?.message)
        .join(', ')}; stages: ${concurrentProgress.join(', ')}`,
    )
  const concurrentResults = concurrentSettled.map((entry) => entry.value)
  if (
    concurrentResults.some(
      (value, index) =>
        value.projectId !==
          (concurrentBenchmark
            ? concurrentBundles[index].deck.id
            : concurrentDocuments[index].projectId) ||
        (!concurrentBenchmark && value.slideCount !== 1) ||
        (concurrentBenchmark
          ? !value.sourceChecked ||
            value.productionDelivery?.pageDigests.length !== 8 ||
            value.productionDelivery.pdfBytes < 100
          : !value.compiledDelivery?.pptxSha256 || value.compiledDelivery.pdfBytes < 100),
    ) ||
    new Set(
      concurrentResults.map((value) =>
        concurrentBenchmark
          ? JSON.stringify(value.productionDelivery?.pageDigests)
          : value.compiledDelivery?.pptxSha256,
      ),
    ).size !== 3
  )
    throw new Error('Electron PC concurrent document sessions crossed project boundaries')
  if (concurrentBenchmark)
    console.log(
      JSON.stringify({
        type: 'ppt_benchmark_timing',
        caseId: 'P0-17:all',
        scope: 'three_source_backed_productions',
        elapsedMs: Date.now() - concurrentStartedAt,
        pageCount: 24,
      }),
    )
  const attachments = join(
    userDataPath,
    'presentation-attachments',
    createHash('sha256').update(documentId).digest('hex'),
  )
  if ((await readdir(attachments)).length !== 0)
    throw new Error('Electron PC test attachments were not cleaned up')
  smokeStage = 'fresh release production'
  const releaseProjectId = `${projectId}-release`
  const releaseResult = await inspectPcBusiness(origin, documentId, releaseProjectId, {
    onCode: (code) => pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
    productionFixture: releaseProductionFixture(releaseProjectId),
    createProduction: true,
    timeoutMs: 45_000,
    ...(p014Fallback
      ? { remoteImageCandidates: { urls: fallbackUrls, attachmentId: fallbackImageSha } }
      : {}),
  })
  if (
    releaseResult.productionDelivery?.pageDigests.length !== 8 ||
    releaseResult.productionDelivery.pdfBytes < 100
  )
    throw new Error('Electron PC fresh release production incomplete')
  if (p014Fallback)
    console.log(`P0-14 controlled image fallback and cache passed: ${fallbackImageSha}`)
  if (!concurrentBenchmark) {
    const cases = benchmarkBatch ? batchCases : [benchmarkCase]
    const failures = []
    for (const caseName of cases) {
      const caseId = `PPT-${caseName}`
      const caseVariant = benchmarkBatch ? undefined : benchmarkVariant
      const bundle = benchmarkBatch
        ? await loadBenchmarkBundle(root, caseId)
        : { plan: researchPlan, deck: researchDeck, sourceAttachments: researchSources }
      const revision =
        caseName === 'P0-18'
          ? benchmarkBatch
            ? {
                requestId: 'P0-18-revised-p04',
                pageId: 'p04',
                slide: JSON.parse(
                  await readFile(
                    join(
                      root,
                      'docs/product/ppt-benchmark-materials/PPT-P0-18/revised-page-deck.json',
                    ),
                    'utf8',
                  ),
                ).slides[0],
              }
            : derivedPageFixture
          : undefined
      smokeStage = `${caseId}${caseVariant ? ` ${caseVariant}` : ''} real-source production`
      const sourceStartedAt = Date.now()
      let sourceOperation = 'pairing'
      const request = () =>
        inspectPcBusiness(
          origin,
          `${caseId}${caseVariant ? `-${caseVariant}` : ''}-local-document`,
          bundle.deck.id,
          {
            onCode: (code) => pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
            onProgress: (stage) => {
              sourceOperation = stage
            },
            timeoutMs: caseName === 'P0-10' ? 120_000 : 60_000,
            sourceAttachments: bundle.sourceAttachments,
            productionFixture: {
              requestId: `${caseId}${caseVariant ? `-${caseVariant}` : ''}-production`,
              deck: bundle.deck,
              plan: bundle.plan,
              expectedSlideTexts: bundle.deck.slides.map((slide) => slide.title),
            },
            ...(revision ? { derivedPageFixture: revision } : {}),
            createProduction: true,
          },
        )
      try {
        let researchResult
        const pairingDeadline = Date.now() + 130_000
        while (!researchResult) {
          try {
            researchResult = await request()
          } catch (error) {
            if (!String(error).includes('create_rate_limited') || Date.now() >= pairingDeadline)
              throw error
            await new Promise((resolveDelay) => setTimeout(resolveDelay, 10_000))
          }
        }
        if (
          !researchResult.sourceChecked ||
          researchResult.productionDelivery?.pageDigests.length !== 8 ||
          researchResult.productionDelivery.pdfBytes < 100 ||
          (revision &&
            (researchResult.productionDelivery.derivedPage?.pageId !== 'p04' ||
              researchResult.productionDelivery.derivedPage.pageDigests.length !== 8))
        )
          throw new Error('source-backed production incomplete')
        console.log(
          JSON.stringify({
            type: 'ppt_benchmark_timing',
            caseId: caseVariant ? `${caseName}:${caseVariant}` : caseName,
            scope: 'source_backed_production',
            elapsedMs: Date.now() - sourceStartedAt,
            pageCount: 8,
          }),
        )
      } catch (error) {
        const message = `${caseId}: ${error.message} after ${Math.round((Date.now() - sourceStartedAt) / 1000)}s at ${sourceOperation}`
        if (!benchmarkBatch) throw new Error(message, { cause: error })
        failures.push(message)
        console.error(message)
      }
    }
    if (failures.length) throw new Error(`benchmark batch failed: ${failures.join(' | ')}`)
  }
  const pendingProjectId = `${projectId}-recovery`
  smokeStage = 'pending production setup'
  const pendingFixture = {
    requestId: 'production-restart-run',
    deck: { ...deck, id: pendingProjectId },
    plan: { ...plan, projectId: pendingProjectId },
    expectedSlideTexts,
  }
  const pending = await inspectPcBusiness(origin, documentId, pendingProjectId, {
    onCode: (code) => pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
    timeoutMs: 15_000,
    productionFixture: pendingFixture,
    beginProductionOnly: true,
  })
  if (pending.productionDelivery?.status !== 'pending' || pending.productionDelivery.total !== 8)
    throw new Error('Electron PC interrupted production did not persist its pending state')
  await stopPc(pc)
  const restartedPc = await startPc(true, true)
  smokeStage = 'pending production resume'
  const resumeProgress = []
  const resumed = await inspectPcBusiness(origin, documentId, pendingProjectId, {
    onCode: (code) => restartedPc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
    onProgress: (stage) => resumeProgress.push(stage),
    timeoutMs: 60_000,
    productionFixture: pendingFixture,
    runExistingProduction: true,
  }).catch((error) => {
    throw new Error(`${error.message}; stages: ${resumeProgress.join(', ')}`)
  })
  if (
    resumed.productionDelivery?.pageDigests.length !== 8 ||
    resumed.productionDelivery.pdfBytes < 100
  )
    throw new Error('Electron PC pending production did not resume after restart')
  smokeStage = 'completed delivery recovery'
  // This smoke opens more than ten pairings from one loopback IP. Respect the
  // Relay's 120-second production limit instead of relaxing it for the test.
  const recoveryDeadline = Date.now() + 130_000
  let recovered
  while (!recovered) {
    try {
      recovered = await inspectPcBusiness(origin, documentId, projectId, {
        onCode: (code) => restartedPc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
        timeoutMs: 15_000,
        compiledRequestId: 'run-1',
        expectedSlideTexts,
        manualObservation: 'read',
        productionFixture: { requestId: 'production-run-1', deck, plan, expectedSlideTexts },
        readExistingProduction: true,
      })
    } catch (error) {
      if (!String(error).includes('create_rate_limited') || Date.now() >= recoveryDeadline)
        throw error
      await new Promise((resolveDelay) => setTimeout(resolveDelay, 10_000))
    }
  }
  // Both PDF responses were independently parsed and checked for page count above.
  // LibreOffice renders on demand, so its output byte length is not a durable receipt.
  const changedAfterRestart = [
    recovered.manualObservation?.after?.digest !== result.manualObservation.after.digest
      ? 'manual_observation'
      : undefined,
    recovered.compiledDelivery?.pptxSha256 !== result.compiledDelivery.pptxSha256
      ? 'compiled_pptx'
      : undefined,
    JSON.stringify(recovered.productionDelivery?.pageDigests) !==
    JSON.stringify(result.productionDelivery.pageDigests)
      ? 'production_pages'
      : undefined,
  ].filter(Boolean)
  if (changedAfterRestart.length)
    throw new Error(`Electron PC delivery changed after restart: ${changedAfterRestart.join(', ')}`)
  smokeStage = 'running production crash'
  const blocked = firstLine(
    restartedPc,
    'Electron PC blocked production',
    30_000,
    (line) => line === 'PRODUCTION_BLOCKED',
  )
  restartedPc.stdin.write(JSON.stringify({ type: 'start_crash_job' }) + '\n')
  await blocked
  if (process.platform === 'linux') process.kill(-restartedPc.pid, 'SIGKILL')
  else restartedPc.kill('SIGKILL')
  await new Promise((resolveExit) => restartedPc.once('exit', resolveExit))
  const recoveredPc = await startPc(true)
  smokeStage = 'running production recovery'
  const recoveredJob = firstLine(
    recoveredPc,
    'Electron PC production recovery',
    30_000,
    (line) => line === 'PRODUCTION_RECOVERED',
  )
  recoveredPc.stdin.write(JSON.stringify({ type: 'resume_crash_job' }) + '\n')
  await recoveredJob
  await stopPc(recoveredPc)
  console.log(
    `Electron PC + Rust Relay business smoke passed (${builtTaskpane ? 'built' : 'development'} Taskpane): ${builtTaskpane ? 'incompatible protocol metadata blocks connection until retry, ' : ''}browser Taskpane pairing, project readback, presentation copy action, real Relay session resume and Taskpane reopen, ${concurrentBenchmark ? 'three parallel P0-17 source-backed eight-page productions' : 'three concurrent documents'}, fresh eight-page release production${concurrentBenchmark ? '' : `, ${benchmarkBatch ? batchCases.join('/') : researchCaseId} frozen source upload and eight-page production`}${derivedPageFixture || benchmarkBatch ? ', P0-18 parent-bound single-page revision preserving seven page packages' : ''}, PPTX/PDF readback, TXT/PNG upload, durable delivery and manual observation, pending production recovery and ${sourceBackedCrash ? 'P0-20 source-backed page-five crash recovery' : 'running job crash recovery'}`,
  )
} catch (error) {
  throw new Error(`Electron PC smoke failed during ${smokeStage}: ${error.message}`, {
    cause: error,
  })
} finally {
  for (const child of children.reverse()) {
    if (child.exitCode !== null) continue
    if (
      process.platform === 'linux' &&
      (child.spawnfile === 'xvfb-run' || child.spawnfile === 'npm') &&
      child.pid
    ) {
      try {
        process.kill(-child.pid, 'SIGTERM')
      } catch {
        // The process group may have exited between inspection and cleanup.
      }
    } else child.kill('SIGTERM')
  }
  await rm(temp, { recursive: true, force: true })
}
