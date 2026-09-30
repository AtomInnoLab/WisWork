import { spawn } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdir, mkdtemp, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'
import electron from 'electron'
import { inspectPcBusiness } from './ppt-agent-pc-business-smoke.mjs'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const temp = await mkdtemp(join(tmpdir(), 'ppt-electron-real-relay-'))
const children = []
const documentId = 'electron-real-relay-document'
const projectId = 'electron-real-relay-project'
const expectedSlideTexts = Array.from(
  { length: 8 },
  (_, index) => `Electron real Relay page ${index + 1}`,
)
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
  await build({
    entryPoints: {
      'presentation-service': join(root, 'apps/shell/src/main/presentation-service.ts'),
      'office-relay-client': join(root, 'apps/shell/src/main/office-relay-client.ts'),
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
const { createPresentationService } = require(${JSON.stringify(serviceBundle)})
const { createOfficeRelayClient } = require(${JSON.stringify(relayBundle)})
const documentId = ${JSON.stringify(documentId)}
const projectId = ${JSON.stringify(projectId)}
app.whenReady().then(async () => {
  const presentation = createPresentationService({ userDataPath: ${JSON.stringify(userDataPath)} })
  const deck = ${JSON.stringify(deck)}
  if (process.env.PPT_AGENT_SMOKE_RESTARTED !== '1') {
    const compiled = JSON.parse(Buffer.from(await presentation({ operation: 'compile', documentId, requestId: 'run-1', deck }, new AbortController().signal)).toString('utf8'))
    if (compiled.status !== 'compiled') throw Error('Electron PC compile failed')
  }
  let client
  client = createOfficeRelayClient({
    endpoint: ${JSON.stringify(origin.replace('http:', 'ws:') + '/office-relay')},
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'local-business-smoke-token',
    proxy: async () => ({ status: 200, body: new Uint8Array() }),
    presentationProxy: (body, signal) => presentation(body, signal),
    onPending: (pending) => { void client.approve(pending.pairingId).catch(error => { console.error(error); app.exit(1) }) },
  })
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
    }
  })
  console.log('ELECTRON_READY')
}).catch(error => { console.error(error); app.exit(1) })
`,
  )
  const command = process.platform === 'linux' ? 'xvfb-run' : electron
  const args = process.platform === 'linux' ? ['-a', electron, '--no-sandbox', driver] : [driver]
  async function startPc(restarted) {
    const pc = spawn(command, args, {
      cwd: root,
      stdio: ['pipe', 'pipe', 'inherit'],
      env: { ...process.env, PPT_AGENT_SMOKE_RESTARTED: restarted ? '1' : '0' },
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
  const result = await inspectPcBusiness(origin, documentId, projectId, {
    onCode: (code) => pc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
    timeoutMs: 15_000,
    uploadFixtures: true,
    compiledRequestId: 'run-1',
    expectedSlideTexts,
    productionFixture: { requestId: 'production-run-1', deck, plan, expectedSlideTexts },
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
    result.productionDelivery.pdfBytes < 100
  )
    throw new Error('Electron PC business response incomplete')
  const attachments = join(
    userDataPath,
    'presentation-attachments',
    createHash('sha256').update(documentId).digest('hex'),
  )
  if ((await readdir(attachments)).length !== 0)
    throw new Error('Electron PC test attachments were not cleaned up')
  await stopPc(pc)
  const restartedPc = await startPc(true)
  const recovered = await inspectPcBusiness(origin, documentId, projectId, {
    onCode: (code) => restartedPc.stdin.write(JSON.stringify({ type: 'claim', code }) + '\n'),
    timeoutMs: 15_000,
    compiledRequestId: 'run-1',
    expectedSlideTexts,
    productionFixture: { requestId: 'production-run-1', deck, plan, expectedSlideTexts },
    readExistingProduction: true,
  })
  if (
    recovered.compiledDelivery?.pptxSha256 !== result.compiledDelivery.pptxSha256 ||
    recovered.compiledDelivery?.pdfBytes !== result.compiledDelivery.pdfBytes ||
    JSON.stringify(recovered.productionDelivery?.pageDigests) !==
      JSON.stringify(result.productionDelivery.pageDigests) ||
    recovered.productionDelivery?.pdfBytes !== result.productionDelivery.pdfBytes
  )
    throw new Error('Electron PC delivery changed after restart')
  await stopPc(restartedPc)
  console.log(
    'Electron PC + Rust Relay business smoke passed: pairing, eight-page compile and planned page production, import-source digests, PPTX/PDF readback, TXT/PNG upload, native image readback, cleanup and durable delivery after PC restart',
  )
} finally {
  for (const child of children.reverse()) {
    if (child.exitCode === null) child.kill('SIGTERM')
  }
  await rm(temp, { recursive: true, force: true })
}
