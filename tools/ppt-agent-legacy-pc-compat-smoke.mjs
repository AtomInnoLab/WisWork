import { spawn, execFileSync } from 'node:child_process'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createRequire } from 'node:module'
import { createInterface } from 'node:readline'
import { build } from 'esbuild'
import WebSocket from 'ws'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const legacyRevision = '1d611fd3'
const token = 'local-business-smoke-token'
const origin = 'https://office.8-216-134-194.sslip.io'
const temp = await mkdtemp(join(tmpdir(), 'wiswork-legacy-pc-'))
let relay
let office
let modernPc
let legacyClient

function receive(socket, type, timeoutMs = 10_000) {
  return new Promise((resolveFrame, reject) => {
    const timer = setTimeout(() => finish(new Error(`timed out waiting for ${type}`)), timeoutMs)
    const onMessage = (data) => {
      let frame
      try {
        frame = JSON.parse(data.toString())
      } catch {
        finish(new Error('invalid Relay frame'))
        return
      }
      if (frame.type === type) finish(undefined, frame)
      else if (frame.type === 'relay.error') finish(new Error(`Relay error: ${frame.code}`))
    }
    const onClose = () => finish(new Error('Relay socket closed'))
    function finish(error, frame) {
      clearTimeout(timer)
      socket.off('message', onMessage)
      socket.off('close', onClose)
      if (error) reject(error)
      else resolveFrame(frame)
    }
    socket.on('message', onMessage)
    socket.on('close', onClose)
  })
}

async function opened(socket) {
  await new Promise((resolveOpen, reject) => {
    const timer = setTimeout(() => reject(new Error('Relay socket open timed out')), 10_000)
    socket.once('open', () => {
      clearTimeout(timer)
      resolveOpen()
    })
    socket.once('error', (error) => {
      clearTimeout(timer)
      reject(error)
    })
  })
}

try {
  const source = execFileSync(
    'git',
    ['show', `${legacyRevision}:apps/shell/src/main/office-relay-client.ts`],
    { cwd: root, encoding: 'utf8' },
  )
  const output = join(temp, 'legacy-client.cjs')
  await build({
    stdin: { contents: source, loader: 'ts', resolveDir: root, sourcefile: 'legacy-client.ts' },
    outfile: output,
    bundle: true,
    platform: 'node',
    format: 'cjs',
  })
  const { createOfficeRelayClient } = createRequire(import.meta.url)(output)
  relay = spawn(
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
  const lines = createInterface({ input: relay.stdout })
  const relayOrigin = await Promise.race([
    new Promise((resolveOrigin, reject) => {
      lines.once('line', resolveOrigin)
      relay.once('exit', (code) => reject(new Error(`Relay exited ${code}`)))
    }),
    new Promise((_resolve, reject) => {
      setTimeout(() => reject(new Error('Relay start timed out')), 300_000).unref()
    }),
  ])
  if (!/^http:\/\/127\.0\.0\.1:\d+$/.test(relayOrigin)) throw new Error('invalid Relay origin')
  const url = `${relayOrigin.replace(/^http:/, 'ws:')}/office-relay`
  office = new WebSocket(url, { headers: { Origin: origin } })
  await opened(office)
  const createdPromise = receive(office, 'office.created')
  office.send(
    JSON.stringify({
      version: 2,
      type: 'office.create',
      host: 'PowerPoint',
      capabilities: ['agent.v1'],
    }),
  )
  const created = await createdPromise
  if (!/^\d{6}$/.test(created.verification_code)) throw new Error('invalid v2 invitation')

  const legacyStatuses = []
  legacyClient = createOfficeRelayClient({
    endpoint: url,
    connect: (endpoint, accessToken) =>
      new WebSocket(endpoint, { headers: { Authorization: `Bearer ${accessToken}` } }),
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => token,
    proxy: {},
    onPending() {},
    onStatus(status) {
      legacyStatuses.push(status)
    },
  })
  const incompatiblePromise = receive(office, 'office.pc_incompatible')
  await legacyClient.claim(created.verification_code)
  const incompatible = await incompatiblePromise
  if (incompatible.pairing_id !== created.pairing_id)
    throw new Error('legacy PC entered an invalid paired state')
  const deadline = Date.now() + 5_000
  while (!legacyStatuses.includes('disconnected:protocol_violation') && Date.now() < deadline)
    await new Promise((resolveNext) => setTimeout(resolveNext, 10))
  if (!legacyStatuses.includes('disconnected:protocol_violation'))
    throw new Error(`legacy PC did not reject v2 invitation: ${legacyStatuses.join(',')}`)
  if (legacyClient.status() !== 'disconnected:protocol_violation')
    throw new Error('legacy PC entered an invalid paired state')

  modernPc = new WebSocket(url, { headers: { Authorization: `Bearer ${token}` } })
  await opened(modernPc)
  const claimedPromise = receive(modernPc, 'pc.claimed')
  modernPc.send(
    JSON.stringify({
      version: 2,
      type: 'pc.claim',
      verification_code: created.verification_code,
      capabilities: ['agent.v1'],
    }),
  )
  const claimed = await claimedPromise
  if (claimed.pairing_id !== created.pairing_id)
    throw new Error('v2 invitation was consumed by legacy PC')
  process.stdout.write(
    `Legacy PC ${legacyRevision} compatibility smoke passed: v1 claim rejected, v2 invitation retained\n`,
  )
} finally {
  legacyClient?.revoke()
  office?.close()
  modernPc?.close()
  relay?.kill('SIGTERM')
  await rm(temp, { recursive: true, force: true })
}
