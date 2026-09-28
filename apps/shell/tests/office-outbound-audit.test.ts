import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, expect, it, vi } from 'vitest'
import { WISWORK_MESSAGES_URL } from '@wiswork/ai-provider'
import { createOfficeOutboundAudit } from '../src/main/office-outbound-audit'
import { createOfficeMessagesProxy } from '../src/main/office-bridge-runtime'

const roots: string[] = []
afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

it('durably records a fixed-destination model request without body or credentials', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'office-outbound-audit-'))
  roots.push(userDataPath)
  const audit = createOfficeOutboundAudit({ userDataPath, now: () => 1_000 })
  const body = { messages: [{ role: 'user', content: 'PRIVATE-SIM-P0-06-7E91' }] }
  const fetchWithAuth = vi.fn(async (request: (token: string) => Promise<Response>) =>
    request('private-token'),
  )
  const fetch = vi.fn(async () => {
    const records = await audit.list()
    expect(records).toMatchObject([{ state: 'attempted', destination: WISWORK_MESSAGES_URL }])
    return new Response('data: ok\n\n', { status: 200 })
  })
  const proxy = createOfficeMessagesProxy({ fetchWithAuth, fetch, audit })
  expect((await proxy({ body, signal: new AbortController().signal })).status).toBe(200)
  const records = await createOfficeOutboundAudit({ userDataPath }).list()
  expect(records).toMatchObject([
    {
      state: 'response_received',
      destination: WISWORK_MESSAGES_URL,
      status: 200,
      requestBytes: Buffer.byteLength(JSON.stringify(body)),
      requestSha256: createHash('sha256').update(JSON.stringify(body)).digest('hex'),
    },
  ])
  const raw = await readFile(join(userDataPath, 'office-outbound-audit.json'), 'utf8')
  expect(raw).not.toContain('PRIVATE-SIM-P0-06-7E91')
  expect(raw).not.toContain('private-token')
  const report = execFileSync(
    process.execPath,
    [
      fileURLToPath(new URL('../../../scripts/office-outbound-audit-report.mjs', import.meta.url)),
      userDataPath,
    ],
    { encoding: 'utf8' },
  )
  expect(JSON.parse(report)).toMatchObject({ total: 1, counts: { response_received: 1 } })
  expect(report).not.toContain('PRIVATE-SIM-P0-06-7E91')
  expect(report).not.toContain('private-token')
  expect(fetch).toHaveBeenCalledOnce()
})

it('retains an honest failed or unfinished request rather than claiming delivery', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'office-outbound-audit-'))
  roots.push(userDataPath)
  const audit = createOfficeOutboundAudit({ userDataPath })
  const first = await audit.begin('{}', WISWORK_MESSAGES_URL)
  const proxy = createOfficeMessagesProxy({
    fetchWithAuth: async (request) => request('token'),
    fetch: async () => {
      throw new Error('offline')
    },
    audit,
  })
  await expect(
    proxy({ body: { messages: [] }, signal: new AbortController().signal }),
  ).rejects.toThrow('offline')
  const records = await createOfficeOutboundAudit({ userDataPath }).list()
  expect(records[0]).toMatchObject({ id: first, state: 'attempted' })
  expect(records[1]).toMatchObject({ state: 'failed' })
  expect(records.every((record) => record.state !== 'response_received')).toBe(true)
})

it('records terminal authentication loss and revokes the bridge capability', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'office-outbound-audit-'))
  roots.push(userDataPath)
  const revoked = vi.fn()
  const proxy = createOfficeMessagesProxy({
    fetchWithAuth: async (request) => {
      await request('expired-token')
      return request('refreshed-token')
    },
    fetch: async () => new Response('', { status: 401 }),
    onTerminalAuthLoss: revoked,
    audit: createOfficeOutboundAudit({ userDataPath }),
  })
  await expect(proxy({ body: {}, signal: new AbortController().signal })).rejects.toThrow(
    'auth_required',
  )
  expect(revoked).toHaveBeenCalledOnce()
  expect(await createOfficeOutboundAudit({ userDataPath }).list()).toMatchObject([
    { state: 'auth_required', destination: WISWORK_MESSAGES_URL },
    { state: 'auth_required', destination: WISWORK_MESSAGES_URL },
  ])
})

it('records both actual HTTP attempts when token refresh retries a 401', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'office-outbound-audit-'))
  roots.push(userDataPath)
  let attempts = 0
  const proxy = createOfficeMessagesProxy({
    fetchWithAuth: async (request) => {
      let response = await request('expired-token')
      if (response.status === 401) response = await request('fresh-token')
      return response
    },
    fetch: async () => new Response('', { status: ++attempts === 1 ? 401 : 200 }),
    audit: createOfficeOutboundAudit({ userDataPath }),
  })
  expect(
    (await proxy({ body: { messages: [] }, signal: new AbortController().signal })).status,
  ).toBe(200)
  expect(attempts).toBe(2)
  expect(await createOfficeOutboundAudit({ userDataPath }).list()).toMatchObject([
    { state: 'auth_required' },
    { state: 'response_received', status: 200 },
  ])
})

it('records cancellation separately from a failed or completed request', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'office-outbound-audit-'))
  roots.push(userDataPath)
  const controller = new AbortController()
  const proxy = createOfficeMessagesProxy({
    fetchWithAuth: async (request) => request('token'),
    fetch: async () => {
      controller.abort()
      throw new Error('request aborted')
    },
    audit: createOfficeOutboundAudit({ userDataPath }),
  })
  await expect(proxy({ body: {}, signal: controller.signal })).rejects.toThrow('request aborted')
  expect(await createOfficeOutboundAudit({ userDataPath }).list()).toMatchObject([
    { state: 'aborted', destination: WISWORK_MESSAGES_URL },
  ])
})

it('fails closed before a model request when the local audit cannot be read', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'office-outbound-audit-'))
  roots.push(userDataPath)
  await writeFile(join(userDataPath, 'office-outbound-audit.json'), '{not valid json')
  const fetch = vi.fn(async () => new Response('ok'))
  const proxy = createOfficeMessagesProxy({
    fetchWithAuth: async (request) => request('token'),
    fetch,
    audit: createOfficeOutboundAudit({ userDataPath }),
  })
  await expect(proxy({ body: {}, signal: new AbortController().signal })).rejects.toThrow(
    'office_audit_invalid_state',
  )
  expect(fetch).not.toHaveBeenCalled()
})
