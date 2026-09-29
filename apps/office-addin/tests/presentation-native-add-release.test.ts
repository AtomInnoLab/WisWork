import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  describePagePackageBackup,
  saveChartPackageBackup,
  readChartPackageBackup,
} from '../src/skills/powerpoint/presentation-chart-backup'
import { officeOperationsForSlideIR } from '../src/skills/powerpoint/presentation-office-ir'
import { createPresentationNativeAddRelease } from '../src/skills/powerpoint/presentation-native-add-release'
import type { PresentationNativeAddBatch } from '../src/skills/powerpoint/presentation-existing-batch'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function fixture() {
  const dir = mkdtempSync(join(tmpdir(), 'native-release-'))
  roots.push(dir)
  let service = createPresentationService({ userDataPath: dir })
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
      ),
  )
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  const binding = () =>
    createPresentationDocumentBinding(
      {
        get: (k) => values.get(k),
        set: (k, v) => {
          values.set(k, v)
        },
        save,
        location: () => 'synthetic',
      },
      () => 'doc',
    )
  const documentId = await binding().documentId(),
    deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck),
    base64 = Buffer.from(bytes).toString('base64'),
    meta = await describePagePackageBackup(base64)
  const scope = { request, documentId, hostSlideId: 'host', slideIds: ['host'] }
  const backup = await saveChartPackageBackup({ ...scope, backupId: 'native-original', base64 })
  const source = await saveChartPackageBackup({
    ...scope,
    backupId: 'page-restoration-source',
    base64,
  })
  const operations = officeOperationsForSlideIR(
    {
      ...deck.slides[0]!,
      elements: [
        {
          kind: 'shape',
          id: 'native',
          shape: 'rect',
          x: 1,
          y: 1,
          w: 1,
          h: 1,
          fill: 'FFFFFF',
          lineColor: '000000',
        },
      ],
      claimIds: [],
    },
    deck.style,
    0,
  )
  const initial: PresentationNativeAddBatch = {
    version: 2,
    kind: 'native_page_add',
    changeId: 'native-add',
    documentId,
    baselineId: 'baseline',
    baselineDigest: meta.packageDigest,
    hostSlideId: 'host',
    slideIndex: 0,
    beforeSlideIds: ['host'],
    scope: { slideIds: ['host'] },
    intent: 'add',
    preserved: ['original'],
    validation: ['package'],
    risk: 'high',
    backups: [{ ...backup, hostSlideId: 'host', packageDigest: meta.packageDigest }],
    operations,
    nextIndex: 0,
    createdShapeIds: [],
    state: 'applying',
  }
  await binding().writeExistingBatch(initial, undefined)
  const undoing = { ...initial, state: 'undoing' as const }
  await binding().writeExistingBatch(undoing, initial)
  const record = { ...undoing, state: 'undone' as const, restoredSlideId: 'restored' }
  await binding().writeExistingBatch(record, undoing)
  const proposals = createStructuredProposalController()
  const options = {
    proposals,
    documentId: () => binding().documentId(),
    request,
    readExistingBatch: (id: string) => binding().readExistingBatch(id),
    writeExistingBatch: (
      next: Parameters<ReturnType<typeof binding>['writeExistingBatch']>[0],
      expected: Parameters<ReturnType<typeof binding>['writeExistingBatch']>[1],
    ) => binding().writeExistingBatch(next, expected),
  }
  return {
    record,
    request,
    scope,
    source,
    base64,
    binding,
    save,
    proposals,
    options,
    factory: (overrides = {}) => createPresentationNativeAddRelease({ ...options, ...overrides }),
    restart: () => {
      service = createPresentationService({ userDataPath: dir })
    },
  }
}
async function confirm(f: Awaited<ReturnType<typeof fixture>>) {
  const p = await f.factory().propose('native-add')
  const decision = f.proposals.waitForDecision(p.id)
  await f.proposals.confirm(p.id)
  expect((await decision).status).toBe('confirmed')
}
it('requires confirmation, releases only native backup, persists receipt and preserves restoration source', async () => {
  const f = await fixture(),
    p = await f.factory().propose('native-add')
  expect(
    f.request.mock.calls.filter(
      ([b]) => (b as { operation?: string }).operation === 'existing_page_backup_release',
    ),
  ).toHaveLength(0)
  expect(
    await readChartPackageBackup({
      ...f.scope,
      backup: f.record.backups[0]!,
      expectedPackageDigest: f.record.baselineDigest,
    }),
  ).toBe(f.base64)
  const decision = f.proposals.waitForDecision(p.id)
  await f.proposals.confirm(p.id)
  expect((await decision).status).toBe('confirmed')
  expect(f.binding().readExistingBatch('native-add')).toHaveProperty('backupReleasedAt')
  await expect(
    readChartPackageBackup({ ...f.scope, backup: f.record.backups[0]! }),
  ).rejects.toThrow()
  expect(await readChartPackageBackup({ ...f.scope, backup: f.source })).toBe(f.base64)
  f.restart()
  expect(f.binding().readExistingBatch('native-add')).toHaveProperty('backupReleasedAt')
  await expect(f.factory().propose('native-add')).rejects.toThrow()
})
it('retries a real PC release whose acknowledgment was lost', async () => {
  const f = await fixture()
  const p = await f
    .factory({
      request: async (body: unknown, s?: AbortSignal) => {
        const response = await f.request(body, s)
        if ((body as { operation?: string }).operation === 'existing_page_backup_release')
          throw Error('lost ACK')
        return response
      },
    })
    .propose('native-add')
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow(
    'presentation_native_add_backup_release_failed',
  )
  expect((await decision).status).toBe('failed')
  expect(f.binding().readExistingBatch('native-add')).not.toHaveProperty('backupReleasedAt')
  f.restart()
  await confirm(f)
  expect(f.binding().readExistingBatch('native-add')).toHaveProperty('backupReleasedAt')
})
it('retries after a real settings save failure without losing the PC release receipt', async () => {
  const f = await fixture()
  f.save.mockRejectedValueOnce(Error('settings failed'))
  const p = await f.factory().propose('native-add')
  const d = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow(
    'presentation_native_add_backup_release_failed',
  )
  await d
  expect(f.binding().readExistingBatch('native-add')).not.toHaveProperty('backupReleasedAt')
  await confirm(f)
})
it('rejects cross-document confirmation and changed records before deletion', async () => {
  const f = await fixture()
  let doc = f.record.documentId
  const p = await f.factory({ documentId: async () => doc }).propose('native-add')
  doc = 'other'
  const d = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  await d
  expect(await readChartPackageBackup({ ...f.scope, backup: f.record.backups[0]! })).toBe(f.base64)
  await expect(
    f
      .factory({ readExistingBatch: () => ({ ...f.record, changeId: 'other' }) })
      .propose('native-add'),
  ).rejects.toThrow()
})
it('does no delete for pre-cancelled calls or unconfirmed cancellation', async () => {
  const f = await fixture(),
    abort = new AbortController()
  abort.abort()
  await expect(f.factory().propose('native-add', abort.signal)).rejects.toThrow('cancelled')
  await f.factory().propose('native-add')
  f.proposals.newTurn()
  expect(await readChartPackageBackup({ ...f.scope, backup: f.record.backups[0]! })).toBe(f.base64)
})
it('does not mark released for a mismatched receipt', async () => {
  const f = await fixture()
  const p = await f
    .factory({ request: async () => Response.json({ status: 'released', backupId: 'wrong' }) })
    .propose('native-add')
  const d = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  await d
  expect(f.binding().readExistingBatch('native-add')).not.toHaveProperty('backupReleasedAt')
})

it('routes the actual Skill release tool through explicit confirmation and real durable PC receipt', async () => {
  const f = await fixture()
  const skill = createPowerPointSkill({
    adapter: {} as PowerPointAdapter,
    proposals: f.proposals,
    nativeAddSavepoint: f.options,
  })
  expect(skill.tools.some((tool) => tool.name === 'release_slide_ir_addition')).toBe(true)
  const result = await skill.executeTool({
    id: 'release',
    name: 'release_slide_ir_addition',
    input: { change_id: 'native-add' },
  })
  expect(result.isError).not.toBe(true)
  expect(result.mutated).toBe(false)
  expect(f.binding().readExistingBatch('native-add')).not.toHaveProperty('backupReleasedAt')
  const p = f.proposals.pending()!,
    d = f.proposals.waitForDecision(p.id)
  expect(p.impact).toMatchObject({ targets: ['backup:native-original'], count: 1 })
  await f.proposals.confirm(p.id)
  expect((await d).status).toBe('confirmed')
  expect(f.binding().readExistingBatch('native-add')).toHaveProperty('backupReleasedAt')
  expect(await readChartPackageBackup({ ...f.scope, backup: f.source })).toBe(f.base64)
})
it('refuses a record that changed after the proposal without deleting its retained backup', async () => {
  const f = await fixture()
  const p = await f.factory().propose('native-add')
  await f
    .binding()
    .writeExistingBatch({ ...f.record, backupReleasedAt: '2026-09-29T00:00:00.000Z' }, f.record)
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  expect((await decision).status).toBe('failed')
  expect(await readChartPackageBackup({ ...f.scope, backup: f.record.backups[0]! })).toBe(f.base64)
})
it('refuses active native records and validates exact release receipt scope', async () => {
  const f = await fixture()
  const { restoredSlideId: _restored, ...rest } = f.record
  await expect(
    f.factory({ readExistingBatch: () => ({ ...rest, state: 'applying' }) }).propose('native-add'),
  ).rejects.toThrow('presentation_native_add_state_invalid')
  const p = await f
    .factory({
      request: async (body: unknown) => {
        const response = await f.request(body)
        const receipt = await response.json()
        return Response.json({ ...receipt, hostSlideId: 'different-host' })
      },
    })
    .propose('native-add')
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow(
    'presentation_native_add_backup_release_failed',
  )
  expect((await decision).status).toBe('failed')
  expect(f.binding().readExistingBatch('native-add')).not.toHaveProperty('backupReleasedAt')
  await confirm(f)
})
