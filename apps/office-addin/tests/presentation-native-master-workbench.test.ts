import { createPresentationExistingEditingSkill } from '../src/skills/powerpoint/presentation-existing-editing'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { expect, it, vi } from 'vitest'
import { createPresentationChangesController } from '../src/agent/presentation-changes'
import type { PresentationNativeMasterChange } from '../src/skills/powerpoint/presentation-native-master-change'
const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 100 })
function record(): PresentationNativeMasterChange {
  return {
    version: 1,
    kind: 'native_master',
    changeId: 'master-change',
    documentId: 'doc',
    intent: 'theme',
    snapshotRef: ref('snapshot'),
    operations: [
      { op: 'set_master_theme_color', master_id: 'm1', theme_color: 'Accent1', color: '#112233' },
    ],
    inverseOperations: [
      { op: 'set_master_theme_color', master_id: 'm1', theme_color: 'Accent1', color: '#FFFFFF' },
    ],
    scope: { masterIds: ['m1'], affectedPageCount: 600 },
    nextIndex: 0,
    state: 'applying',
    currentProofRef: ref('receipt-0'),
    receipts: [],
    reviews: [],
  }
}
async function fixture(r = record()) {
  let current = r,
    available = true
  const executeTool = vi.fn(async () => ({ output: '{}', mutated: false, summary: 'proposed' }))
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => false,
    nativeMasterAvailable: () => available,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      {
        id: 'native_master:master-change',
        kind: 'native_master',
        record: current,
        legacy: false,
        sequence: 1,
      },
    ],
    executeTool,
  })
  await controller.refresh()
  return {
    controller,
    executeTool,
    setRecord: (value: PresentationNativeMasterChange) => {
      current = value
    },
    disconnect: () => {
      available = false
    },
  }
}
it('projects master savepoint with full dependency count and never uses page replacement recovery', async () => {
  const f = await fixture()
  expect(f.controller.snapshot().entries).toHaveLength(1)
  expect(f.controller.snapshot().entries[0]).toMatchObject({
    kind: 'master',
    source: 'native_master',
    affectedPageCount: 600,
    cursor: 0,
    operationCount: 1,
  })
  await f.controller.run('native_master:master-change', 'resume')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'resume_slide_master_change',
      input: { change_id: 'master-change' },
    }),
    expect.any(AbortSignal),
  )
  expect(f.executeTool.mock.calls.some((call) => String(call).includes('stage_existing'))).toBe(
    false,
  )
})
it('offers explicit receipt reconciliation for pending native writes without a replay action', async () => {
  const r = record()
  r.pending = { direction: 'forward', index: 0, beforeProofRef: r.currentProofRef }
  const f = await fixture(r)
  expect(f.controller.snapshot().entries[0].actions).toEqual(['inspect', 'reconcile'])
  await f.controller.run('native_master:master-change', 'reconcile')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({ name: 'reconcile_slide_master_change' }),
    expect.any(AbortSignal),
  )
})
it('routes applied master undo by change ID and blocks stale ledger changes', async () => {
  const r = record()
  r.state = 'applied'
  r.nextIndex = 1
  r.currentProofRef = ref('receipt-1')
  r.receipts = [{ direction: 'forward', index: 0, proofRef: r.currentProofRef }]
  const f = await fixture(r)
  await f.controller.run('native_master:master-change', 'undo')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'undo_slide_master_change',
      input: { change_id: 'master-change' },
    }),
    expect.any(AbortSignal),
  )
  f.executeTool.mockClear()
  f.setRecord({ ...r, intent: 'changed baseline' })
  await f.controller.run('native_master:master-change', 'undo')
  expect(f.executeTool).not.toHaveBeenCalled()
  expect(f.controller.snapshot().error).toBeDefined()
})
it('drops master actions after PC availability changes', async () => {
  const f = await fixture()
  f.disconnect()
  await f.controller.run('native_master:master-change', 'resume')
  expect(f.executeTool).not.toHaveBeenCalled()
  await f.controller.refresh()
  expect(f.controller.snapshot().entries).toEqual([])
})

it('lists standalone master history without requiring an artifact or certifying current host', async () => {
  const r = record()
  const hostRead = vi.fn(async () => {
    throw Error('unexpected_host_read')
  })
  const skill = createPresentationExistingEditingSkill({
    baseline: {} as never,
    baselineAdapter: { readPage: hostRead } as never,
    adapter: { inspectSlideMasters: hostRead } as never,
    proposals: createStructuredProposalController(),
    documentId: async () => 'doc',
    request: async () => {
      throw Error('unexpected_pc_request')
    },
    listChangeHistory: () => [
      {
        id: 'native_master:master-change',
        kind: 'native_master',
        record: r,
        legacy: false,
        sequence: 1,
      },
    ],
    readExistingChange: () => undefined,
    writeExistingChange: async () => {
      throw Error('unexpected_write')
    },
  })
  const result = await skill.executeTool({
    id: 'list',
    name: 'list_existing_presentation_changes',
    input: {},
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    currentHostVerified: false,
    changes: [
      {
        kind: 'native_master',
        changeId: 'master-change',
        masterIds: ['m1'],
        affectedPageCount: 600,
        pending: false,
        currentHostVerified: false,
      },
    ],
  })
  expect(hostRead).not.toHaveBeenCalled()
})
