import { createPresentationExistingEditingSkill } from '../src/skills/powerpoint/presentation-existing-editing'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { expect, it, vi } from 'vitest'
import { createPresentationChangesController } from '../src/agent/presentation-changes'
import type { PresentationPackageChange } from '../src/skills/powerpoint/presentation-package-change'
const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 100 })
function record(): PresentationPackageChange {
  return {
    version: 1,
    kind: 'package_xml',
    changeId: 'xml-change',
    documentId: 'doc',
    sourceKind: 'slide',
    sourceSlideId: 'host-slide',
    packageSourceSlideId: '256#',
    intent: 'XML背景修改',
    snapshotRef: ref('snapshot'),
    originalRef: ref('page-0'),
    preparedRef: ref('page-1'),
    currentProofRef: ref('receipt-0'),
    state: 'prepared',
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
    packageAvailable: () => available,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      {
        id: 'package_xml:xml-change',
        kind: 'package_xml',
        record: current,
        legacy: false,
        sequence: 1,
      },
    ],
    executeTool,
  } as any)
  await controller.refresh()
  return {
    controller,
    executeTool,
    setRecord: (r: PresentationPackageChange) => {
      current = r
    },
    disconnect: () => {
      available = false
    },
  }
}
it('shows XML savepoints independently of a generated artifact and routes resume by exact change ID', async () => {
  const f = await fixture()
  expect(f.controller.snapshot().entries[0]).toMatchObject({
    source: 'package_xml',
    kind: 'page',
    affectedPageCount: 1,
    actions: ['inspect', 'resume', 'discard'],
  })
  await f.controller.run('package_xml:xml-change', 'resume')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'resume_package_xml_change',
      input: { change_id: 'xml-change' },
    }),
    expect.any(AbortSignal),
  )
})
it('offers reconciliation only for uncertain XML import without replay or original-page routing', async () => {
  const r = record()
  r.pending = { action: 'import', beforeProofRef: r.currentProofRef }
  const f = await fixture(r)
  expect(f.controller.snapshot().entries[0].actions).toEqual(['inspect', 'reconcile'])
  await f.controller.run('package_xml:xml-change', 'reconcile')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({ name: 'reconcile_package_xml_change' }),
    expect.any(AbortSignal),
  )
  expect(
    f.executeTool.mock.calls.some((call) => String(call).includes('existing_presentation_page')),
  ).toBe(false)
})
it('blocks stale XML snapshot changes before proposing an action', async () => {
  const f = await fixture()
  f.setRecord({ ...record(), intent: 'changed source' })
  await f.controller.run('package_xml:xml-change', 'resume')
  expect(f.executeTool).not.toHaveBeenCalled()
  expect(f.controller.snapshot().error).toBeDefined()
})
it('removes XML savepoints after paired backup capability disconnects', async () => {
  const f = await fixture()
  f.disconnect()
  await f.controller.run('package_xml:xml-change', 'resume')
  expect(f.executeTool).not.toHaveBeenCalled()
  await f.controller.refresh()
  expect(f.controller.snapshot().entries).toEqual([])
})

it('lists XML history as historical evidence without reading or certifying the current host', async () => {
  const hostRead = vi.fn(async () => {
    throw Error('unexpected_host_read')
  })
  const skill = createPresentationExistingEditingSkill({
    baseline: {} as never,
    baselineAdapter: { readPage: hostRead } as never,
    adapter: {} as never,
    proposals: createStructuredProposalController(),
    documentId: async () => 'doc',
    request: async () => {
      throw Error('unexpected_pc_read')
    },
    listChangeHistory: () => [
      {
        id: 'package_xml:xml-change',
        kind: 'package_xml',
        record: record(),
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
        kind: 'package_xml',
        changeId: 'xml-change',
        sourceKind: 'slide',
        sourceSlideId: 'host-slide',
        state: 'prepared',
        pending: false,
        currentHostVerified: false,
      },
    ],
  })
  expect(hostRead).not.toHaveBeenCalled()
})
