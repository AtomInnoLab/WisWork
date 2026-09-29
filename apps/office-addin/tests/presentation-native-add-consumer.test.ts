import { expect, it, vi } from 'vitest'
import { createPresentationChangesController } from '../src/agent/presentation-changes.js'
import { presentationChangeSetSummary } from '../src/skills/powerpoint/presentation-history'
import { createPresentationExistingBatchEditingSkill } from '../src/skills/powerpoint/presentation-existing-batch-editing'
import type { PresentationNativeAddBatch } from '../src/skills/powerpoint/presentation-existing-batch'

const record: PresentationNativeAddBatch = {
  version: 2,
  kind: 'native_page_add',
  changeId: 'add',
  documentId: 'doc',
  baselineId: 'before',
  baselineDigest: 'a'.repeat(64),
  hostSlideId: 'slide',
  slideIndex: 0,
  beforeSlideIds: ['slide'],
  scope: { slideIds: ['slide'] },
  intent: 'Add a title',
  preserved: ['Original page'],
  validation: ['Native readback'],
  risk: 'high',
  backups: [
    {
      hostSlideId: 'slide',
      backupId: 'backup',
      sha256: 'b'.repeat(64),
      sizeBytes: 100,
      packageDigest: 'a'.repeat(64),
    },
  ],
  operations: [
    {
      op: 'add_text_box',
      slide_index: 0,
      name: 'new-title',
      text: 'Title',
      left: 72,
      top: 72,
      width: 720,
      height: 72,
    },
  ],
  createdShapeIds: [],
  nextIndex: 0,
  state: 'applying',
}
it('summarizes native additions using the actual journal page without inventing per-object host IDs', () => {
  const summary = presentationChangeSetSummary({
    id: 'existing_batch:add',
    kind: 'existing_batch',
    record,
    legacy: false,
    sequence: 1,
  })
  expect(summary).toMatchObject({
    scope: { slideIds: ['slide'] },
    intent: 'Add a title',
    operations: [{ kind: 'existing_batch', pageId: 'slide' }],
  })
})
it.each(['inspect', 'resume', 'undo', 'reapply', 'release'])(
  'rejects a V2 journal in legacy target-batch %s before reading or writing the host',
  async (action) => {
    const readPage = vi.fn(async () => {
      throw new Error('legacy read should not occur')
    })
    const executeDeclarative = vi.fn()
    const options = {
      documentId: async () => 'doc',
      readExistingBatch: () => structuredClone(record),
      writeExistingBatch: vi.fn(),
      baseline: {},
      baselineAdapter: { readPage },
      adapter: { executeDeclarative },
      proposals: {},
      request: vi.fn(),
    } as unknown as Parameters<typeof createPresentationExistingBatchEditingSkill>[0]
    const skill = createPresentationExistingBatchEditingSkill(options)
    const result = await skill.executeTool({
      id: 'old-batch',
      name: `${action}_existing_presentation_batch`,
      input: { change_id: 'add' },
    })
    expect(result).toMatchObject({
      isError: true,
      mutated: false,
      output: 'office_api_unsupported',
    })
    expect(readPage).not.toHaveBeenCalled()
    expect(executeDeclarative).not.toHaveBeenCalled()
    expect(options.writeExistingBatch).not.toHaveBeenCalled()
  },
)

it('shows a pending native addition as historical metadata without offering unconnected actions', async () => {
  const executeTool = vi.fn(async () => ({ output: '{}', summary: 'Historical metadata' }))
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      {
        id: 'existing_batch:add',
        kind: 'existing_batch',
        sequence: 1,
        legacy: false,
        record: structuredClone(record),
      },
    ],
    executeTool,
  })
  await controller.refresh()
  expect(controller.snapshot().entries).toHaveLength(1)
  expect(controller.snapshot().entries[0]).toMatchObject({
    id: 'existing_batch:add',
    kind: 'addition',
    pageId: 'slide',
    state: 'applying',
    cursor: 0,
    operationCount: 1,
    affectedPageCount: 1,
    actions: [],
  })
  expect(controller.snapshot().entries[0].after).toContain('new-title')
  expect(controller.snapshot().entries[0].after).toContain('尚未记录创建身份')
  expect(executeTool).not.toHaveBeenCalled()
})
