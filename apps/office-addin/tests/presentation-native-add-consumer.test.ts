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

it('routes an enabled native-addition workbench inspection to the dedicated read-only tool', async () => {
  const executeTool = vi.fn(async () => ({
    output: JSON.stringify({
      visualQaVerified: false,
      operationCount: 1,
      nextIndex: 0,
      observation: { status: 'complete', completedCount: 1 },
    }),
    summary: 'Read-only inspection',
  }))
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeAdditionAvailable: () => true,
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
  expect(controller.snapshot().entries[0].actions).toEqual(['inspect', 'resume'])
  expect(executeTool).not.toHaveBeenCalled()
  await controller.run('existing_batch:add', 'inspect')
  expect(controller.snapshot().notice).toContain('已核对新增对象 1/1；持久回执 0/1')
  expect(controller.snapshot().notice).toContain('不代表视觉或专业 QA 通过')
  expect(executeTool).toHaveBeenCalledOnce()
  expect(executeTool).toHaveBeenCalledWith(
    expect.objectContaining({ name: 'inspect_slide_ir_addition', input: { change_id: 'add' } }),
    expect.any(AbortSignal),
  )
})

it('prepares a native original-page restore, reads a fresh exact baseline and stages only a confirmation proposal', async () => {
  const history = [
    {
      id: 'existing_batch:add',
      kind: 'existing_batch' as const,
      sequence: 1,
      legacy: false,
      record: structuredClone(record),
    },
  ]
  const hostWrite = vi.fn()
  const executeTool = vi.fn(async (call: { name: string; input: Record<string, unknown> }) => {
    if (call.name === 'prepare_existing_presentation_original_page_restore')
      return {
        output: JSON.stringify({
          path: '/home/user/presentation-original-restore-source.pptx',
          slideId: 'slide',
          sourceKind: 'batch',
          sourceChangeId: 'add',
          packageDigest: record.baselineDigest,
          nextTool: 'stage_existing_presentation_page_change',
          nextInput: {
            path: '/home/user/presentation-original-restore-source.pptx',
            slide_id: 'slide',
            restore_source_kind: 'batch',
            restore_source_change_id: 'add',
          },
        }),
        mutated: false,
        summary: 'Prepared',
      }
    if (call.name === 'read_presentation_baseline')
      return {
        output: JSON.stringify({
          baselineId: 'fresh',
          documentId: 'doc',
          scope: { slideIds: ['slide'] },
          context: { slideIds: ['slide'] },
          pages: [{ slideId: 'slide' }],
          coverage: { pagePackages: 'read' },
          contentDigest: 'c'.repeat(64),
          qaPassed: false,
        }),
        mutated: false,
        summary: 'Fresh baseline',
      }
    if (call.name === 'stage_existing_presentation_page_change')
      return {
        output: JSON.stringify({
          proposalId: 'proposal',
          status: 'awaiting_confirmation',
          changeId: 'restore',
          state: 'pending',
          oldSlideId: 'slide',
        }),
        mutated: false,
        summary: 'Stage proposal',
      }
    hostWrite()
    throw Error('unexpected mutation')
  })
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeRestorationAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => history,
    executeTool,
  })
  await controller.refresh()
  expect(controller.snapshot().entries[0].actions).toEqual(['undo'])
  await controller.run('existing_batch:add', 'undo')
  expect(executeTool.mock.calls.map(([call]) => call.name)).toEqual([
    'prepare_existing_presentation_original_page_restore',
    'read_presentation_baseline',
    'stage_existing_presentation_page_change',
  ])
  expect(executeTool.mock.calls[2]![0].input).toEqual({
    baseline_id: 'fresh',
    path: '/home/user/presentation-original-restore-source.pptx',
    slide_id: 'slide',
    restore_source_kind: 'batch',
    restore_source_change_id: 'add',
  })
  expect(hostWrite).not.toHaveBeenCalled()
  expect(history[0]!.record.state).toBe('applying')
  expect(controller.snapshot().notice).toContain('恢复提案')
})

function finalizationHistory() {
  const native: PresentationNativeAddBatch = { ...structuredClone(record), inFlightIndex: 0 }
  const page: import('../src/skills/powerpoint/presentation-existing-page').PresentationExistingPageChange =
    {
      version: 1,
      changeId: 'restore',
      documentId: 'doc',
      baselineId: 'fresh',
      baselineDigest: 'c'.repeat(64),
      scope: { slideIds: ['slide'] },
      oldSlideId: 'slide',
      beforeSlideIds: ['slide'],
      originalPackageDigest: 'd'.repeat(64),
      replacementPackageDigest: native.baselineDigest,
      sourceSlideId: '256#',
      backup: { backupId: 'edited-backup', sha256: 'e'.repeat(64), sizeBytes: 100 },
      sourceBackup: { backupId: 'source-copy', sha256: native.backups[0]!.sha256, sizeBytes: 100 },
      state: 'applied',
      newSlideId: 'restored',
      restores: {
        sourceKind: 'batch',
        sourceChangeId: 'add',
        sourceHostSlideId: 'slide',
        originalBackupId: 'backup',
        originalPackageDigest: native.baselineDigest,
      },
    }
  const history: import('../src/skills/powerpoint/presentation-change-history').PresentationHistoryEntry[] =
    [
      {
        id: 'existing_batch:add',
        kind: 'existing_batch',
        sequence: 1,
        legacy: false,
        record: native,
      },
      {
        id: 'existing_page:restore',
        kind: 'existing_page',
        sequence: 2,
        legacy: false,
        record: page,
      },
    ]
  return { native, page, history }
}
it('offers exact unique restore finalization and reflects only the tool-written terminal journal', async () => {
  const f = finalizationHistory(),
    executeTool = vi.fn(async () => {
      delete f.native.inFlightIndex
      f.native.state = 'undone'
      f.native.restoredSlideId = 'restored'
      return {
        output: JSON.stringify({
          changeId: 'add',
          state: 'undone',
          restoredSlideId: 'restored',
          historicalOnly: true,
          visualQaVerified: false,
        }),
        mutated: false,
        summary: 'Reconciled',
      }
    })
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeRestorationAvailable: () => true,
    nativeRestoreFinalizationAvailable: () => true,
    nativeReleaseAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => f.history,
    executeTool,
  })
  await controller.refresh()
  expect(controller.snapshot().entries.find((e) => e.id === 'existing_batch:add')?.actions).toEqual(
    ['finalize'],
  )
  expect(executeTool).not.toHaveBeenCalled()
  await controller.run('existing_batch:add', 'finalize')
  expect(executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'finalize_slide_ir_addition_restore',
      input: { change_id: 'add', restoration_change_id: 'restore' },
    }),
    expect.any(AbortSignal),
  )
  expect(controller.snapshot().entries.find((e) => e.id === 'existing_batch:add')).toMatchObject({
    state: 'undone',
    pageId: 'restored',
    actions: ['release'],
  })
  expect(controller.snapshot().error).toBeUndefined()
  expect(controller.snapshot().notice).toContain('不代表页面 QA 通过')
})
it.each([
  'ambiguous',
  'not-applied',
  'source-sha',
  'source-host',
  'original-backup',
  'before-order',
])('does not offer finalization for %s restoration evidence', async (scenario) => {
  const f = finalizationHistory()
  if (scenario === 'ambiguous')
    f.history.push({
      id: 'existing_page:other',
      kind: 'existing_page',
      sequence: 3,
      legacy: false,
      record: { ...f.page, changeId: 'other' },
    })
  if (scenario === 'not-applied') f.page.state = 'staged'
  if (scenario === 'source-sha') f.page.sourceBackup!.sha256 = 'f'.repeat(64)
  if (scenario === 'source-host') {
    f.page.oldSlideId = 'other'
    f.page.beforeSlideIds = ['other']
    f.page.scope.slideIds = ['other']
    f.page.restores!.sourceHostSlideId = 'other'
  }
  if (scenario === 'original-backup') f.page.restores!.originalBackupId = 'other'
  if (scenario === 'before-order') f.page.beforeSlideIds = ['slide', 'other']
  const executeTool = vi.fn(async () => ({ output: '{}', summary: 'Unused' })),
    controller = createPresentationChangesController({
      available: () => false,
      existingAvailable: () => true,
      nativeRestoreFinalizationAvailable: () => true,
      artifact: () => undefined,
      documentId: async () => 'doc',
      listChangeHistory: () => f.history,
      executeTool,
    })
  await controller.refresh()
  expect(
    controller.snapshot().entries.find((e) => e.id === 'existing_batch:add')?.actions ?? [],
  ).not.toContain('finalize')
  await controller.run('existing_batch:add', 'finalize')
  expect(executeTool).not.toHaveBeenCalled()
})
it('does not fabricate undone state from a successful-looking finalization response without a durable journal', async () => {
  const f = finalizationHistory(),
    executeTool = vi.fn(async () => ({
      output: JSON.stringify({
        changeId: 'add',
        state: 'undone',
        restoredSlideId: 'restored',
        historicalOnly: true,
        visualQaVerified: false,
      }),
      mutated: false,
      summary: 'Claimed',
    })),
    controller = createPresentationChangesController({
      available: () => false,
      existingAvailable: () => true,
      nativeRestoreFinalizationAvailable: () => true,
      artifact: () => undefined,
      documentId: async () => 'doc',
      listChangeHistory: () => f.history,
      executeTool,
    })
  await controller.refresh()
  await controller.run('existing_batch:add', 'finalize')
  expect(f.native.state).toBe('applying')
  expect(controller.snapshot().error).toBeDefined()
})

it.each([
  'cancel',
  'record-changed',
  'document-changed',
  'prepare-error',
  'wrong-next-slide',
  'wrong-source',
])('stops native restoration before baseline/stage on %s', async (scenario) => {
  let documentId = 'doc'
  const history = [
    {
      id: 'existing_batch:add',
      kind: 'existing_batch' as const,
      sequence: 1,
      legacy: false,
      record: structuredClone(record),
    },
  ]
  const executeTool = vi.fn(async () => {
    if (scenario === 'cancel') controller.clear()
    if (scenario === 'record-changed') history[0]!.record.intent = 'Changed by another action'
    if (scenario === 'document-changed') documentId = 'other-document'
    const nextInput = {
      path: '/home/user/presentation-original-restore-source.pptx',
      slide_id: scenario === 'wrong-next-slide' ? 'foreign' : 'slide',
      restore_source_kind: 'batch',
      restore_source_change_id: 'add',
    }
    return {
      output: JSON.stringify({
        path: nextInput.path,
        slideId: 'slide',
        sourceKind: 'batch',
        sourceChangeId: scenario === 'wrong-source' ? 'foreign' : 'add',
        packageDigest: record.baselineDigest,
        nextTool: 'stage_existing_presentation_page_change',
        nextInput,
      }),
      isError: scenario === 'prepare-error',
      mutated: false,
      summary: 'Prepared',
    }
  })
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeRestorationAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => documentId,
    listChangeHistory: () => history,
    executeTool,
  })
  await controller.refresh()
  await controller.run('existing_batch:add', 'undo')
  expect(executeTool).toHaveBeenCalledOnce()
  expect(history[0]!.record.state).toBe('applying')
  if (scenario !== 'cancel') expect(controller.snapshot().error).toBeDefined()
})

it('offers another restore after a discarded staging attempt and finalizes one applied restore beside discarded history', async () => {
  const f = finalizationHistory()
  f.page.state = 'discarded'
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeRestorationAvailable: () => true,
    nativeRestoreFinalizationAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => f.history,
    executeTool: vi.fn(async () => ({ output: '{}', summary: 'Unused' })),
  })
  await controller.refresh()
  expect(controller.snapshot().entries.find((e) => e.id === 'existing_batch:add')?.actions).toEqual(
    ['undo'],
  )
  f.history.push({
    id: 'existing_page:applied',
    kind: 'existing_page',
    sequence: 3,
    legacy: false,
    record: { ...f.page, changeId: 'applied', state: 'applied' },
  })
  await controller.refresh()
  expect(controller.snapshot().entries.find((e) => e.id === 'existing_batch:add')?.actions).toEqual(
    ['finalize'],
  )
})
it('stops restoration if a same-source staging journal appears during preparation', async () => {
  const f = finalizationHistory(),
    history = [f.history[0]!]
  f.page.state = 'staged'
  const executeTool = vi.fn(async () => {
    history.push(f.history[1]!)
    return {
      output: JSON.stringify({
        path: '/home/user/presentation-original-restore-source.pptx',
        slideId: 'slide',
        sourceKind: 'batch',
        sourceChangeId: 'add',
        packageDigest: record.baselineDigest,
        nextTool: 'stage_existing_presentation_page_change',
        nextInput: {
          path: '/home/user/presentation-original-restore-source.pptx',
          slide_id: 'slide',
          restore_source_kind: 'batch',
          restore_source_change_id: 'add',
        },
      }),
      mutated: false,
      summary: 'Prepared',
    }
  })
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeRestorationAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => history,
    executeTool,
  })
  await controller.refresh()
  await controller.run('existing_batch:add', 'undo')
  expect(executeTool).toHaveBeenCalledOnce()
  expect(controller.snapshot().error).toBeDefined()
})
