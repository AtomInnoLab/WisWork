import { expect, it, vi } from 'vitest'
import { createPresentationChangesController } from '../src/agent/presentation-changes.js'
import type { CompiledPresentationArtifact } from '../src/skills/powerpoint/presentation-delivery.js'
import type { PresentationGeometryChange } from '../src/skills/powerpoint/presentation-geometry-change.js'
import type { PresentationExistingPageChange } from '../src/skills/powerpoint/presentation-existing-page.js'
import type { PresentationExistingChartChange } from '../src/skills/powerpoint/presentation-existing-chart.js'
const artifact: CompiledPresentationArtifact = {
  documentId: 'doc',
  projectId: 'project',
  requestId: 'request',
  pptxBase64: 'base64',
  slideCount: 1,
  pages: [{ id: 'page', title: 'Title', sourceSlideId: 'slide' }],
}
async function setup() {
  const digest = Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode('base64'))),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
  let record: PresentationGeometryChange | undefined = {
    version: 1,
    changeId: 'change',
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    artifactDigest: digest,
    pageId: 'page',
    hostSlideId: 'host',
    shapeId: 'shape',
    before: { left: 1, top: 2, width: 3, height: 4 },
    after: { left: 2, top: 2, width: 3, height: 4 },
    state: 'applied',
  }
  let current = artifact,
    documentId = 'doc'
  const executeTool = vi.fn(async () => ({ output: '<unsafe>payload</unsafe>', summary: 'Done' }))
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => current,
    documentId: async () => documentId,
    readGeometryChange: () => record,
    executeTool,
  })
  await controller.refresh()
  return {
    controller,
    executeTool,
    change: (value: Partial<PresentationGeometryChange>) => {
      record = { ...record!, ...value }
    },
    switchTask: () => {
      current = { ...artifact, requestId: 'other' }
    },
    switchDoc: () => {
      documentId = 'other'
    },
  }
}
it('reads only current records and clones snapshots; routes undo through the tool', async () => {
  const { controller, executeTool } = await setup()
  const s = controller.snapshot()
  expect(s.entries).toHaveLength(1)
  s.entries[0].actions.length = 0
  expect(controller.snapshot().entries[0].actions).toEqual(['undo'])
  expect(executeTool).not.toHaveBeenCalled()
  await controller.run(s.entries[0].id, 'undo')
  expect(executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'undo_presentation_geometry_change',
      input: { project_id: 'project', page_id: 'page' },
    }),
    expect.any(AbortSignal),
  )
  expect(JSON.stringify(controller.snapshot())).not.toContain('unsafe')
})
it('shows existing-page identity diff offline and routes commit by exact change ID', async () => {
  const record: PresentationExistingPageChange = {
    version: 1,
    changeId: 'native-page',
    documentId: 'doc',
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    scope: { slideIds: ['old'] },
    oldSlideId: 'old',
    beforeSlideIds: ['old', 'other'],
    originalPackageDigest: 'b'.repeat(64),
    replacementPackageDigest: 'c'.repeat(64),
    sourceSlideId: '256#',
    backup: { backupId: 'backup', sha256: 'd'.repeat(64), sizeBytes: 120 },
    state: 'staged',
    newSlideId: 'new',
  }
  const executeTool = vi.fn(async () => ({ output: '{}', summary: 'proposed' }))
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      { id: 'existing_page:native-page', kind: 'existing_page', sequence: 1, legacy: false, record },
    ],
    listExistingPageBackups: async () => [{ backupId: 'backup', status: 'ready', hostSlideId: 'old', slideIds: ['old', 'other'], sha256: 'd'.repeat(64), sizeBytes: 120 }],
    executeTool,
  })
  await controller.refresh()
  const row = controller.snapshot().entries[0]!
  expect(row).toMatchObject({ source: 'existing_page', kind: 'page', actions: ['inspect', 'commit', 'discard'] })
  expect(row.before).toContain('old')
  expect(row.after).toContain('new')
  expect(controller.snapshot().backupAudit).toEqual({ active: 1, unmatched: 0 })
  await controller.run(row.id, 'commit')
  expect(executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'commit_existing_presentation_page_change',
      input: { change_id: 'native-page' },
    }),
    expect.any(AbortSignal),
  )
  executeTool.mockResolvedValueOnce({ output: JSON.stringify({ inspection: { status: 'staged' }, visualReceipts: [
    { hostSlideId: 'old', status: 'matched' }, { hostSlideId: 'new', status: 'different' },
  ] }), summary: 'checked' })
  await controller.run(row.id, 'inspect')
  expect(controller.snapshot().notice).toContain('当前截图与历史回执不同')
})
it('offers one chart backup release after cancellation and accepts the receipt update', async () => {
  let record: PresentationExistingChartChange = {
    version: 1, changeId: 'chart-release', documentId: 'doc', oldSlideId: 'old', shapeId: '7',
    slideIndex: 0, beforeSlideIds: ['old'], beforePackageDigest: 'a'.repeat(64),
    afterPackageDigest: 'b'.repeat(64), backup: { backupId: 'backup', sha256: 'c'.repeat(64), sizeBytes: 120 },
    state: 'cancelled',
  }
  let inventoryAvailable = true
  let inventoryDigest = 'c'.repeat(64)
  const executeTool = vi.fn(async () => ({ output: '{}', mutated: false, summary: 'proposed' }))
  const controller = createPresentationChangesController({
    available: () => false, existingAvailable: () => true, artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [{ id: 'existing_chart:chart-release', kind: 'existing_chart', sequence: 1, legacy: false, record }],
    listExistingPageBackups: async () => {
      if (!inventoryAvailable) throw new Error('pc_offline')
      return [
        { backupId: 'backup', status: 'ready', hostSlideId: 'old', slideIds: ['old'], sha256: inventoryDigest, sizeBytes: 120 },
        { backupId: 'unmatched', status: 'ready', hostSlideId: 'old', slideIds: ['old'], sha256: 'c'.repeat(64), sizeBytes: 120 },
      ]
    },
    executeTool,
  })
  await controller.refresh()
  const row = controller.snapshot().entries[0]!
  expect(row.actions).toEqual(['inspect', 'release'])
  expect(controller.snapshot().backupAudit).toEqual({ active: 2, unmatched: 1 })
  inventoryDigest = 'd'.repeat(64)
  await controller.refresh()
  expect(controller.snapshot().backupAudit).toEqual({ active: 2, unmatched: 2 })
  inventoryDigest = 'c'.repeat(64)
  await controller.run(row.id, 'release')
  expect(executeTool).toHaveBeenCalledWith(expect.objectContaining({
    name: 'release_slide_chart_values_change', input: { change_id: 'chart-release' },
  }), expect.any(AbortSignal))
  expect(controller.snapshot().notice).toContain('确认后执行')
  expect(controller.snapshot().entries[0]?.actions).toEqual(['inspect', 'release'])
  record = { ...record, backupReleasedAt: '2026-09-24T09:00:00.000Z' }
  await controller.refresh()
  expect(controller.snapshot().entries[0]?.actions).toEqual(['inspect'])
  expect(controller.snapshot().error).toBeUndefined()
  inventoryAvailable = false
  await controller.refresh()
  expect(controller.snapshot().entries).toHaveLength(1)
  expect(controller.snapshot().backupAudit).toBeUndefined()
})
it('rejects changed fingerprints and invalid actions', async () => {
  const { controller, executeTool, change } = await setup()
  const id = controller.snapshot().entries[0].id
  await controller.run(id, 'commit')
  change({ shapeId: 'other' })
  await controller.run(id, 'undo')
  expect(executeTool).not.toHaveBeenCalled()
  expect(controller.snapshot().error).toBeTruthy()
})
it.each(['switchTask', 'switchDoc'] as const)('rejects %s before action', async (key) => {
  const s = await setup()
  const id = s.controller.snapshot().entries[0].id
  s[key]()
  await s.controller.run(id, 'undo')
  expect(s.executeTool).not.toHaveBeenCalled()
})
it('reports malformed records, hides other requests and digests', async () => {
  const s = await setup()
  s.change({ state: 'bad' as never })
  await s.controller.refresh()
  expect(s.controller.snapshot().error).toBeTruthy()
  s.change({ state: 'applied', requestId: 'other' })
  await s.controller.refresh()
  expect(s.controller.snapshot().entries).toEqual([])
  s.change({ requestId: 'request', artifactDigest: 'a'.repeat(64) })
  await s.controller.refresh()
  expect(s.controller.snapshot().entries).toEqual([])
})
it('clear invalidates pending reads', async () => {
  let release!: (v: string) => void
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => artifact,
    documentId: () =>
      new Promise((r) => {
        release = r
      }),
    executeTool: vi.fn(),
  })
  const pending = controller.refresh()
  controller.clear()
  release('doc')
  await pending
  expect(controller.snapshot()).toEqual({ phase: 'idle', entries: [] })
})
it('single action ignores double clicks and clear cancels late results', async () => {
  const s = await setup()
  let finish!: (v: { output: string; summary: string }) => void
  s.executeTool.mockImplementation(
    () =>
      new Promise((r) => {
        finish = r
      }),
  )
  const id = s.controller.snapshot().entries[0].id
  const first = s.controller.run(id, 'undo')
  await vi.waitFor(() => expect(s.executeTool).toHaveBeenCalledTimes(1))
  await s.controller.run(id, 'undo')
  s.controller.clear()
  finish({ output: 'ok', summary: 'Done' })
  await first
  expect(s.controller.snapshot()).toEqual({ phase: 'idle', entries: [] })
  expect(s.executeTool).toHaveBeenCalledTimes(1)
})
it('rejects source mismatch and missing current pages', async () => {
  const s = await setup()
  s.change({ source: 'production' })
  await s.controller.refresh()
  expect(s.controller.snapshot().entries).toEqual([])
  s.change({ source: undefined, pageId: 'missing' })
  await s.controller.refresh()
  expect(s.controller.snapshot().entries).toEqual([])
})
it('shows safe errors for thrown and error tool outcomes', async () => {
  const s = await setup()
  s.executeTool.mockRejectedValue(new Error('<script>secret</script>'))
  await s.controller.run(s.controller.snapshot().entries[0].id, 'undo')
  expect(s.controller.snapshot().error).toBeTruthy()
  expect(JSON.stringify(s.controller.snapshot())).not.toContain('secret')
})
it('refresh notification inside action is nonblocking and picks up terminal records', async () => {
  const s = await setup()
  s.executeTool.mockImplementation(async () => {
    s.change({ state: 'undone' })
    await s.controller.refresh()
    return { output: 'ok', summary: 'Done' }
  })
  await s.controller.run(s.controller.snapshot().entries[0].id, 'undo')
  expect(s.controller.snapshot().entries[0].actions).toEqual([])
})
it('task switch during action aborts and suppresses the old result', async () => {
  const s = await setup()
  let finish!: (v: { output: string; summary: string }) => void
  s.executeTool.mockImplementation(
    () =>
      new Promise((r) => {
        finish = r
      }),
  )
  const pending = s.controller.run(s.controller.snapshot().entries[0].id, 'undo')
  await vi.waitFor(() => expect(s.executeTool).toHaveBeenCalledOnce())
  s.switchTask()
  await s.controller.refresh()
  finish({ output: 'ok', summary: 'Done' })
  await pending
  expect(s.controller.snapshot().entries).toEqual([])
  expect(s.controller.snapshot().notice).toBeUndefined()
})
it('older refresh cannot replace a newer task snapshot', async () => {
  let release!: (v: string) => void
  let first = true,
    current = artifact
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => current,
    documentId: () =>
      first
        ? ((first = false),
          new Promise((r) => {
            release = r
          }))
        : Promise.resolve('doc'),
    executeTool: vi.fn(),
  })
  const older = controller.refresh()
  current = { ...artifact, requestId: 'other' }
  await controller.refresh()
  release('doc')
  await older
  expect(controller.snapshot().requestId).toBe('other')
})
it.each(['pending', 'applied', 'undo_pending', 'undone'] as const)(
  'maps text %s and invokes exact tool',
  async (state) => {
    const digest = Array.from(
      new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode('base64'))),
      (b) => b.toString(16).padStart(2, '0'),
    ).join('')
    const record = {
      version: 1 as const,
      changeId: 'text',
      documentId: 'doc',
      projectId: 'project',
      requestId: 'request',
      artifactDigest: digest,
      pageId: 'page',
      hostSlideId: 'host',
      shapeId: 'shape',
      before: 'before',
      after: 'after',
      state,
    }
    const executeTool = vi.fn(async () => ({ output: 'ok', summary: 'Done' }))
    const controller = createPresentationChangesController({
      available: () => true,
      artifact: () => artifact,
      documentId: async () => 'doc',
      readTextChange: () => record,
      executeTool,
    })
    await controller.refresh()
    const item = controller.snapshot().entries[0]
    expect(item.actions).toEqual(
      state === 'applied' ? ['undo'] : state === 'undone' ? [] : ['inspect', 'resume'],
    )
    for (const action of item.actions) {
      await controller.run(item.id, action)
      expect(executeTool).toHaveBeenLastCalledWith(
        expect.objectContaining({
          name: `${action}_presentation_text_change`,
          input: { project_id: 'project', page_id: 'page' },
        }),
        expect.any(AbortSignal),
      )
    }
  },
)
it('image pending exposes inspection only without a recovery baseline; complete has no undo', async () => {
  let record = {
    version: 1 as const,
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    hostSlideId: 'host',
    oldShapeId: 'old',
    assetDigest: 'a'.repeat(64),
    state: 'pending' as 'pending' | 'complete',
    newShapeId: undefined as string | undefined,
  }
  const executeTool = vi.fn(async () => ({ output: 'ok', summary: 'Done' }))
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    listImageReplacements: () => [record],
    executeTool,
  })
  await controller.refresh()
  const row = controller.snapshot().entries[0]
  expect(row.actions).toEqual(['inspect'])
  await controller.run(row.id, 'inspect')
  expect(executeTool).toHaveBeenLastCalledWith(
    expect.objectContaining({
      name: 'inspect_presentation_image_replacement',
      input: { project_id: 'project', page_id: 'page', shape_id: 'old' },
    }),
    expect.any(AbortSignal),
  )
  record = { ...record, state: 'complete', newShapeId: 'new' }
  await controller.refresh()
  expect(controller.snapshot().entries[0].actions).toEqual([])
})
it('inspection presents only whitelisted recovery meaning', async () => {
  const s = await setup()
  s.change({ state: 'pending' })
  await s.controller.refresh()
  s.executeTool.mockResolvedValue({
    output: JSON.stringify({ status: 'manual_review', reason: '<secret>' }),
    summary: 'Done',
  })
  await s.controller.run(s.controller.snapshot().entries[0].id, 'inspect')
  expect(s.controller.snapshot().notice).toContain('需要人工检查')
  expect(JSON.stringify(s.controller.snapshot())).not.toContain('secret')
})
it.each([2, 33])('rejects duplicate or excessive image records (%s)', async (count) => {
  const record = {
    version: 1 as const,
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    hostSlideId: 'host',
    oldShapeId: 'old',
    assetDigest: 'a'.repeat(64),
    state: 'pending' as const,
  }
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    listImageReplacements: () =>
      Array.from({ length: count }, (_, i) => ({
        ...record,
        oldShapeId: count === 2 ? 'old' : `shape-${i}`,
      })),
    executeTool: vi.fn(),
  })
  await controller.refresh()
  expect(controller.snapshot().error).toBeTruthy()
  expect(controller.snapshot().entries).toEqual([])
})
it.each([
  ['pending', ['inspect']],
  ['inserted', ['inspect', 'resume']],
  ['staged', ['inspect', 'commit', 'discard']],
  ['discard_pending', ['inspect', 'discard']],
  ['discarded', []],
  ['commit_pending', ['inspect', 'commit']],
  ['applied', ['inspect', 'undo']],
  ['undo_pending', ['inspect', 'undo']],
  ['restore_inserted', ['inspect', 'undo']],
  ['undone', []],
] as const)('maps page %s and routes existing tools', async (state, actions) => {
  const production = {
    ...artifact,
    pptxBase64: '',
    planRevision: 1,
    pagePptxBase64: [Buffer.from('PK\x03\x04revision').toString('base64')],
    pages: [{ id: 'page', title: 'Page', sourceSlideId: '256#' }],
  }
  const receipt = (slideId: string) => ({
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: [slideId],
    checkpoint: {
      version: 2 as const,
      artifactDigest: 'a'.repeat(64),
      sourceSlideIds: ['256#'],
      pageIds: ['page'],
      baselineSlideIds: ['original'],
      completed: [{ sourceSlideId: '256#', slideId }],
    },
  })
  const record = {
    version: 1 as const,
    changeId: 'replacement',
    documentId: 'doc',
    projectId: 'project',
    parentRequestId: 'request',
    requestId: 'child',
    pageId: 'page',
    backupId: 'backup',
    parentArtifactDigest: 'a'.repeat(64),
    backupDigest: 'b'.repeat(64),
    originalPackageDigest: 'c'.repeat(64),
    replacementPackageDigest: 'd'.repeat(64),
    sourceSlideId: '256#',
    oldSlideId: 'old',
    beforeSlideIds: ['old'],
    state,
    ...(state !== 'pending' ? { newSlideId: 'new' } : {}),
    ...(['commit_pending', 'applied', 'undo_pending', 'restore_inserted', 'undone'].includes(state)
      ? { parentReceipt: receipt('old'), childReceipt: receipt('new') }
      : {}),
    ...(['restore_inserted', 'undone'].includes(state) ? { restoredSlideId: 'restored' } : {}),
  }
  let current: CompiledPresentationArtifact = production
  const executeTool = vi.fn(async () => ({
    output: JSON.stringify({ inspection: { status: 'staged' } }),
    summary: 'Done',
  }))
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => current,
    documentId: async () => 'doc',
    readPageReplacement: () => record,
    executeTool,
  })
  await controller.refresh()
  expect(controller.snapshot().error).toBeUndefined()
  const row = controller.snapshot().entries[0]
  expect(row.actions).toEqual(actions)
  for (const action of row.actions) {
    await controller.run(row.id, action)
    expect(executeTool).toHaveBeenLastCalledWith(
      expect.objectContaining({
        name: `${action}_presentation_page_replacement`,
        input: { project_id: 'project', change_id: 'replacement' },
      }),
      expect.any(AbortSignal),
    )
  }
  if (state === 'undone') {
    expect(row.after).toContain(`恢复页面：restored\n恢复包摘要：${'c'.repeat(64)}`)
    expect(row.after).toContain(`页面：new\n包摘要：${'d'.repeat(64)}`)
  }
  current = { ...production, requestId: 'child' }
  await controller.refresh()
  expect(controller.snapshot().entries).toHaveLength(1)
  current = artifact
  await controller.refresh()
  expect(controller.snapshot().entries).toEqual([])
})

it.each(['complete', 'undo_pending', 'undone'] as const)(
  'routes backed-up image %s actions and labels restored bytes accurately',
  async (state) => {
    const baseline = {
      slideId: 'host',
      shapeId: 'old',
      geometry: { left: 0, top: 0, width: 10, height: 10 },
      rotation: 0,
      name: 'Picture',
      altTextTitle: '',
      altTextDescription: '',
      zOrderPosition: 0,
      shapeIds: ['old'],
      pictureFingerprint: 'a'.repeat(64),
      mediaDigest: 'b'.repeat(64),
    }
    const after = {
      ...baseline,
      shapeId: 'new',
      shapeIds: ['new'],
      pictureFingerprint: 'c'.repeat(64),
      mediaDigest: 'd'.repeat(64),
    }
    const record = {
      version: 1 as const,
      documentId: 'doc',
      projectId: 'project',
      requestId: 'request',
      pageId: 'page',
      hostSlideId: 'host',
      oldShapeId: 'old',
      newShapeId: 'new',
      assetDigest: after.mediaDigest,
      state,
      baseline,
      after,
      backup: { attachmentId: baseline.mediaDigest, sizeBytes: 100, mime: 'image/png' as const },
      ...(state !== 'complete' ? { undoBaseline: after, restoredShapeId: 'restored' } : {}),
    }
    const executeTool = vi.fn(async () => ({ output: 'ok', summary: 'Done' }))
    const controller = createPresentationChangesController({
      available: () => true,
      artifact: () => artifact,
      documentId: async () => 'doc',
      listImageReplacements: () => [record],
      executeTool,
    })
    await controller.refresh()
    expect(controller.snapshot().error).toBeUndefined()
    const item = controller.snapshot().entries[0]
    expect(item.actions).toEqual(
      state === 'complete' ? ['undo'] : state === 'undo_pending' ? ['inspect', 'resume'] : [],
    )
    expect(item.before).toContain(baseline.mediaDigest)
    if (state === 'undone')
      expect(item.after).toContain(`恢复图片：restored\n恢复资源摘要：${baseline.mediaDigest}`)
    for (const action of item.actions) {
      await controller.run(item.id, action)
      expect(executeTool).toHaveBeenLastCalledWith(
        expect.objectContaining({
          name: `${action}_presentation_image_replacement`,
          input: { project_id: 'project', page_id: 'page', shape_id: 'old' },
        }),
        expect.any(AbortSignal),
      )
    }
  },
)

it('lists ordered historical text records and routes the selected change ID', async () => {
  const digest = Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode('base64'))),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
  const base = {
    version: 1 as const,
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    artifactDigest: digest,
    pageId: 'page',
    hostSlideId: 'host',
    shapeId: 'shape',
    state: 'applied' as const,
  }
  const older = { ...base, changeId: 'old', before: 'a', after: 'b' }
  const newer = { ...base, changeId: 'new', before: 'b', after: 'c' }
  const executeTool = vi.fn(async () => ({ output: 'ok', summary: 'Done' }))
  const controller = createPresentationChangesController({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      { id: 'text:old', sequence: 1, legacy: true, kind: 'text', record: older },
      { id: 'text:new', sequence: 2, legacy: false, kind: 'text', record: newer },
    ],
    executeTool,
  })
  await controller.refresh()
  const items = controller.snapshot().entries
  expect(items.map((e) => e.id)).toEqual(['text:new', 'text:old'])
  expect(items[1].legacy).toBe(true)
  expect(items[1].changeSet?.scope).toEqual({ slideIds: ['host'], shapeIds: ['shape'] })
  await controller.run('text:old', 'undo')
  expect(executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'undo_presentation_text_change',
      input: { project_id: 'project', page_id: 'page', change_id: 'old' },
    }),
    expect.any(AbortSignal),
  )
})
