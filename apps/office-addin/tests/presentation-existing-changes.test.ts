import { expect, it, vi } from 'vitest'
import { createPresentationChangesController } from '../src/agent/presentation-changes.js'
import type { PresentationHistoryEntry } from '../src/skills/powerpoint/presentation-change-history.js'
import type { PresentationExistingImageChange } from '../src/skills/powerpoint/presentation-existing-image.js'

it('routes an undone image reapply by its saved ID only when a durable replacement source exists', async () => {
  const picture = (shapeId: string, mediaDigest: string) => ({
    slideId: 'slide',
    shapeId,
    geometry: { left: 0, top: 0, width: 10, height: 10 },
    rotation: 0,
    name: 'Picture',
    altTextTitle: '',
    altTextDescription: '',
    zOrderPosition: 0,
    shapeIds: [shapeId],
    pictureFingerprint: mediaDigest,
    mediaDigest,
  })
  const record: PresentationExistingImageChange = {
    version: 1,
    changeId: 'image-change',
    documentId: 'doc',
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    scope: { slideIds: ['slide'], shapeIds: ['old'] },
    hostSlideId: 'slide',
    oldShapeId: 'old',
    assetDigest: 'b'.repeat(64),
    original: picture('old', 'c'.repeat(64)),
    backup: { attachmentId: 'c'.repeat(64), sizeBytes: 68, mime: 'image/png' },
    sourceBackup: { attachmentId: 'b'.repeat(64), sizeBytes: 68, mime: 'image/png' },
    state: 'undone',
    insertedShapeId: 'new',
    after: picture('new', 'b'.repeat(64)),
    undoBaseline: picture('new', 'b'.repeat(64)),
    restoredShapeId: 'restored',
  }
  const executeTool = vi.fn(async () => ({ output: '{}', summary: 'Done' }))
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      {
        id: 'existing_image:image-change',
        sequence: 1,
        legacy: false,
        kind: 'existing_image',
        record,
      },
    ],
    executeTool,
  })
  await controller.refresh()
  expect(controller.snapshot().entries[0]?.actions).toEqual(['inspect', 'reapply'])
  await controller.run('existing_image:image-change', 'reapply')
  expect(executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'reapply_existing_presentation_image_change',
      input: { change_id: 'image-change' },
    }),
    expect.any(AbortSignal),
  )
  delete record.sourceBackup
  await controller.refresh()
  expect(controller.snapshot().entries[0]?.actions).toEqual(['inspect'])
  executeTool.mockClear()
  await controller.run('existing_image:image-change', 'reapply')
  expect(executeTool).not.toHaveBeenCalled()
})
function setup() {
  let documentId = 'doc',
    available = true
  let history: PresentationHistoryEntry[] = [
    {
      id: 'existing:change',
      sequence: 1,
      legacy: false,
      kind: 'existing',
      agentRunId: 'run-1',
      toolCallId: 'call-1',
      record: {
        version: 1,
        changeId: 'change',
        documentId: 'doc',
        baselineId: 'baseline',
        baselineDigest: 'a'.repeat(64),
        scope: { slideIds: ['slide'], shapeIds: ['shape'] },
        hostSlideId: 'slide',
        shapeId: 'shape',
        shapeType: 'TextBox',
        kind: 'text',
        before: 'Before',
        after: 'After',
        state: 'applied',
      },
    },
  ]
  const executeTool = vi.fn(async () => ({ output: '{}', summary: 'Done' }))
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => available,
    artifact: () => undefined,
    documentId: async () => documentId,
    listChangeHistory: () => history,
    executeTool,
  })
  return {
    controller,
    executeTool,
    switchDoc: () => {
      documentId = 'other'
    },
    disable: () => {
      available = false
    },
    setState: (state: string) => {
      history = structuredClone(history)
      history[0].record.state = state as never
    },
    retainBackup: () => {
      const record = history[0].record as Extract<
        PresentationHistoryEntry,
        { kind: 'existing' }
      >['record']
      record.beforeSlideIds = ['slide']
      record.backup = {
        hostSlideId: 'slide',
        backupId: 'backup',
        sha256: 'a'.repeat(64),
        packageDigest: 'b'.repeat(64),
        sizeBytes: 128,
      }
    },
    corrupt: () => {
      history = [history[0], history[0]]
    },
  }
}
it('shows offline existing history and dispatches by exact change ID without project identity', async () => {
  const s = setup()
  await s.controller.refresh()
  expect(s.controller.snapshot().entries).toEqual([
    expect.objectContaining({
      id: 'existing:change',
      origin: { agentRunId: 'run-1', toolCallId: 'call-1' },
      source: 'existing',
      kind: 'text',
      pageId: 'slide',
      actions: ['inspect', 'undo'],
    }),
  ])
  await s.controller.run('existing:change', 'undo')
  expect(s.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'undo_existing_presentation_change',
      input: { change_id: 'change' },
    }),
    expect.any(AbortSignal),
  )
})
it.each(['pending', 'undo_pending', 'undone'])(
  'shows state-appropriate existing actions for %s',
  async (state) => {
    const s = setup()
    s.setState(state)
    await s.controller.refresh()
    expect(s.controller.snapshot().entries[0]?.actions).toEqual(
      state === 'undone' ? ['inspect'] : ['inspect', 'resume'],
    )
  },
)
it.each(['switchDoc', 'disable', 'corrupt'] as const)(
  'rejects existing action after %s',
  async (key) => {
    const s = setup()
    await s.controller.refresh()
    s[key]()
    await s.controller.run('existing:change', 'undo')
    expect(s.executeTool).not.toHaveBeenCalled()
    expect(s.controller.snapshot().error).toBeTruthy()
  },
)
it('does not publish late success after document switches during an action', async () => {
  const s = setup()
  await s.controller.refresh()
  s.executeTool.mockImplementation(async () => {
    s.switchDoc()
    return { output: '{}', summary: 'Done' }
  })
  await s.controller.run('existing:change', 'undo')
  expect(s.controller.snapshot().error).toBeTruthy()
  expect(s.controller.snapshot().entries).toEqual([])
})
it('rejects record changes during inspection', async () => {
  const s = setup()
  await s.controller.refresh()
  s.executeTool.mockImplementation(async () => {
    s.setState('undone')
    return { output: '{}', summary: 'Done' }
  })
  await s.controller.run('existing:change', 'inspect')
  expect(s.controller.snapshot().error).toBeTruthy()
})
it('accepts tool state transitions and refreshes exact existing actions', async () => {
  const s = setup()
  await s.controller.refresh()
  s.executeTool.mockImplementation(async () => {
    s.setState('undone')
    return { output: '{}', summary: 'Done' }
  })
  await s.controller.run('existing:change', 'undo')
  expect(s.controller.snapshot().error).toBeUndefined()
  expect(s.controller.snapshot().entries[0]?.actions).toEqual(['inspect'])
})
it('clear aborts existing work and drops late results while double click stays single', async () => {
  const s = setup()
  await s.controller.refresh()
  let finish!: (value: { output: string; summary: string }) => void
  s.executeTool.mockImplementation(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const work = s.controller.run('existing:change', 'undo')
  await vi.waitFor(() => expect(s.executeTool).toHaveBeenCalledOnce())
  await s.controller.run('existing:change', 'undo')
  s.controller.clear()
  finish({ output: '{}', summary: 'Done' })
  await work
  expect(s.executeTool).toHaveBeenCalledOnce()
  expect(s.controller.snapshot()).toEqual({ phase: 'idle', entries: [] })
})

it('hides another document history on refresh', async () => {
  const s = setup()
  s.switchDoc()
  await s.controller.refresh()
  expect(s.controller.snapshot().entries).toEqual([])
  expect(s.controller.snapshot().error).toBeUndefined()
})
it('rejects damaged history returned after tool execution', async () => {
  const s = setup()
  await s.controller.refresh()
  s.executeTool.mockImplementation(async () => {
    s.corrupt()
    return { output: '{}', summary: 'Done' }
  })
  await s.controller.run('existing:change', 'undo')
  expect(s.controller.snapshot().error).toBeTruthy()
  expect(s.controller.snapshot().entries).toEqual([])
})

it('offers single reapply only with retained backup and routes the exact change ID', async () => {
  const s = setup()
  s.setState('undone')
  s.retainBackup()
  await s.controller.refresh()
  expect(s.controller.snapshot().entries[0]?.actions).toEqual(['inspect', 'reapply', 'release'])
  await s.controller.run('existing:change', 'reapply')
  expect(s.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'reapply_existing_presentation_change',
      input: { change_id: 'change' },
    }),
    expect.any(AbortSignal),
  )
})
