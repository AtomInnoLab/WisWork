import { expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationChangesController } from '../src/agent/presentation-changes'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { createPresentationExistingImageEditingSkill } from '../src/skills/powerpoint/presentation-existing-image-editing'
import type {
  PictureSnapshot,
  ImageRecoveryStatus,
} from '../src/skills/powerpoint/browser-presentation-image-adapter'
import type { PresentationBaselineSkill } from '../src/skills/powerpoint/presentation-baseline'
import type { InMemoryVfs } from '../src/skills/shared/vfs'

const originalPng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
const replacementPng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
vi.mock('../src/skills/shared/import-media', () => ({
  MAX_IMPORT_BYTES: 2 * 1024 * 1024,
  supportsBrowserMediaValidation: () => true,
  readBoundedImage: async () => ({
    base64: replacementPng,
    mime: 'image/png',
    bytes: 68,
    width: 1,
    height: 1,
  }),
}))
const bytes = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const digest = async (base64: string) =>
  Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes(base64))), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('')

async function fixture() {
  const originalDigest = await digest(originalPng),
    replacementDigest = await digest(replacementPng)
  const settings = new Map<string, string>()
  const save = vi.fn(async () => {})
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => settings.get(key),
      set: (key, value) => {
        settings.set(key, value)
      },
      save,
      location: () => 'file://existing-image.pptx',
    },
    () => 'doc',
  )
  const documentId = await binding.documentId()
  const picture = (id: string, mediaDigest: string): PictureSnapshot => ({
    slideId: 'slide',
    shapeId: id,
    geometry: { left: 10, top: 20, width: 100, height: 80 },
    rotation: 0,
    name: 'Picture',
    altTextTitle: 'title',
    altTextDescription: 'description',
    zOrderPosition: 0,
    shapeIds: [id],
    pictureFingerprint: mediaDigest,
    mediaDigest,
  })
  let current = picture('old', originalDigest)
  const baselineSnapshot = {
    version: 1,
    baselineId: 'baseline',
    documentId,
    contentDigest: 'a'.repeat(64),
    scope: { kind: 'selected', slideIds: ['slide'], shapeIds: ['old'] },
    pages: [
      {
        slideId: 'slide',
        shapes: [{ id: 'old', type: 'Image', left: 10, top: 20, width: 100, height: 80 }],
      },
    ],
  }
  const baseline = {
    snapshot: () => structuredClone(baselineSnapshot),
    executeTool: async () => ({
      output: JSON.stringify({ unchanged: true }),
      mutated: false,
      summary: 'checked',
    }),
  } as unknown as PresentationBaselineSkill
  const adapter = {
    inspect: vi.fn(async (_slide: string, id: string) => {
      if (id !== current.shapeId) throw new Error('office_read_failed')
      return structuredClone(current)
    }),
    captureOriginal: vi.fn(async () => ({
      snapshot: structuredClone(current),
      base64: originalPng,
    })),
    replace: vi.fn(
      async (
        _slide: string,
        id: string,
        base64: string,
        expected: PictureSnapshot,
        onInserted: (id: string) => Promise<void>,
      ) => {
        if (id !== current.shapeId || JSON.stringify(expected) !== JSON.stringify(current))
          throw new Error('office_concurrent_change')
        const nextId = id === 'old' ? 'new' : 'restored'
        await onInserted(nextId)
        current = picture(nextId, base64 === replacementPng ? replacementDigest : originalDigest)
        return { shapeId: nextId }
      },
    ),
    inspectRecovery: vi.fn(async (): Promise<ImageRecoveryStatus> => ({ status: 'manual_review' })),
    finishRecovery: vi.fn(async (): Promise<{ shapeId: string }> => {
      throw new Error('not_expected')
    }),
  }
  const backup = {
    available: () => true,
    save: vi.fn(async () => ({
      attachmentId: originalDigest,
      sizeBytes: bytes(originalPng).length,
      mime: 'image/png' as const,
    })),
    load: vi.fn(async () => originalPng),
  }
  let assetBytes: Uint8Array = bytes(replacementPng)
  const proposals = createStructuredProposalController()
  const create = () =>
    createPresentationExistingImageEditingSkill({
      baseline,
      imageAdapter: adapter,
      imageBackup: backup,
      vfs: { readBytes: () => assetBytes } as unknown as InMemoryVfs,
      proposals,
      documentId: binding.documentId,
      readExistingImageChange: binding.readExistingImageChange,
      writeExistingImageChange: binding.writeExistingImageChange,
    })
  let skill = create()
  const call = (name: string, input: Record<string, unknown>) =>
    skill.executeTool({ id: 'call', name, input })
  const confirm = () => proposals.confirm(proposals.pending()!.id)
  return {
    call,
    confirm,
    adapter,
    backup,
    binding,
    save,
    picture,
    replacementDigest,
    setAsset: (value: Uint8Array) => {
      assetBytes = value
    },
    setCurrent: (value: PictureSnapshot) => {
      current = value
    },
    reopen: () => {
      skill.clear()
      skill = create()
    },
    current: () => current,
  }
}

it('backs up before a confirmed native replacement, then undoes from a reopened savepoint', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  expect(f.adapter.replace).not.toHaveBeenCalled()
  expect(f.backup.save).not.toHaveBeenCalled()
  expect(f.binding.listChangeHistory()).toHaveLength(0)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  expect(f.binding.readExistingImageChange(changeId)).toMatchObject({
    state: 'complete',
    insertedShapeId: 'new',
  })
  expect(f.current().shapeId).toBe('new')
  expect(f.backup.save).toHaveBeenCalledOnce()
  f.reopen()
  const undo = await f.call('undo_existing_presentation_image_change', { change_id: changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect(f.binding.readExistingImageChange(changeId)).toMatchObject({
    state: 'undone',
    restoredShapeId: 'restored',
  })
  expect(f.current().shapeId).toBe('restored')
})

it('does not insert when the durable savepoint fails', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.confirm()).rejects.toThrow()
  expect(f.adapter.replace).not.toHaveBeenCalled()
  expect(f.binding.listChangeHistory()).toHaveLength(0)
})

it('rejects source drift during original-image backup before the host insertion', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.backup.save.mockImplementationOnce(async () => {
    f.setAsset(bytes(originalPng))
    return {
      attachmentId: await digest(originalPng),
      sizeBytes: bytes(originalPng).length,
      mime: 'image/png' as const,
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.adapter.replace).not.toHaveBeenCalled()
  expect(f.binding.listChangeHistory()).toHaveLength(0)
})

it('resumes only a recorded candidate after an interrupted native insertion', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  f.adapter.replace.mockImplementationOnce(async (_slide, _id, _base64, _expected, onInserted) => {
    f.setCurrent(f.picture('new', f.replacementDigest))
    await onInserted('new')
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.binding.readExistingImageChange(changeId)).toMatchObject({
    state: 'pending',
    insertedShapeId: 'new',
  })
  f.reopen()
  f.adapter.inspectRecovery.mockResolvedValue({ status: 'already_applied' })
  f.adapter.finishRecovery.mockResolvedValue({ shapeId: 'new' })
  const resume = await f.call('resume_existing_presentation_image_change', { change_id: changeId })
  expect(resume.isError, resume.output).not.toBe(true)
  await f.confirm()
  expect(f.adapter.replace).toHaveBeenCalledTimes(1)
  expect(f.binding.readExistingImageChange(changeId)).toMatchObject({
    state: 'complete',
    insertedShapeId: 'new',
  })
})

it('refuses to replay insertion when a pending image has no recorded candidate', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  f.adapter.replace.mockRejectedValueOnce(new Error('office_state_uncertain'))
  await expect(f.confirm()).rejects.toThrow()
  expect(f.binding.readExistingImageChange(changeId)).toMatchObject({ state: 'pending' })
  f.reopen()
  const resume = await f.call('resume_existing_presentation_image_change', { change_id: changeId })
  expect(resume).toMatchObject({
    isError: true,
    output: 'presentation_existing_image_manual_review',
  })
  expect(f.adapter.replace).toHaveBeenCalledTimes(1)
})

it('shows a native image replacement in the offline change workbench', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  await f.confirm()
  f.reopen()
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    artifact: () => undefined,
    documentId: f.binding.documentId,
    listChangeHistory: f.binding.listChangeHistory,
    executeTool: (call) => f.call(call.name, call.input),
  })
  await controller.refresh()
  const row = controller.snapshot().entries[0]
  expect(row).toMatchObject({
    source: 'existing_image',
    kind: 'image',
    pageId: 'slide',
    state: 'complete',
    actions: ['inspect', 'undo'],
  })
  expect(row.before).toContain('原图已持久备份')
  await controller.run(row.id, 'undo')
  await f.confirm()
  expect(f.binding.listChangeHistory()[0].record.state).toBe('undone')
})
