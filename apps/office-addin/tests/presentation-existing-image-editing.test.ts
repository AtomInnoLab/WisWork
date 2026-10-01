import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
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
  readBoundedImage: async (vfs: InMemoryVfs, path: string) => {
    const image = vfs.readBytes(path)
    return {
      base64: Buffer.from(image).toString('base64'),
      mime: 'image/png',
      bytes: image.length,
      width: Buffer.from(image).readUInt32BE(16),
      height: Buffer.from(image).readUInt32BE(20),
    }
  },
}))
const bytes = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const digest = async (base64: string) =>
  Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes(base64))), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('')

async function fixture(
  originalImage = originalPng,
  replacementImage = replacementPng,
  imageDetails?: Pick<PictureSnapshot, 'geometry' | 'name' | 'altTextDescription' | 'shapeIds'>,
) {
  const originalDigest = await digest(originalImage),
    replacementDigest = await digest(replacementImage)
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
    geometry: imageDetails?.geometry ?? { left: 10, top: 20, width: 100, height: 80 },
    rotation: 0,
    name: imageDetails?.name ?? 'Picture',
    altTextTitle: 'title',
    altTextDescription: imageDetails?.altTextDescription ?? 'description',
    zOrderPosition: 0,
    shapeIds: imageDetails?.shapeIds.map((shapeId) => (shapeId === 'old' ? id : shapeId)) ?? [id],
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
        shapes: [
          {
            id: 'old',
            type: 'Image',
            ...(imageDetails?.geometry ?? { left: 10, top: 20, width: 100, height: 80 }),
          },
        ],
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
      base64: originalImage,
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
        const nextId = id === 'old' ? 'new' : id === 'new' ? 'restored' : `${id}-next`
        await onInserted(nextId)
        current = picture(nextId, base64 === replacementImage ? replacementDigest : originalDigest)
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
    save: vi.fn(async (_documentId: string, base64: string) => ({
      attachmentId: await digest(base64),
      sizeBytes: bytes(base64).length,
      mime: 'image/png' as const,
    })),
    load: vi.fn(async (_documentId: string, metadata: { attachmentId: string }): Promise<string> =>
      metadata.attachmentId === originalDigest ? originalImage : replacementImage,
    ),
  }
  let assetBytes: Uint8Array = bytes(replacementImage)
  const proposals = createStructuredProposalController()
  const inspectPage = vi.fn(async (slideId: string) => ({
    slideId,
    shapesTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: originalPng },
  }))
  const create = () =>
    createPresentationExistingImageEditingSkill({
      baseline,
      imageAdapter: adapter,
      imageBackup: backup,
      vfs: { readBytes: () => assetBytes } as unknown as InMemoryVfs,
      proposals,
      inspectPage,
      documentId: binding.documentId,
      readExistingImageChange: binding.readExistingImageChange,
      writeExistingImageChange: binding.writeExistingImageChange,
    })
  let skill = create()
  const call = (name: string, input: Record<string, unknown>) =>
    skill.executeTool({ id: 'call', name, input })
  const confirm = async () => {
    const proposalId = proposals.pending()!.id
    const decision = proposals.waitForDecision(proposalId)
    await proposals.confirm(proposalId)
    return decision
  }
  return {
    call,
    confirm,
    adapter,
    inspectPage,
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

it('captures the exact native page after confirmed image replacement and undo', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  const first = await f.confirm()
  expect(first.status).toBe('confirmed')
  if (first.status !== 'confirmed') throw new Error('not confirmed')
  expect(first.postWrite).toMatchObject({ status: 'captured', pages: [{ slideId: 'slide' }] })
  expect(f.inspectPage).toHaveBeenCalledWith('slide')
  const changeId = JSON.parse(proposed.output).changeId
  expect(f.binding.readExistingImageChange(changeId)?.capture).toMatchObject({
    hostSlideId: 'slide',
  })
  await f.call('undo_existing_presentation_image_change', { change_id: changeId })
  const second = await f.confirm()
  expect(second.status).toBe('confirmed')
  if (second.status !== 'confirmed') throw new Error('not confirmed')
  expect(second.postWrite).toMatchObject({ status: 'captured', pages: [{ slideId: 'slide' }] })
  expect(f.binding.readExistingImageChange(changeId)?.capture).toMatchObject({
    hostSlideId: 'slide',
  })
})

it('keeps the frozen P0-13 picture geometry and neighboring object in the native replacement flow', async () => {
  const material = new URL(
    '../../../docs/product/ppt-benchmark-materials/PPT-P0-13/',
    import.meta.url,
  )
  const zip = await JSZip.loadAsync(
    readFileSync(new URL('wiswork-image-dense-research-draft.pptx', material)),
  )
  const slide = await zip.file('ppt/slides/slide4.xml')!.async('string')
  const pictures = [...slide.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map(([xml]) => xml)
  expect(pictures).toHaveLength(2)
  const left = pictures[0]!
  const right = pictures[1]!
  const original = readFileSync(new URL('images/schematic-07.png', material))
  const replacement = readFileSync(new URL('images/schematic-12.png', material))
  const relId = /<a:blip r:embed="([^"]+)"/.exec(left)![1]
  const rels = await zip.file('ppt/slides/_rels/slide4.xml.rels')!.async('string')
  const media = new RegExp(`<Relationship\\b(?=[^>]*Id="${relId}")[^>]*Target="([^"]+)"`).exec(
    rels,
  )![1]
  expect(
    Buffer.from(await zip.file(`ppt/${media.replace(/^\.\.\//, '')}`)!.async('uint8array')).equals(
      original,
    ),
  ).toBe(true)
  const off = /<a:off x="(\d+)" y="(\d+)"/.exec(left)!
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"/.exec(left)!
  const geometry = {
    left: Number(off[1]) / 12700,
    top: Number(off[2]) / 12700,
    width: Number(ext[1]) / 12700,
    height: Number(ext[2]) / 12700,
  }
  const f = await fixture(original.toString('base64'), replacement.toString('base64'), {
    geometry,
    name: /name="([^"]+)"/.exec(left)![1],
    altTextDescription: /descr="([^"]+)"/.exec(left)![1],
    shapeIds: ['old', /<p:cNvPr id="(\d+)"/.exec(right)![1]],
  })
  const before = f.current()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/schematic-12.png',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  expect(f.adapter.replace).not.toHaveBeenCalled()
  expect((await f.confirm()).status).toBe('confirmed')
  const after = f.current()
  expect(after.shapeId).not.toBe(before.shapeId)
  expect(after.mediaDigest).toBe(await digest(replacement.toString('base64')))
  expect(after.geometry).toEqual(geometry)
  expect(after.shapeIds).toEqual(['new', before.shapeIds[1]])
  expect(after.name).toBe(before.name)
  expect(after.altTextDescription).toBe(before.altTextDescription)
  expect(f.backup.save.mock.calls[0]?.[1]).toBe(original.toString('base64'))
  const changeId = JSON.parse(proposed.output).changeId as string
  expect(f.binding.readExistingImageChange(changeId)?.state).toBe('complete')
  expect(
    (await f.call('undo_existing_presentation_image_change', { change_id: changeId })).isError,
  ).not.toBe(true)
  expect((await f.confirm()).status).toBe('confirmed')
  expect(f.current().mediaDigest).toBe(before.mediaDigest)
  expect(f.current().geometry).toEqual(geometry)
})

it('records a reviewed screenshot and rejects reuse after undo', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  await f.confirm()
  const changeId = JSON.parse(proposed.output).changeId as string
  const captured = await f.call('capture_existing_presentation_image_review', {
    change_id: changeId,
  })
  expect(captured.isError).toBeUndefined()
  const screenshotDigest = JSON.parse(captured.output).screenshotDigest as string
  const reviewed = await f.call('record_existing_presentation_image_review', {
    change_id: changeId,
    screenshot_digest: screenshotDigest,
    status: 'pass',
    notes: 'checked',
  })
  expect(reviewed.isError).toBeUndefined()
  expect(f.binding.readExistingImageChange(changeId)?.review?.status).toBe('pass')
  await f.call('undo_existing_presentation_image_change', { change_id: changeId })
  await f.confirm()
  expect(f.binding.readExistingImageChange(changeId)?.review).toBeUndefined()
  const stale = await f.call('record_existing_presentation_image_review', {
    change_id: changeId,
    screenshot_digest: screenshotDigest,
    status: 'pass',
    notes: 'old shot',
  })
  expect(stale.isError).toBe(true)
})

it('compares the current page with the persisted image capture during inspect', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  await f.confirm()
  const changeId = JSON.parse(proposed.output).changeId as string
  const matched = await f.call('inspect_existing_presentation_image_change', {
    change_id: changeId,
  })
  expect(JSON.parse(matched.output)).toMatchObject({ visualReceipt: 'matched', qaPassed: false })
  f.inspectPage.mockResolvedValue({
    slideId: 'slide',
    shapesTruncated: false,
    screenshot: { mime: 'image/png', base64: replacementPng },
  })
  const different = await f.call('inspect_existing_presentation_image_change', {
    change_id: changeId,
  })
  expect(JSON.parse(different.output)).toMatchObject({
    visualReceipt: 'different',
    qaPassed: false,
  })
})

it('keeps a confirmed image write while reporting unavailable evidence if the host changes during capture', async () => {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  f.inspectPage.mockImplementationOnce(async (slideId) => {
    f.setCurrent(f.picture('new', 'f'.repeat(64)))
    return {
      slideId,
      shapesTruncated: false,
      screenshot: { mime: 'image/png', base64: originalPng },
    }
  })
  const decision = await f.confirm()
  expect(decision).toMatchObject({ status: 'confirmed', postWrite: { status: 'unavailable' } })
  expect(f.binding.readExistingImageChange(JSON.parse(proposed.output).changeId)?.state).toBe(
    'complete',
  )
  expect(
    f.binding.readExistingImageChange(JSON.parse(proposed.output).changeId)?.capture,
  ).toBeUndefined()
})

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
  expect(f.backup.save).toHaveBeenCalledTimes(2)
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

it('durably reapplies after reopen as an independent change with a fresh native ID', async () => {
  const f = await fixture()
  const first = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  await f.confirm()
  const oldId = JSON.parse(first.output).changeId
  expect(f.binding.readExistingImageChange(oldId)?.sourceBackup?.attachmentId).toBe(
    f.replacementDigest,
  )
  await f.call('undo_existing_presentation_image_change', { change_id: oldId })
  await f.confirm()
  const oldRecord = structuredClone(f.binding.readExistingImageChange(oldId))
  f.reopen()
  f.setAsset(new Uint8Array())
  const reapplied = await f.call('reapply_existing_presentation_image_change', { change_id: oldId })
  expect(reapplied.isError).toBeUndefined()
  const newId = JSON.parse(reapplied.output).changeId
  expect(newId).not.toBe(oldId)
  expect(JSON.parse(reapplied.output).reapplies).toBe(oldId)
  expect((await f.confirm()).status).toBe('confirmed')
  expect(f.binding.readExistingImageChange(oldId)).toEqual(oldRecord)
  expect(f.binding.readExistingImageChange(newId)).toMatchObject({
    state: 'complete',
    oldShapeId: oldRecord!.restoredShapeId,
    reapplies: oldId,
  })
  expect(f.backup.save).toHaveBeenCalledTimes(2)
  await f.call('undo_existing_presentation_image_change', { change_id: newId })
  expect((await f.confirm()).status).toBe('confirmed')
  await f.call('reapply_existing_presentation_image_change', { change_id: newId })
  expect((await f.confirm()).status).toBe('confirmed')
})

async function undoneFixture() {
  const f = await fixture()
  const proposed = await f.call('replace_existing_presentation_image', {
    baseline_id: 'baseline',
    slide_id: 'slide',
    shape_id: 'old',
    path: '/image.png',
  })
  await f.confirm()
  const changeId = JSON.parse(proposed.output).changeId
  await f.call('undo_existing_presentation_image_change', { change_id: changeId })
  await f.confirm()
  f.reopen()
  return { f, changeId }
}

it.each(['original', 'source'])(
  'rejects missing or tampered %s backup without recreating it',
  async (kind) => {
    const { f, changeId } = await undoneFixture()
    const before = structuredClone(f.binding.readExistingImageChange(changeId))
    const replaceCount = f.adapter.replace.mock.calls.length
    const load = f.backup.load.getMockImplementation()!
    f.backup.load.mockImplementation(async (doc, metadata) => {
      const target = kind === 'source' ? before!.assetDigest : before!.original.mediaDigest
      return metadata.attachmentId === target ? btoa('tampered') : load(doc, metadata)
    })
    const rejected = await f.call('reapply_existing_presentation_image_change', {
      change_id: changeId,
    })
    expect(rejected).toMatchObject({ isError: true, output: 'presentation_image_backup_invalid' })
    expect(f.adapter.replace).toHaveBeenCalledTimes(replaceCount)
    expect(f.backup.save).toHaveBeenCalledTimes(2)
    expect(f.binding.readExistingImageChange(changeId)).toEqual(before)
    f.backup.load.mockRejectedValue(new Error('presentation_image_backup_unavailable'))
    expect(
      await f.call('reapply_existing_presentation_image_change', { change_id: changeId }),
    ).toMatchObject({ isError: true, output: 'presentation_image_backup_unavailable' })
  },
)

it.each([
  'mediaDigest',
  'name',
  'altTextTitle',
  'altTextDescription',
  'rotation',
  'geometry',
  'shapeIds',
])('rejects manual restored %s conflicts', async (field) => {
  const { f, changeId } = await undoneFixture()
  const changed = structuredClone(f.current())
  Object.assign(changed, {
    [field]:
      field === 'geometry'
        ? { ...changed.geometry, left: 12 }
        : field === 'rotation'
          ? 5
          : field === 'shapeIds'
            ? [changed.shapeId, 'extra']
            : 'changed',
  })
  f.setCurrent(changed)
  expect(
    (await f.call('reapply_existing_presentation_image_change', { change_id: changeId })).isError,
  ).toBe(true)
  expect(f.backup.save).toHaveBeenCalledTimes(2)
})

it('revalidates restored observation and both backups on confirmation', async () => {
  const { f, changeId } = await undoneFixture()
  const proposed = await f.call('reapply_existing_presentation_image_change', {
    change_id: changeId,
  })
  expect(proposed.isError).toBeUndefined()
  f.setCurrent({ ...f.current(), pictureFingerprint: 'e'.repeat(64) })
  await expect(f.confirm()).rejects.toThrow('proposal_stale')
  expect(f.binding.readExistingImageChange(JSON.parse(proposed.output).changeId)).toBeUndefined()
})

it('keeps reapply pending and refuses replay after a lost insertion receipt', async () => {
  const { f, changeId } = await undoneFixture()
  const old = structuredClone(f.binding.readExistingImageChange(changeId))
  const proposed = await f.call('reapply_existing_presentation_image_change', {
    change_id: changeId,
  })
  f.adapter.replace.mockImplementationOnce(async () => {
    f.setCurrent(f.picture('receipt-lost', f.replacementDigest))
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  const newId = JSON.parse(proposed.output).changeId
  expect(f.binding.readExistingImageChange(newId)).toMatchObject({
    state: 'pending',
    reapplies: changeId,
  })
  expect(f.binding.readExistingImageChange(changeId)).toEqual(old)
  f.reopen()
  expect(
    await f.call('resume_existing_presentation_image_change', { change_id: newId }),
  ).toMatchObject({ isError: true, output: 'presentation_existing_image_manual_review' })
})

it.each(['original', 'source'])(
  'rejects %s backup loss after proposal before host write',
  async (kind) => {
    const { f, changeId } = await undoneFixture()
    const old = f.binding.readExistingImageChange(changeId)!
    const proposed = await f.call('reapply_existing_presentation_image_change', {
      change_id: changeId,
    })
    const load = f.backup.load.getMockImplementation()!
    f.backup.load.mockImplementation(async (doc, metadata) => {
      if (
        metadata.attachmentId === (kind === 'source' ? old.assetDigest : old.original.mediaDigest)
      )
        throw new Error('presentation_image_backup_unavailable')
      return load(doc, metadata)
    })
    await expect(f.confirm()).rejects.toThrow('proposal_stale')
    expect(f.binding.readExistingImageChange(JSON.parse(proposed.output).changeId)).toBeUndefined()
    expect(f.backup.save).toHaveBeenCalledTimes(2)
  },
)
