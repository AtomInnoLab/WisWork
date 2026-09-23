import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import { PNG } from 'pngjs'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import type { ImageReplacementRecord } from '../src/skills/powerpoint/presentation-image-replacement-record.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationPageEditingSkill } from '../src/skills/powerpoint/presentation-page-editing.js'
function setup() {
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: 'UEsDBAAAAAA=',
    slideCount: 2,
    pages: [
      { id: 'page1', title: 'First', sourceSlideId: '256#' },
      { id: 'page2', title: 'Second', sourceSlideId: '257#' },
    ],
  }
  const receipt = {
    state: 'pending' as const,
    documentId: 'doc',
    checkpoint: {
      version: 1 as const,
      artifactDigest: createHash('sha256').update(artifact.pptxBase64).digest('hex'),
      sourceSlideIds: ['256#', '257#'],
      baselineSlideIds: ['original'],
      completed: [{ sourceSlideId: '256#', slideId: 'host-42' }],
    },
  }
  let text = 'Before'
  const adapter = {
    listPresentationPageShapes: vi.fn(async () => ({
      slideId: 'host-42',
      shapes: [
        { id: 'shape1', name: 'Title', type: 'TextBox', left: 0, top: 0, width: 100, height: 50 },
      ],
      shapesTruncated: false,
    })),
    readPresentationPageText: vi.fn(async () => ({
      slideId: 'host-42',
      shapeId: 'shape1',
      text,
      paragraphs: [text],
    })),
    editPresentationPageText: vi.fn(
      async (_slide: string, _shape: string, next: string, expected: string) => {
        if (text !== expected) throw new Error('proposal_stale')
        text = next
      },
    ),
  }
  const documentId = vi.fn(async () => 'doc'),
    proposals = createStructuredProposalController()
  const options = {
    available: () => true,
    artifact: () => artifact,
    documentId,
    readReceipt: () => receipt,
    adapter,
    proposals,
  }
  const skill = createPresentationPageEditingSkill(options)
  const read = { id: 'read', name: 'read_presentation_page', input: { page_id: 'page1' } },
    edit = {
      id: 'edit',
      name: 'edit_presentation_page_text',
      input: { page_id: 'page1', shape_id: 'shape1', text: 'After' },
    }
  return {
    options,
    artifact,
    receipt,
    adapter,
    documentId,
    proposals,
    skill,
    read,
    edit,
    setText: (value: string) => {
      text = value
    },
  }
}
it('reads stable host IDs and edits only after a fresh confirmed text comparison', async () => {
  const f = setup()
  expect(JSON.parse((await f.skill.executeTool(f.read)).output)).toMatchObject({
    pageId: 'page1',
    hostSlideId: 'host-42',
    shapes: [{ id: 'shape1' }],
  })
  expect(f.adapter.listPresentationPageShapes).toHaveBeenCalledWith('host-42', undefined)
  expect(await f.skill.executeTool(f.edit)).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
    mutated: false,
  })
  expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledWith(
    'host-42',
    'shape1',
    'After',
    'Before',
    expect.any(AbortSignal),
  )
  expect(f.adapter.readPresentationPageText).toHaveBeenLastCalledWith(
    'host-42',
    'shape1',
    expect.any(AbortSignal),
  )
})
it('rejects pending pages, changed source packages, and unknown input fields', async () => {
  const f = setup()
  expect(await f.skill.executeTool({ ...f.read, input: { page_id: 'page2' } })).toMatchObject({
    isError: true,
    output: 'presentation_page_not_imported',
  })
  expect(
    await f.skill.executeTool({ ...f.read, input: { page_id: 'page1', slide_index: 2 } }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  f.artifact.pptxBase64 += 'AA'
  expect(await f.skill.executeTool(f.read)).toMatchObject({
    isError: true,
    output: 'presentation_page_binding_invalid',
  })
  expect(f.adapter.listPresentationPageShapes).not.toHaveBeenCalled()
})
it('does not overwrite changed text or revive a proposal after clear', async () => {
  const f = setup()
  await f.skill.executeTool(f.edit)
  f.setText('User edit')
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
  await f.skill.executeTool(f.edit)
  f.skill.clear()
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
})
it('rejects foreign document switches and late shape reads after lifecycle clear', async () => {
  const f = setup(),
    original = f.adapter.listPresentationPageShapes.getMockImplementation()!
  f.adapter.listPresentationPageShapes.mockImplementation(async () => {
    f.documentId.mockResolvedValue('other')
    return original()
  })
  expect(await f.skill.executeTool(f.read)).toMatchObject({
    isError: true,
    output: 'presentation_document_changed',
  })
  const g = setup(),
    read = g.adapter.readPresentationPageText.getMockImplementation()!
  g.adapter.readPresentationPageText.mockImplementation(async () => {
    g.skill.clear()
    return read()
  })
  expect(await g.skill.executeTool(g.edit)).toMatchObject({ isError: true, output: 'cancelled' })
  expect(g.proposals.pending()).toBeUndefined()
})
it('rechecks text after mutation hooks and refuses missing targets or failed readback', async () => {
  const f = setup(),
    proposals = createStructuredProposalController(undefined, {
      beforeWrite: async () => {
        f.setText('Concurrent user change')
      },
      afterWrite: () => {},
    })
  const skill = createPresentationPageEditingSkill({ ...f.options, proposals })
  await skill.executeTool(f.edit)
  await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
  const g = setup()
  await g.skill.executeTool(g.edit)
  g.adapter.readPresentationPageText.mockRejectedValue(new Error('office_read_failed'))
  await expect(g.proposals.confirm(g.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(g.adapter.editPresentationPageText).not.toHaveBeenCalled()
  const h = setup()
  h.adapter.editPresentationPageText.mockResolvedValue()
  await h.skill.executeTool(h.edit)
  await expect(h.proposals.confirm(h.proposals.pending()!.id)).rejects.toThrow(
    'office_verify_failed',
  )
})
it('bounds read results and proposals before publication while retaining full expected text', async () => {
  const f = setup()
  f.setText('中'.repeat(12000))
  const next = '新'.repeat(12000)
  expect(
    await f.skill.executeTool({ ...f.edit, input: { ...f.edit.input, text: next } }),
  ).not.toHaveProperty('isError', true)
  const proposal = f.proposals.pending()!
  expect(proposal.preview).toMatchObject({ beforeTruncated: true, beforeLength: 12000 })
  expect(proposal.after).toBe(next)
  expect(new TextEncoder().encode(JSON.stringify(proposal)).byteLength).toBeLessThan(64 * 1024)
  await f.proposals.confirm(proposal.id)
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledWith(
    'host-42',
    'shape1',
    next,
    '中'.repeat(12000),
    expect.any(AbortSignal),
  )
  const g = setup()
  expect(
    await g.skill.executeTool({
      ...g.edit,
      input: { ...g.edit.input, text: '\u0000'.repeat(12000) },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_page_preview_too_large' })
  expect(g.proposals.pending()).toBeUndefined()
  g.adapter.listPresentationPageShapes.mockResolvedValue({
    slideId: 'host-42',
    shapes: [
      {
        id: 'shape1',
        name: 'x'.repeat(65 * 1024),
        type: 'TextBox',
        left: 0,
        top: 0,
        width: 1,
        height: 1,
      },
    ],
    shapesTruncated: false,
  })
  expect(await g.skill.executeTool(g.read)).toMatchObject({
    isError: true,
    output: 'presentation_page_output_too_large',
  })
})
it('fails closed for old receipts and changed mappings and preserves pre-abort behavior', async () => {
  const f = setup(),
    legacy = createPresentationPageEditingSkill({
      ...f.options,
      readReceipt: () => ({ state: 'complete', documentId: 'doc', slideIds: ['host-42'] }),
    })
  expect(await legacy.executeTool(f.read)).toMatchObject({
    isError: true,
    output: 'presentation_page_binding_invalid',
  })
  const original = f.adapter.listPresentationPageShapes.getMockImplementation()!
  f.adapter.listPresentationPageShapes.mockImplementation(async () => {
    f.receipt.checkpoint.completed[0]!.slideId = 'other'
    return original()
  })
  expect(await f.skill.executeTool(f.read)).toMatchObject({
    isError: true,
    output: 'presentation_page_stale',
  })
  const g = setup(),
    controller = new AbortController()
  controller.abort()
  expect(await g.skill.executeTool(g.edit, controller.signal)).toMatchObject({
    isError: true,
    output: 'cancelled',
  })
  expect(g.adapter.readPresentationPageText).not.toHaveBeenCalled()
})
function geometrySetup() {
  const f = setup()
  let geometry = { left: 10, top: 20, width: 100, height: 50 }
  const readPresentationPageGeometry = vi.fn(async () => ({
    slideId: 'host-42',
    shapeId: 'shape1',
    geometry: { ...geometry },
  }))
  const editPresentationPageGeometry = vi.fn(
    async (_slide: string, _shape: string, next: typeof geometry, expected: typeof geometry) => {
      expect(expected).toEqual(geometry)
      geometry = { ...next }
    },
  )
  const adapter = { ...f.adapter, readPresentationPageGeometry, editPresentationPageGeometry }
  const options = { ...f.options, adapter },
    skill = createPresentationPageEditingSkill(options)
  const read = {
    id: 'geo-read',
    name: 'read_presentation_page_geometry',
    input: { page_id: 'page1', shape_id: 'shape1' },
  }
  const edit = {
    id: 'geo-edit',
    name: 'edit_presentation_page_geometry',
    input: {
      page_id: 'page1',
      shape_id: 'shape1',
      geometry: { left: -5, top: 25.5, width: 0, height: 60 },
    },
  }
  return {
    ...f,
    adapter,
    options,
    skill,
    read,
    edit,
    setGeometry: (value: typeof geometry) => {
      geometry = value
    },
  }
}
it('reads and confirms native geometry in points using stable host IDs', async () => {
  const f = geometrySetup()
  expect(f.skill.tools.map((tool) => tool.name)).toContain('edit_presentation_page_geometry')
  expect(JSON.parse((await f.skill.executeTool(f.read)).output)).toMatchObject({
    hostSlideId: 'host-42',
    shapeId: 'shape1',
    unit: 'pt',
    geometry: { left: 10, top: 20, width: 100, height: 50 },
  })
  await f.skill.executeTool(f.edit)
  expect(f.proposals.pending()).toMatchObject({
    preview: { unit: 'pt' },
    before: { left: 10, top: 20, width: 100, height: 50 },
    after: f.edit.input.geometry,
  })
  expect(f.adapter.editPresentationPageGeometry).not.toHaveBeenCalled()
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.adapter.editPresentationPageGeometry).toHaveBeenCalledWith(
    'host-42',
    'shape1',
    f.edit.input.geometry,
    { left: 10, top: 20, width: 100, height: 50 },
    expect.any(AbortSignal),
  )
  expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
})
it('hides unsupported geometry tools and rejects invalid geometry before reading Office', async () => {
  const old = setup()
  expect(old.skill.tools.map((tool) => tool.name)).not.toContain('read_presentation_page_geometry')
  expect(
    await old.skill.executeTool({
      id: 'geo',
      name: 'read_presentation_page_geometry',
      input: { page_id: 'page1', shape_id: 'shape1' },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_unavailable' })
  const f = geometrySetup()
  for (const geometry of [
    { left: 0, top: 0, width: 1 },
    { left: 0, top: 0, width: 1, height: 1, rotation: 1 },
    { left: NaN, top: 0, width: 1, height: 1 },
    { left: 0, top: Infinity, width: 1, height: 1 },
    { left: -100001, top: 0, width: 1, height: 1 },
    { left: 0, top: 0, width: -1, height: 1 },
    null,
  ]) {
    expect(
      await f.skill.executeTool({ ...f.edit, input: { ...f.edit.input, geometry } }),
    ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  }
  expect(f.adapter.readPresentationPageGeometry).not.toHaveBeenCalled()
})
it('uses exact pre-write geometry and accepts only 0.01pt readback rounding', async () => {
  const f = geometrySetup()
  await f.skill.executeTool(f.edit)
  f.setGeometry({ left: 10.001, top: 20, width: 100, height: 50 })
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.editPresentationPageGeometry).not.toHaveBeenCalled()
  const g = geometrySetup()
  g.adapter.editPresentationPageGeometry.mockImplementation(async () =>
    g.setGeometry({ ...g.edit.input.geometry, top: 25.505 }),
  )
  await g.skill.executeTool(g.edit)
  await g.proposals.confirm(g.proposals.pending()!.id)
  const h = geometrySetup()
  h.adapter.editPresentationPageGeometry.mockImplementation(async () =>
    h.setGeometry({ ...h.edit.input.geometry, top: 25.52 }),
  )
  await h.skill.executeTool(h.edit)
  await expect(h.proposals.confirm(h.proposals.pending()!.id)).rejects.toThrow(
    'office_verify_failed',
  )
})
it('rejects wrong geometry objects and missing shapes and checks the mutation hook again', async () => {
  const f = geometrySetup()
  f.adapter.readPresentationPageGeometry.mockResolvedValue({
    slideId: 'wrong',
    shapeId: 'shape1',
    geometry: { left: 0, top: 0, width: 1, height: 1 },
  })
  expect(await f.skill.executeTool(f.read)).toMatchObject({
    isError: true,
    output: 'office_read_failed',
  })
  expect(f.proposals.pending()).toBeUndefined()
  const g = geometrySetup(),
    proposals = createStructuredProposalController(undefined, {
      beforeWrite: async () => g.setGeometry({ left: 11, top: 20, width: 100, height: 50 }),
      afterWrite: () => {},
    })
  const skill = createPresentationPageEditingSkill({ ...g.options, proposals })
  await skill.executeTool(g.edit)
  await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(g.adapter.editPresentationPageGeometry).not.toHaveBeenCalled()
  const h = geometrySetup()
  await h.skill.executeTool(h.edit)
  h.adapter.readPresentationPageGeometry.mockRejectedValue(new Error('office_read_failed'))
  await expect(h.proposals.confirm(h.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(h.adapter.editPresentationPageGeometry).not.toHaveBeenCalled()
})
it('guards geometry awaits against cancellation and context changes and freezes the requested values', async () => {
  const f = geometrySetup(),
    original = f.adapter.readPresentationPageGeometry.getMockImplementation()!
  f.adapter.readPresentationPageGeometry.mockImplementation(async () => {
    f.skill.clear()
    return original()
  })
  expect(await f.skill.executeTool(f.read)).toMatchObject({ isError: true, output: 'cancelled' })
  const g = geometrySetup(),
    read = g.adapter.readPresentationPageGeometry.getMockImplementation()!
  g.adapter.readPresentationPageGeometry.mockImplementation(async () => {
    g.documentId.mockResolvedValue('foreign')
    return read()
  })
  expect(await g.skill.executeTool(g.edit)).toMatchObject({
    isError: true,
    output: 'presentation_document_changed',
  })
  expect(g.proposals.pending()).toBeUndefined()
  const h = geometrySetup()
  await h.skill.executeTool(h.edit)
  const proposed = { ...h.edit.input.geometry }
  h.edit.input.geometry.left = 999
  await h.proposals.confirm(h.proposals.pending()!.id)
  expect(h.adapter.editPresentationPageGeometry).toHaveBeenCalledWith(
    'host-42',
    'shape1',
    proposed,
    expect.any(Object),
    expect.any(AbortSignal),
  )
})

afterEach(() => vi.unstubAllGlobals())
function imageSetup() {
  const f = setup(),
    vfs = new InMemoryVfs(),
    records = new Map<string, ImageReplacementRecord>()
  const png = new PNG({ width: 1, height: 1 })
  png.data.fill(120)
  const bytes = PNG.sync.write(png)
  vfs.writeFile('/home/user/new.png', bytes)
  vi.stubGlobal(
    'createImageBitmap',
    vi.fn(async () => ({ width: 1, height: 1, close: () => {} })),
  )
  const snapshot = {
    slideId: 'host-42',
    shapeId: 'shape1',
    geometry: { left: 1, top: 2, width: 100, height: 50 },
    rotation: 0,
    name: 'Picture',
    altTextTitle: 'alt',
    altTextDescription: 'description',
    zOrderPosition: 0,
    shapeIds: ['shape1', 'title'],
    pictureFingerprint: 'a'.repeat(64),
    mediaDigest: 'b'.repeat(64),
  }
  const imageAdapter = {
    inspect: vi.fn(async () => structuredClone(snapshot)),
    replace: vi.fn(
      async (
        _slide: string,
        _shape: string,
        _base64: string,
        _expected: typeof snapshot,
        onInserted: (id: string) => Promise<void>,
      ) => {
        await onInserted('new-picture')
        return { shapeId: 'new-picture' }
      },
    ),
  }
  const readImageReplacement = (key: string) => records.get(key),
    writeImageReplacement = vi.fn(async (key: string, record: ImageReplacementRecord) => {
      records.set(key, structuredClone(record))
    })
  const options = { ...f.options, vfs, imageAdapter, readImageReplacement, writeImageReplacement },
    skill = createPresentationPageEditingSkill(options)
  const replace = {
    id: 'replace',
    name: 'replace_presentation_page_image',
    input: { page_id: 'page1', shape_id: 'shape1', path: '/home/user/new.png' },
  }
  const status = {
    id: 'status',
    name: 'read_presentation_image_replacement',
    input: { page_id: 'page1', shape_id: 'shape1' },
  }
  return {
    ...f,
    vfs,
    records,
    imageAdapter,
    readImageReplacement,
    writeImageReplacement,
    options,
    skill,
    replace,
    status,
    snapshot,
  }
}
it('confirms image replacement and persists pending, candidate, then complete without replay', async () => {
  const f = imageSetup()
  expect(f.skill.tools.map((tool) => tool.name)).toContain('replace_presentation_page_image')
  await f.skill.executeTool(f.replace)
  expect(f.imageAdapter.replace).not.toHaveBeenCalled()
  f.imageAdapter.replace.mockImplementation(
    async (_slide, _shape, _base64, _expected, onInserted) => {
      expect([...f.records.values()][0]).toMatchObject({ state: 'pending', oldShapeId: 'shape1' })
      await onInserted('new-picture')
      expect([...f.records.values()][0]).toMatchObject({
        state: 'pending',
        newShapeId: 'new-picture',
      })
      return { shapeId: 'new-picture' }
    },
  )
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect([...f.records.values()][0]).toMatchObject({ state: 'complete', newShapeId: 'new-picture' })
  expect(await f.skill.executeTool(f.replace)).toMatchObject({
    output: expect.stringContaining('already_replaced'),
  })
  const result = JSON.parse((await f.skill.executeTool(f.status)).output)
  expect(result).toMatchObject({
    historical: true,
    record: { state: 'complete', newShapeId: 'new-picture' },
  })
  expect(f.imageAdapter.replace).toHaveBeenCalledOnce()
})
it('keeps failed image replacements pending across recreation and rejects changed VFS sources', async () => {
  const f = imageSetup()
  f.imageAdapter.replace.mockRejectedValue(new Error('office_state_uncertain'))
  await f.skill.executeTool(f.replace)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
    'office_state_uncertain',
  )
  expect(await createPresentationPageEditingSkill(f.options).executeTool(f.replace)).toMatchObject({
    isError: true,
    output: 'presentation_image_replacement_uncertain',
  })
  expect(f.imageAdapter.replace).toHaveBeenCalledOnce()
  const g = imageSetup()
  await g.skill.executeTool(g.replace)
  g.vfs.writeFile('/home/user/new.png', new Uint8Array([1, 2, 3]))
  await expect(g.proposals.confirm(g.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(g.writeImageReplacement).not.toHaveBeenCalled()
  expect(g.imageAdapter.replace).not.toHaveBeenCalled()
})
it('does not call image adapter if reservation fails and retains candidate on completion save failure', async () => {
  const f = imageSetup()
  f.writeImageReplacement.mockRejectedValue(new Error('save_failed'))
  await f.skill.executeTool(f.replace)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
  expect(f.imageAdapter.replace).not.toHaveBeenCalled()
  const g = imageSetup(),
    write = g.writeImageReplacement.getMockImplementation()!
  g.writeImageReplacement.mockImplementation(async (key, record) => {
    if (record.state === 'complete') throw new Error('save_failed')
    await write(key, record)
  })
  await g.skill.executeTool(g.replace)
  await expect(g.proposals.confirm(g.proposals.pending()!.id)).rejects.toThrow()
  expect([...g.records.values()][0]).toMatchObject({ state: 'pending', newShapeId: 'new-picture' })
  expect(await g.skill.executeTool(g.replace)).toMatchObject({
    isError: true,
    output: 'presentation_image_replacement_uncertain',
  })
})
it('keeps history readable without browser decode and hides unsupported replacement', async () => {
  const f = imageSetup()
  await f.skill.executeTool(f.replace)
  await f.proposals.confirm(f.proposals.pending()!.id)
  vi.stubGlobal('createImageBitmap', undefined)
  const skill = createPresentationPageEditingSkill({ ...f.options, imageAdapter: undefined })
  expect(skill.tools.map((t) => t.name)).not.toContain('replace_presentation_page_image')
  expect(skill.tools.map((t) => t.name)).toContain('read_presentation_image_replacement')
  expect(JSON.parse((await skill.executeTool(f.status)).output)).toMatchObject({
    historical: true,
    record: { state: 'complete' },
  })
  expect(await skill.executeTool(f.replace)).toMatchObject({
    isError: true,
    output: 'presentation_unavailable',
  })
})
it('rejects changed original pictures, document changes, and lifecycle cancellation before insertion', async () => {
  for (const change of ['picture', 'document', 'clear'] as const) {
    const f = imageSetup()
    await f.skill.executeTool(f.replace)
    if (change === 'picture') f.snapshot.mediaDigest = 'c'.repeat(64)
    if (change === 'document') f.documentId.mockResolvedValue('other-document')
    if (change === 'clear') f.skill.clear()
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    expect(f.writeImageReplacement).not.toHaveBeenCalled()
    expect(f.imageAdapter.replace).not.toHaveBeenCalled()
  }
})
it('retains pending state when candidate persistence fails and rejects forged history scope', async () => {
  const f = imageSetup(),
    write = f.writeImageReplacement.getMockImplementation()!
  f.writeImageReplacement.mockImplementation(async (key, record) => {
    if (record.newShapeId) throw new Error('save_failed')
    await write(key, record)
  })
  await f.skill.executeTool(f.replace)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect([...f.records.values()][0]).toMatchObject({ state: 'pending' })
  expect([...f.records.values()][0].newShapeId).toBeUndefined()
  const [key, record] = [...f.records.entries()][0]
  f.records.set(key, { ...record, documentId: 'another-doc' })
  expect(await f.skill.executeTool(f.status)).toMatchObject({
    isError: true,
    output: 'presentation_image_replacement_invalid',
  })
})
it('does not insert if source changes while the pending reservation is saving', async () => {
  const f = imageSetup(),
    write = f.writeImageReplacement.getMockImplementation()!
  f.writeImageReplacement.mockImplementation(async (key, record) => {
    await write(key, record)
    f.vfs.writeFile('/home/user/new.png', new Uint8Array([1, 2, 3]))
  })
  await f.skill.executeTool(f.replace)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.imageAdapter.replace).not.toHaveBeenCalled()
  expect([...f.records.values()][0]).toMatchObject({ state: 'pending' })
})
async function recoverySetup() {
  const f = imageSetup()
  f.imageAdapter.replace.mockImplementation(async (_a, _b, _c, _d, onInserted) => {
    await onInserted('new-picture')
    throw new Error('office_state_uncertain')
  })
  await f.skill.executeTool(f.replace)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
  const imageAdapter = {
    ...f.imageAdapter,
    inspectRecovery: vi.fn(async () => ({
      status: 'ready_to_finish' as 'ready_to_finish' | 'already_applied' | 'manual_review',
    })),
    finishRecovery: vi.fn(async () => ({ shapeId: 'new-picture' })),
  }
  const options = { ...f.options, vfs: undefined, imageAdapter },
    skill = createPresentationPageEditingSkill(options)
  const inspect = { ...f.status, name: 'inspect_presentation_image_replacement' },
    resume = { ...f.status, name: 'resume_presentation_image_replacement' }
  return { ...f, options, skill, imageAdapter, inspect, resume }
}
it('inspects and resumes a checkpointed replacement without source or browser decode', async () => {
  const f = await recoverySetup()
  expect([...f.records.values()][0]).toMatchObject({ baseline: f.snapshot })
  vi.stubGlobal('createImageBitmap', undefined)
  expect(f.skill.tools.map((t) => t.name)).toContain('resume_presentation_image_replacement')
  expect(JSON.parse((await f.skill.executeTool(f.inspect)).output)).toMatchObject({
    status: 'ready_to_finish',
  })
  await f.skill.executeTool(f.resume)
  expect(f.imageAdapter.finishRecovery).not.toHaveBeenCalled()
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect([...f.records.values()][0].state).toBe('complete')
  expect(f.imageAdapter.finishRecovery).toHaveBeenCalledWith(
    expect.objectContaining({ state: 'pending' }),
    'ready_to_finish',
    expect.any(AbortSignal),
  )
  expect(f.imageAdapter.replace).toHaveBeenCalledOnce()
})
it('does not propose recovery for legacy or manual records and rejects changed inspection', async () => {
  const f = await recoverySetup()
  f.imageAdapter.inspectRecovery.mockResolvedValue({ status: 'manual_review' })
  expect(await f.skill.executeTool(f.resume)).toMatchObject({ isError: true })
  expect(f.proposals.pending()).toBeUndefined()
  f.imageAdapter.inspectRecovery.mockResolvedValue({ status: 'ready_to_finish' })
  await f.skill.executeTool(f.resume)
  f.imageAdapter.inspectRecovery.mockResolvedValue({ status: 'already_applied' })
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.imageAdapter.finishRecovery).not.toHaveBeenCalled()
  const g = await recoverySetup(),
    [key, record] = [...g.records.entries()][0]
  delete record.baseline
  g.records.set(key, record)
  expect(JSON.parse((await g.skill.executeTool(g.inspect)).output)).toMatchObject({
    status: 'manual_review',
  })
  expect(await g.skill.executeTool(g.resume)).toMatchObject({ isError: true })
})
it('retains pending after recovered host succeeds but completion save fails', async () => {
  const f = await recoverySetup()
  f.imageAdapter.inspectRecovery.mockResolvedValue({ status: 'already_applied' })
  f.writeImageReplacement.mockRejectedValue(new Error('save_failed'))
  await f.skill.executeTool(f.resume)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect([...f.records.values()][0].state).toBe('pending')
  expect(f.imageAdapter.finishRecovery).toHaveBeenCalledOnce()
})
it('rejects changed recovery records, documents, and cleared lifecycle without finishing', async () => {
  for (const change of ['record', 'document', 'clear'] as const) {
    const f = await recoverySetup()
    await f.skill.executeTool(f.resume)
    if (change === 'record') {
      const [key, record] = [...f.records.entries()][0]
      f.records.set(key, { ...record, assetDigest: 'c'.repeat(64) })
    }
    if (change === 'document') f.documentId.mockResolvedValue('another-doc')
    if (change === 'clear') f.skill.clear()
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    expect(f.imageAdapter.finishRecovery).not.toHaveBeenCalled()
  }
})
it('checks recovery context after asynchronous inspection and hides legacy adapter recovery tools', async () => {
  const f = await recoverySetup()
  f.imageAdapter.inspectRecovery.mockImplementation(async () => {
    f.documentId.mockResolvedValue('other-doc')
    return { status: 'ready_to_finish' }
  })
  expect(await f.skill.executeTool(f.resume)).toMatchObject({
    isError: true,
    output: 'presentation_document_changed',
  })
  expect(f.proposals.pending()).toBeUndefined()
  const g = imageSetup()
  expect(g.skill.tools.map((t) => t.name)).not.toContain('resume_presentation_image_replacement')
  expect(
    await g.skill.executeTool({ ...g.status, name: 'resume_presentation_image_replacement' }),
  ).toMatchObject({ isError: true, output: 'presentation_unavailable' })
  const h = await recoverySetup()
  expect(
    await h.skill.executeTool({
      ...h.resume,
      input: { ...h.resume.input, path: '/home/user/new.png' },
    }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
})
function productionEditing(f = setup()) {
  const artifact = {
    ...f.artifact,
    pptxBase64: '',
    planRevision: 1,
    pagePptxBase64: ['UEsDBAAAAAA=', 'UEsDBAEAAAA='],
    pages: f.artifact.pages.map((p) => ({ ...p, sourceSlideId: '256#' })),
  }
  const content = JSON.stringify({
    documentId: artifact.documentId,
    projectId: artifact.projectId,
    requestId: artifact.requestId,
    planRevision: artifact.planRevision,
    pages: artifact.pages,
    pagePptxBase64: artifact.pagePptxBase64,
  })
  const receipt = {
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: ['host-first', 'host-42'],
    checkpoint: {
      version: 2 as const,
      artifactDigest: createHash('sha256').update(content).digest('hex'),
      pageIds: ['page1', 'page2'],
      sourceSlideIds: ['256#', '256#'],
      baselineSlideIds: ['original'],
      completed: [
        { sourceSlideId: '256#', slideId: 'host-first' },
        { sourceSlideId: '256#', slideId: 'host-42' },
      ],
    },
  }
  const readReceipt = vi.fn((key: string) =>
    key === 'production/project/request' ? receipt : undefined,
  )
  const options = { ...f.options, artifact: () => artifact, readReceipt },
    skill = createPresentationPageEditingSkill(options)
  return {
    ...f,
    artifact,
    receipt,
    options,
    skill,
    readReceipt,
    edit: { ...f.edit, input: { ...f.edit.input, page_id: 'page2' } },
  }
}
it('edits the second production page by business ID despite duplicate source slide IDs', async () => {
  const f = productionEditing()
  const result = await f.skill.executeTool(f.edit)
  expect(result.isError, result.output).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledWith(
    'host-42',
    'shape1',
    'After',
    'Before',
    expect.any(AbortSignal),
  )
  expect(f.readReceipt).toHaveBeenCalledWith('production/project/request')
})
it('rejects production digest or business order changes and stale bytes after proposal', async () => {
  for (const mode of ['digest', 'order', 'stale']) {
    const f = productionEditing()
    if (mode === 'digest') f.receipt.checkpoint.artifactDigest = 'a'.repeat(64)
    if (mode === 'order') f.receipt.checkpoint.pageIds.reverse()
    const result = await f.skill.executeTool(f.edit)
    if (mode !== 'stale') expect(result.isError).toBe(true)
    else {
      expect(result.isError, result.output).not.toBe(true)
      f.artifact.pagePptxBase64.reverse()
      await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    }
    expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
  }
})
it('uses isolated production image recovery records for the second page and rejects legacy source substitution', async () => {
  const f = imageSetup(),
    p = productionEditing(f)
  const options = { ...f.options, artifact: () => p.artifact, readReceipt: p.readReceipt }
  const skill = createPresentationPageEditingSkill(options),
    call = { ...f.replace, input: { ...f.replace.input, page_id: 'page2' } }
  f.imageAdapter.replace.mockImplementation(async (_a, _b, _c, _d, onInserted) => {
    await onInserted('new-picture')
    throw new Error('office_state_uncertain')
  })
  expect((await skill.executeTool(call)).isError).not.toBe(true)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
    'office_state_uncertain',
  )
  const [key, record] = [...f.records.entries()][0]
  expect(record).toMatchObject({ source: 'production', pageId: 'page2', hostSlideId: 'host-42' })
  const { imageReplacementKey } =
    await import('../src/skills/powerpoint/presentation-image-replacement-record.js')
  expect(key).toBe(await imageReplacementKey('project', 'request', 'page2', 'shape1', 'production'))
  expect(key).not.toBe(await imageReplacementKey('project', 'request', 'page2', 'shape1'))
  const imageAdapter = {
    ...f.imageAdapter,
    inspectRecovery: vi.fn(async () => ({ status: 'already_applied' as const })),
    finishRecovery: vi.fn(async () => ({ shapeId: 'new-picture' })),
  }
  const resumed = createPresentationPageEditingSkill({ ...options, imageAdapter }),
    resume = {
      id: 'resume',
      name: 'resume_presentation_image_replacement',
      input: { page_id: 'page2', shape_id: 'shape1' },
    }
  expect((await resumed.executeTool(resume)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect([...f.records.values()][0]).toMatchObject({ source: 'production', state: 'complete' })
  const legacy = { ...record }
  delete legacy.source
  f.records.set(key, legacy)
  expect(
    await resumed.executeTool({ ...resume, name: 'read_presentation_image_replacement' }),
  ).toMatchObject({ isError: true, output: 'presentation_image_replacement_invalid' })
})
it('keeps legacy image keys unchanged and validates only the explicit production source', async () => {
  const { imageReplacementKey, validateImageReplacementRecord } =
    await import('../src/skills/powerpoint/presentation-image-replacement-record.js')
  expect(await imageReplacementKey('project', 'request', 'page1', 'shape1')).toBe(
    createHash('sha256')
      .update(JSON.stringify(['project', 'request', 'page1', 'shape1']))
      .digest('hex'),
  )
  const record = {
    version: 1,
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pageId: 'page1',
    hostSlideId: 'host',
    oldShapeId: 'old',
    assetDigest: 'a'.repeat(64),
    state: 'pending',
  }
  expect(validateImageReplacementRecord({ ...record, source: 'production' })).toBe(true)
  for (const source of [null, 'legacy', '', false])
    expect(validateImageReplacementRecord({ ...record, source })).toBe(false)
})
it('rejects mixed receipt versions and incomplete production pages before host reads', async () => {
  const f = productionEditing()
  const legacy = { ...f.receipt, checkpoint: { ...f.receipt.checkpoint, version: 1 as const } }
  const mixed = createPresentationPageEditingSkill({ ...f.options, readReceipt: () => legacy })
  expect(await mixed.executeTool(f.edit)).toMatchObject({ isError: true })
  const partial = {
    ...f.receipt,
    state: 'pending' as const,
    slideIds: undefined,
    checkpoint: {
      ...f.receipt.checkpoint,
      completed: f.receipt.checkpoint.completed.slice(0, 1),
      inFlight: { sourceSlideId: '256#' },
    },
  }
  const skill = createPresentationPageEditingSkill({ ...f.options, readReceipt: () => partial })
  expect(await skill.executeTool(f.edit)).toMatchObject({ isError: true })
  expect(f.adapter.readPresentationPageText).not.toHaveBeenCalled()
})
it('checks production content again after asynchronous host reads', async () => {
  const f = productionEditing()
  f.adapter.readPresentationPageText.mockImplementation(async () => {
    f.artifact.planRevision = 2
    return { slideId: 'host-42', shapeId: 'shape1', text: 'Before', paragraphs: ['Before'] }
  })
  expect(await f.skill.executeTool(f.edit)).toMatchObject({
    isError: true,
    output: 'presentation_page_stale',
  })
  expect(f.proposals.pending()).toBeUndefined()
})
function undoSetup() {
  const f = geometrySetup()
  let record:
    | import('../src/skills/powerpoint/presentation-geometry-change.js').PresentationGeometryChange
    | undefined
  const readGeometryChange = () => record,
    writeGeometryChange = vi.fn(
      async (next: NonNullable<typeof record>, expected: typeof record) => {
        expect(record).toEqual(expected)
        record = structuredClone(next)
      },
    )
  const options = { ...f.options, readGeometryChange, writeGeometryChange },
    skill = createPresentationPageEditingSkill(options),
    undo = { id: 'undo', name: 'undo_presentation_geometry_change', input: { page_id: 'page1' } }
  return { ...f, options, skill, undo, readGeometryChange, writeGeometryChange }
}
it('persists the latest geometry before writing and supports confirmed idempotent undo after recreation', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.writeGeometryChange.mock.calls.map(([r]) => r.state)).toEqual(['pending', 'applied'])
  const skill = createPresentationPageEditingSkill(f.options)
  expect((await skill.executeTool(f.undo)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.readGeometryChange()?.state).toBe('undone')
  expect((await skill.executeTool(f.undo)).output).toContain('already_undone')
  expect(f.adapter.editPresentationPageGeometry).toHaveBeenCalledTimes(2)
})
it('rejects manual geometry drift and preserves pending after an uncertain write', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  f.setGeometry({ left: 999, top: 0, width: 1, height: 1 })
  expect(await f.skill.executeTool(f.undo)).toMatchObject({ isError: true })
  const g = undoSetup()
  g.adapter.editPresentationPageGeometry.mockRejectedValue(new Error('office_state_uncertain'))
  await g.skill.executeTool(g.edit)
  await expect(g.proposals.confirm(g.proposals.pending()!.id)).rejects.toThrow()
  expect(g.readGeometryChange()?.state).toBe('pending')
  expect(await g.skill.executeTool(g.edit)).toMatchObject({ isError: true })
})
it('does not write geometry when pending persistence fails or is cancelled while saving', async () => {
  for (const mode of ['failed', 'cancelled']) {
    const f = undoSetup(),
      save = f.writeGeometryChange.getMockImplementation()!
    f.writeGeometryChange.mockImplementation(async (next, expected) => {
      if (mode === 'failed') throw new Error('save_failed')
      await save(next, expected)
      f.skill.clear()
    })
    await f.skill.executeTool(f.edit)
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    expect(f.adapter.editPresentationPageGeometry).not.toHaveBeenCalled()
  }
})
it('retains applied or uncertain undo records on storage failures and uses actual tolerated host value', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  const close = { ...f.edit.input.geometry, left: f.edit.input.geometry.left + 0.005 }
  f.setGeometry(close)
  await f.skill.executeTool(f.undo)
  const save = f.writeGeometryChange.getMockImplementation()!
  f.writeGeometryChange.mockImplementation(async (next, expected) => {
    if (next.state === 'undone') throw new Error('save_failed')
    await save(next, expected)
  })
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect(f.adapter.editPresentationPageGeometry).toHaveBeenLastCalledWith(
    'host-42',
    'shape1',
    expect.any(Object),
    close,
    expect.any(AbortSignal),
  )
  expect(f.readGeometryChange()?.state).toBe('undo_pending')
  expect(await f.skill.executeTool(f.undo)).toMatchObject({ isError: true })
})
it('leaves no-op geometry without a record and rejects changed journal proposals', async () => {
  const f = undoSetup()
  expect(
    (
      await f.skill.executeTool({
        ...f.edit,
        input: { ...f.edit.input, geometry: { left: 10, top: 20, width: 100, height: 50 } },
      })
    ).output,
  ).toContain('unchanged')
  expect(f.writeGeometryChange).not.toHaveBeenCalled()
  await f.skill.executeTool(f.edit)
  const id = f.proposals.pending()!.id
  f.options.readGeometryChange = () => undefined
  const g = undoSetup()
  await g.skill.executeTool(g.edit)
  await g.proposals.confirm(g.proposals.pending()!.id)
  f.options.readGeometryChange = () => g.readGeometryChange()
  await expect(f.proposals.confirm(id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.editPresentationPageGeometry).not.toHaveBeenCalled()
})
it('reads only a matching saved geometry source and keeps historical undone idempotent', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  const read = { ...f.undo, name: 'read_presentation_geometry_change' }
  expect(JSON.parse((await f.skill.executeTool(read)).output)).toMatchObject({
    historical: true,
    record: { state: 'applied' },
  })
  const record = f.readGeometryChange()!
  f.options.readGeometryChange = () => ({ ...record, source: 'production' })
  expect(await f.skill.executeTool(read)).toMatchObject({ isError: true })
})
it('blocks new geometry across projects while a document journal is pending and keeps pending on failed applied save', async () => {
  const f = undoSetup(),
    save = f.writeGeometryChange.getMockImplementation()!
  f.writeGeometryChange.mockImplementation(async (next, expected) => {
    if (next.state === 'applied') throw new Error('save_failed')
    await save(next, expected)
  })
  await f.skill.executeTool(f.edit)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect(f.readGeometryChange()?.state).toBe('pending')
  const record = f.readGeometryChange()!
  f.options.readGeometryChange = () => ({
    ...record,
    projectId: 'another-project',
    source: 'production',
  })
  expect(
    await f.skill.executeTool({
      ...f.edit,
      input: { ...f.edit.input, geometry: { left: 1, top: 2, width: 3, height: 4 } },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_geometry_change_uncertain' })
  expect(f.adapter.editPresentationPageGeometry).toHaveBeenCalledOnce()
})
async function interruptedGeometry(undo = false, applied = false) {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  const record = f.readGeometryChange()!
  const state = undo ? 'undo_pending' : 'pending'
  f.options.readGeometryChange = () => ({ ...record, state }) as typeof record
  let current = { ...record, state } as typeof record
  f.options.readGeometryChange = () => current
  f.options.writeGeometryChange = vi.fn(async (next, expected) => {
    expect(current).toEqual(expected)
    current = structuredClone(next)
  })
  f.setGeometry(
    applied ? (undo ? record.before : record.after) : undo ? record.after : record.before,
  )
  const skill = createPresentationPageEditingSkill(f.options)
  return {
    ...f,
    skill,
    record,
    resume: {
      id: 'resume-geo',
      name: 'resume_presentation_geometry_change',
      input: { page_id: 'page1' },
    },
    inspect: {
      id: 'inspect-geo',
      name: 'inspect_presentation_geometry_change',
      input: { page_id: 'page1' },
    },
    getRecord: () => current,
  }
}
it.each([
  [false, false],
  [false, true],
  [true, false],
  [true, true],
])(
  'recovers geometry pending undo=%s applied=%s only after confirmation',
  async (undo, applied) => {
    const f = await interruptedGeometry(undo, applied)
    const writes = f.adapter.editPresentationPageGeometry.mock.calls.length
    expect(JSON.parse((await f.skill.executeTool(f.inspect)).output).status).toBe(
      applied ? 'already_applied' : 'ready_to_apply',
    )
    expect((await f.skill.executeTool(f.resume)).isError).not.toBe(true)
    await f.proposals.confirm(f.proposals.pending()!.id)
    expect(f.getRecord().state).toBe(undo ? 'undone' : 'applied')
    expect(f.adapter.editPresentationPageGeometry).toHaveBeenCalledTimes(writes + (applied ? 0 : 1))
    expect(JSON.parse((await f.skill.executeTool(f.resume)).output).status).toBe('not_pending')
  },
)
it('requires manual review when pending geometry matches both or neither endpoint', async () => {
  for (const both of [true, false]) {
    const f = await interruptedGeometry()
    if (both) {
      f.getRecord().after = { ...f.record.before, left: f.record.before.left + 0.005 }
      f.setGeometry(f.record.before)
    } else f.setGeometry({ left: 999, top: 9, width: 9, height: 9 })
    expect(JSON.parse((await f.skill.executeTool(f.inspect)).output).status).toBe('manual_review')
    expect(await f.skill.executeTool(f.resume)).toMatchObject({
      isError: true,
      output: 'presentation_geometry_change_manual_review',
    })
    expect(f.proposals.pending()).toBeUndefined()
  }
})
it('rejects tiny observed drift, cancellation and document changes after recovery proposal', async () => {
  for (const mode of ['drift', 'clear', 'document']) {
    const f = await interruptedGeometry()
    await f.skill.executeTool(f.resume)
    if (mode === 'drift') f.setGeometry({ ...f.record.before, left: f.record.before.left + 0.001 })
    if (mode === 'clear') f.skill.clear()
    if (mode === 'document') f.documentId.mockResolvedValue('other')
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    expect(f.getRecord().state).toBe('pending')
    expect(f.adapter.editPresentationPageGeometry).toHaveBeenCalledTimes(1)
  }
})
it('passes quantized actual geometry as expected and leaves pending if completion save fails', async () => {
  const f = await interruptedGeometry()
  const actual = { ...f.record.before, left: f.record.before.left + 0.005 }
  f.setGeometry(actual)
  await f.skill.executeTool(f.resume)
  f.options.writeGeometryChange = vi.fn(async () => {
    throw new Error('save_failed')
  })
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect(f.adapter.editPresentationPageGeometry).toHaveBeenLastCalledWith(
    'host-42',
    'shape1',
    f.record.after,
    actual,
    expect.any(AbortSignal),
  )
  expect(f.getRecord().state).toBe('pending')
})
it('does not read or write the host for terminal geometry inspection', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  f.adapter.readPresentationPageGeometry.mockClear()
  expect(
    JSON.parse(
      (
        await f.skill.executeTool({
          id: 'inspect',
          name: 'inspect_presentation_geometry_change',
          input: { page_id: 'page1' },
        })
      ).output,
    ),
  ).toMatchObject({ status: 'not_pending', historical: true })
  expect(f.adapter.readPresentationPageGeometry).not.toHaveBeenCalled()
})
it('still verifies a committed geometry recovery when Stop races the host write', async () => {
  const f = await interruptedGeometry(),
    write = f.adapter.editPresentationPageGeometry.getMockImplementation()!
  f.adapter.editPresentationPageGeometry.mockImplementation(async (...args) => {
    await write(...args)
    f.proposals.newTurn()
  })
  await f.skill.executeTool(f.resume)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.getRecord().state).toBe('applied')
})
it('rejects recovery classification drift and record changes during host inspection', async () => {
  const f = await interruptedGeometry()
  await f.skill.executeTool(f.resume)
  f.setGeometry(f.record.after)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  const g = await interruptedGeometry(),
    read = g.adapter.readPresentationPageGeometry.getMockImplementation()!
  g.adapter.readPresentationPageGeometry.mockImplementation(async () => {
    const result = await read()
    g.getRecord().changeId = 'changed'
    return result
  })
  expect(await g.skill.executeTool(g.inspect)).toMatchObject({ isError: true })
})
it('describes geometry inspection as read-only and recovery as direction-aware reconciliation', () => {
  const f = undoSetup(),
    inspect = f.skill.tools.find((t) => t.name === 'inspect_presentation_geometry_change')!,
    resume = f.skill.tools.find((t) => t.name === 'resume_presentation_geometry_change')!
  expect(inspect.description).toContain('Read-only')
  expect(resume.description).toContain('pending direction')
  expect(resume.description).toContain('without another write')
  expect(inspect.description).not.toContain('Propose undoing')
})
