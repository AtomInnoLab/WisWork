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
