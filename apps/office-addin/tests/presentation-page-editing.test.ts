import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
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
