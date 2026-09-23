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
