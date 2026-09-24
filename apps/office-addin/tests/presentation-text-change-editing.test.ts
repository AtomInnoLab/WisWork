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
function undoSetup() {
  const f = setup()
  let record:
    | import('../src/skills/powerpoint/presentation-text-change.js').PresentationTextChange
    | undefined
  const readTextChange = () => record,
    writeTextChange = vi.fn(async (next: NonNullable<typeof record>, expected: typeof record) => {
      expect(record).toEqual(expected)
      record = structuredClone(next)
    })
  const options = { ...f.options, readTextChange, writeTextChange },
    skill = createPresentationPageEditingSkill(options),
    undo = { id: 'undo', name: 'undo_presentation_text_change', input: { page_id: 'page1' } }
  return { ...f, options, skill, undo, readTextChange, writeTextChange }
}
it('persists the latest text before writing and supports confirmed idempotent undo after recreation', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.writeTextChange.mock.calls.map(([r]) => r.state)).toEqual(['pending', 'applied'])
  const skill = createPresentationPageEditingSkill(f.options)
  expect((await skill.executeTool(f.undo)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.readTextChange()?.state).toBe('undone')
  expect((await skill.executeTool(f.undo)).output).toContain('already_undone')
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledTimes(2)
})
it('rejects manual text drift and preserves pending after an uncertain write', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  f.setText('manual edit')
  expect(await f.skill.executeTool(f.undo)).toMatchObject({ isError: true })
  const g = undoSetup()
  g.adapter.editPresentationPageText.mockRejectedValue(new Error('office_state_uncertain'))
  await g.skill.executeTool(g.edit)
  await expect(g.proposals.confirm(g.proposals.pending()!.id)).rejects.toThrow()
  expect(g.readTextChange()?.state).toBe('pending')
  expect(await g.skill.executeTool(g.edit)).toMatchObject({ isError: true })
})
it('does not write text when pending persistence fails or is cancelled while saving', async () => {
  for (const mode of ['failed', 'cancelled']) {
    const f = undoSetup(),
      save = f.writeTextChange.getMockImplementation()!
    f.writeTextChange.mockImplementation(async (next, expected) => {
      if (mode === 'failed') throw new Error('save_failed')
      await save(next, expected)
      f.skill.clear()
    })
    await f.skill.executeTool(f.edit)
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
  }
})
it('retains applied or uncertain undo records on storage failures and uses exact host value', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  const close = f.edit.input.text
  f.setText(close)
  await f.skill.executeTool(f.undo)
  const save = f.writeTextChange.getMockImplementation()!
  f.writeTextChange.mockImplementation(async (next, expected) => {
    if (next.state === 'undone') throw new Error('save_failed')
    await save(next, expected)
  })
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect(f.adapter.editPresentationPageText).toHaveBeenLastCalledWith(
    'host-42',
    'shape1',
    expect.any(String),
    close,
    expect.any(AbortSignal),
  )
  expect(f.readTextChange()?.state).toBe('undo_pending')
  expect(await f.skill.executeTool(f.undo)).toMatchObject({ isError: true })
})
it('leaves no-op text without a record and rejects changed journal proposals', async () => {
  const f = undoSetup()
  expect(
    (
      await f.skill.executeTool({
        ...f.edit,
        input: { ...f.edit.input, text: 'Before' },
      })
    ).output,
  ).toContain('unchanged')
  expect(f.writeTextChange).not.toHaveBeenCalled()
  await f.skill.executeTool(f.edit)
  const id = f.proposals.pending()!.id
  f.options.readTextChange = () => undefined
  const g = undoSetup()
  await g.skill.executeTool(g.edit)
  await g.proposals.confirm(g.proposals.pending()!.id)
  f.options.readTextChange = () => g.readTextChange()
  await expect(f.proposals.confirm(id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.editPresentationPageText).not.toHaveBeenCalled()
})
it('reads only a matching saved text source and keeps historical undone idempotent', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  const read = { ...f.undo, name: 'read_presentation_text_change' }
  expect(JSON.parse((await f.skill.executeTool(read)).output)).toMatchObject({
    historical: true,
    record: { state: 'applied' },
  })
  const record = f.readTextChange()!
  f.options.readTextChange = () => ({ ...record, source: 'production' })
  expect(await f.skill.executeTool(read)).toMatchObject({ isError: true })
})
it('blocks new text across projects while a document journal is pending and keeps pending on failed applied save', async () => {
  const f = undoSetup(),
    save = f.writeTextChange.getMockImplementation()!
  f.writeTextChange.mockImplementation(async (next, expected) => {
    if (next.state === 'applied') throw new Error('save_failed')
    await save(next, expected)
  })
  await f.skill.executeTool(f.edit)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect(f.readTextChange()?.state).toBe('pending')
  const record = f.readTextChange()!
  f.options.readTextChange = () => ({
    ...record,
    projectId: 'another-project',
    source: 'production',
  })
  expect(
    await f.skill.executeTool({
      ...f.edit,
      input: { ...f.edit.input, text: 'new' },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_text_change_uncertain' })
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledOnce()
})
async function interruptedText(undo = false, applied = false) {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  const record = f.readTextChange()!
  const state = undo ? 'undo_pending' : 'pending'
  f.options.readTextChange = () => ({ ...record, state }) as typeof record
  let current = { ...record, state } as typeof record
  f.options.readTextChange = () => current
  f.options.writeTextChange = vi.fn(async (next, expected) => {
    expect(current).toEqual(expected)
    current = structuredClone(next)
  })
  f.setText(applied ? (undo ? record.before : record.after) : undo ? record.after : record.before)
  const skill = createPresentationPageEditingSkill(f.options)
  return {
    ...f,
    skill,
    record,
    resume: {
      id: 'resume-geo',
      name: 'resume_presentation_text_change',
      input: { page_id: 'page1' },
    },
    inspect: {
      id: 'inspect-geo',
      name: 'inspect_presentation_text_change',
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
])('recovers text pending undo=%s applied=%s only after confirmation', async (undo, applied) => {
  const f = await interruptedText(undo, applied)
  const writes = f.adapter.editPresentationPageText.mock.calls.length
  expect(JSON.parse((await f.skill.executeTool(f.inspect)).output).status).toBe(
    applied ? 'already_applied' : 'ready_to_apply',
  )
  expect((await f.skill.executeTool(f.resume)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.getRecord().state).toBe(undo ? 'undone' : 'applied')
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledTimes(writes + (applied ? 0 : 1))
  expect(JSON.parse((await f.skill.executeTool(f.resume)).output).status).toBe('not_pending')
})
it('requires manual review when pending text matches both or neither endpoint', async () => {
  for (const both of [true, false]) {
    const f = await interruptedText()
    if (both) {
      f.getRecord().after = f.record.before
      f.setText(f.record.before)
    } else f.setText('manual')
    expect(JSON.parse((await f.skill.executeTool(f.inspect)).output).status).toBe('manual_review')
    expect(await f.skill.executeTool(f.resume)).toMatchObject({
      isError: true,
      output: 'presentation_text_change_manual_review',
    })
    expect(f.proposals.pending()).toBeUndefined()
  }
})
it('rejects tiny observed drift, cancellation and document changes after recovery proposal', async () => {
  for (const mode of ['drift', 'clear', 'document']) {
    const f = await interruptedText()
    await f.skill.executeTool(f.resume)
    if (mode === 'drift') f.setText(f.record.before + ' ')
    if (mode === 'clear') f.skill.clear()
    if (mode === 'document') f.documentId.mockResolvedValue('other')
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
    expect(f.getRecord().state).toBe('pending')
    expect(f.adapter.editPresentationPageText).toHaveBeenCalledTimes(1)
  }
})
it('passes exact actual text as expected and leaves pending if completion save fails', async () => {
  const f = await interruptedText()
  const actual = f.record.before
  f.setText(actual)
  await f.skill.executeTool(f.resume)
  f.options.writeTextChange = vi.fn(async () => {
    throw new Error('save_failed')
  })
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('save_failed')
  expect(f.adapter.editPresentationPageText).toHaveBeenLastCalledWith(
    'host-42',
    'shape1',
    f.record.after,
    actual,
    expect.any(AbortSignal),
  )
  expect(f.getRecord().state).toBe('pending')
})
it('does not read or write the host for terminal text inspection', async () => {
  const f = undoSetup()
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  f.adapter.readPresentationPageText.mockClear()
  expect(
    JSON.parse(
      (
        await f.skill.executeTool({
          id: 'inspect',
          name: 'inspect_presentation_text_change',
          input: { page_id: 'page1' },
        })
      ).output,
    ),
  ).toMatchObject({ status: 'not_pending', historical: true })
  expect(f.adapter.readPresentationPageText).not.toHaveBeenCalled()
})
it('still verifies a committed text recovery when Stop races the host write', async () => {
  const f = await interruptedText(),
    write = f.adapter.editPresentationPageText.getMockImplementation()!
  f.adapter.editPresentationPageText.mockImplementation(async (...args) => {
    await write(...args)
    f.proposals.newTurn()
  })
  await f.skill.executeTool(f.resume)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.getRecord().state).toBe('applied')
})
it('rejects recovery classification drift and record changes during host inspection', async () => {
  const f = await interruptedText()
  await f.skill.executeTool(f.resume)
  f.setText(f.record.after)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  const g = await interruptedText(),
    read = g.adapter.readPresentationPageText.getMockImplementation()!
  g.adapter.readPresentationPageText.mockImplementation(async () => {
    const result = await read()
    g.getRecord().changeId = 'changed'
    return result
  })
  expect(await g.skill.executeTool(g.inspect)).toMatchObject({ isError: true })
})
it('describes text inspection as read-only and recovery as direction-aware reconciliation', () => {
  const f = undoSetup(),
    inspect = f.skill.tools.find((t) => t.name === 'inspect_presentation_text_change')!,
    resume = f.skill.tools.find((t) => t.name === 'resume_presentation_text_change')!
  expect(inspect.description).toContain('Read-only')
  expect(resume.description).toContain('pending direction')
  expect(resume.description).toContain('without another write')
  expect(inspect.description).not.toContain('Propose undoing')
})

it.each(['edit', 'undo'] as const)(
  'reconciles text %s if Stop races committed native write',
  async (action) => {
    const f = undoSetup()
    if (action === 'undo') {
      await f.skill.executeTool(f.edit)
      await f.proposals.confirm(f.proposals.pending()!.id)
    }
    const write = f.adapter.editPresentationPageText.getMockImplementation()!
    f.adapter.editPresentationPageText.mockImplementation(async (...args) => {
      await write(...args)
      f.proposals.newTurn()
    })
    await f.skill.executeTool(action === 'edit' ? f.edit : f.undo)
    await f.proposals.confirm(f.proposals.pending()!.id)
    expect(f.readTextChange()?.state).toBe(action === 'edit' ? 'applied' : 'undone')
  },
)
it('retains legacy text edits without storage and hides journal tools', async () => {
  const f = setup()
  expect(f.skill.tools.some((tool) => tool.name.endsWith('_text_change'))).toBe(false)
  await f.skill.executeTool(f.edit)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.adapter.editPresentationPageText).toHaveBeenCalledOnce()
})
it('supports full-length text undo with bounded, explicitly truncated proposal previews', async () => {
  const f = undoSetup()
  f.setText('原'.repeat(12000))
  await f.skill.executeTool({ ...f.edit, input: { ...f.edit.input, text: '新'.repeat(12000) } })
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(
    (await f.skill.executeTool({ ...f.undo, name: 'read_presentation_text_change' })).isError,
  ).not.toBe(true)
  expect((await f.skill.executeTool(f.undo)).isError).not.toBe(true)
  expect(f.proposals.pending()?.preview).toMatchObject({
    beforeTruncated: true,
    afterTruncated: true,
  })
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.readTextChange()?.state).toBe('undone')
  expect((await f.adapter.readPresentationPageText()).text).toBe('原'.repeat(12000))
})

function historySetup(kind: 'text' | 'geometry') {
  const f = setup()
  type Record =
    | import('../src/skills/powerpoint/presentation-text-change.js').PresentationTextChange
    | import('../src/skills/powerpoint/presentation-geometry-change.js').PresentationGeometryChange
  const records = new Map<string, Record>()
  let head: Record | undefined
  let geometry = { left: 0, top: 0, width: 100, height: 50 }
  const adapter = {
    ...f.adapter,
    readPresentationPageGeometry: vi.fn(async () => ({
      slideId: 'host-42',
      shapeId: 'shape1',
      geometry: { ...geometry },
    })),
    editPresentationPageGeometry: vi.fn(
      async (_slide: string, _shape: string, next: typeof geometry, expected: typeof geometry) => {
        expect(geometry).toEqual(expected)
        geometry = { ...next }
      },
    ),
  }
  const read = (id?: string) => (id === undefined ? head : records.get(id))
  const write = async (next: Record, expected: Record | undefined) => {
    expect(records.get(next.changeId) ?? head).toEqual(expected)
    records.set(next.changeId, structuredClone(next))
    head = records.get(next.changeId)
  }
  const options = {
    ...f.options,
    adapter,
    readTextChange: read as (
      id?: string,
    ) =>
      | import('../src/skills/powerpoint/presentation-text-change.js').PresentationTextChange
      | undefined,
    writeTextChange: write,
    readGeometryChange: read as (
      id?: string,
    ) =>
      | import('../src/skills/powerpoint/presentation-geometry-change.js').PresentationGeometryChange
      | undefined,
    writeGeometryChange: write,
  }
  const skill = createPresentationPageEditingSkill(options)
  const action = (verb: string, id?: string) =>
    skill.executeTool({
      id: verb,
      name: `${verb}_presentation_${kind}_change`,
      input: { page_id: 'page1', ...(id === undefined ? {} : { change_id: id }) },
    })
  const edit = async (step: number) => {
    expect(
      (
        await skill.executeTool({
          id: 'edit',
          name: `edit_presentation_page_${kind}`,
          input: {
            page_id: 'page1',
            shape_id: 'shape1',
            ...(kind === 'text'
              ? { text: `After${step}` }
              : { geometry: { ...geometry, left: step * 10 } }),
          },
        })
      ).isError,
    ).not.toBe(true)
    await f.proposals.confirm(f.proposals.pending()!.id)
    return head!
  }
  return { ...f, options, skill, action, edit, records, read, adapter }
}
it.each(['text', 'geometry'] as const)(
  'undoes two %s changes in reverse by exact ID',
  async (kind) => {
    const f = historySetup(kind)
    const first = await f.edit(1),
      second = await f.edit(2)
    for (const record of [second, first]) {
      expect((await f.action('read', record.changeId)).output).toContain(record.changeId)
      expect((await f.action('undo', record.changeId)).isError).not.toBe(true)
      await f.proposals.confirm(f.proposals.pending()!.id)
      expect(f.read(record.changeId)?.state).toBe('undone')
    }
    expect(f.records.size).toBe(2)
  },
)
it.each(['text', 'geometry'] as const)(
  'recovers an older pending %s record without selecting the latest slot',
  async (kind) => {
    const f = historySetup(kind)
    const first = await f.edit(1),
      second = await f.edit(2)
    f.records.set(first.changeId, { ...first, state: 'pending' })
    if (kind === 'text') f.setText(first.after as string)
    else
      f.adapter.readPresentationPageGeometry.mockResolvedValue({
        slideId: 'host-42',
        shapeId: 'shape1',
        geometry: first.after as { left: number; top: number; width: number; height: number },
      })
    const writes =
      kind === 'text' ? f.adapter.editPresentationPageText : f.adapter.editPresentationPageGeometry
    const count = writes.mock.calls.length
    expect((await f.action('inspect', first.changeId)).output).toContain('already_applied')
    expect((await f.action('resume', first.changeId)).isError).not.toBe(true)
    await f.proposals.confirm(f.proposals.pending()!.id)
    expect(writes).toHaveBeenCalledTimes(count)
    expect(f.read(first.changeId)?.state).toBe('applied')
    expect(f.read(second.changeId)).toEqual(second)
  },
)
it.each(['text', 'geometry'] as const)(
  'rejects absent IDs and readers ignoring selected %s identity',
  async (kind) => {
    const f = historySetup(kind)
    const first = await f.edit(1)
    await f.edit(2)
    expect((await f.action('undo', 'missing')).isError).toBe(true)
    f.options.readTextChange = () =>
      f.read() as import('../src/skills/powerpoint/presentation-text-change.js').PresentationTextChange
    f.options.readGeometryChange = () =>
      f.read() as import('../src/skills/powerpoint/presentation-geometry-change.js').PresentationGeometryChange
    for (const verb of ['read', 'inspect', 'undo', 'resume'])
      expect((await f.action(verb, first.changeId)).isError).toBe(true)
    expect(f.proposals.pending()).toBeUndefined()
  },
)

it.each(['text', 'geometry'] as const)(
  'validates %s history input and preserves saved scope checks',
  async (kind) => {
    const f = historySetup(kind),
      record = await f.edit(1)
    for (const id of ['', 'a'.repeat(129), 'bad/id', 42, null]) {
      expect(
        await f.skill.executeTool({
          id: 'bad',
          name: `read_presentation_${kind}_change`,
          input: { page_id: 'page1', change_id: id },
        }),
      ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
    }
    for (const patch of [
      { pageId: 'page2' },
      { requestId: 'other' },
      { projectId: 'other' },
      { documentId: 'other' },
      { hostSlideId: 'other' },
      { source: 'production' as const },
      { artifactDigest: '0'.repeat(64) },
    ]) {
      f.records.set(record.changeId, { ...record, ...patch })
      expect((await f.action('read', record.changeId)).isError).toBe(true)
    }
    const longest = { ...record, changeId: 'a'.repeat(128) }
    f.records.set(longest.changeId, longest)
    expect((await f.action('read', longest.changeId)).isError).not.toBe(true)
    expect(
      await f.skill.executeTool({
        ...f.edit,
        id: 'bad-edit',
        name: `edit_presentation_page_${kind}`,
        input: {
          page_id: 'page1',
          shape_id: 'shape1',
          change_id: record.changeId,
          ...(kind === 'text'
            ? { text: 'New' }
            : { geometry: { left: 0, top: 0, width: 100, height: 50 } }),
        },
      }),
    ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  },
)
it.each(['text', 'geometry'] as const)(
  'keeps %s confirmation pinned to the selected ID when the reader changes',
  async (kind) => {
    const f = historySetup(kind),
      first = await f.edit(1)
    expect((await f.action('undo', first.changeId)).isError).not.toBe(true)
    const changed = { ...first, changeId: 'other' }
    f.options.readTextChange = () =>
      changed as import('../src/skills/powerpoint/presentation-text-change.js').PresentationTextChange
    f.options.readGeometryChange = () =>
      changed as import('../src/skills/powerpoint/presentation-geometry-change.js').PresentationGeometryChange
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(f.records.get(first.changeId)?.state).toBe('applied')
  },
)
