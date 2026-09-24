import type { PresentationHistoryEnvelope } from '../src/skills/powerpoint/presentation-change-history'
import type { PresentationPageReplacement } from '../src/skills/powerpoint/presentation-page-replacement-record'
import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import type { PresentationTextChange } from '../src/skills/powerpoint/presentation-text-change'
const historyKey = 'wiswork.presentation.change-history.v1'
const textKey = 'wiswork.presentation.text-change.v1'
async function fixture() {
  const values = new Map<string, string>()
  const save = vi.fn(async () => {})
  let location = 'deck'
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = create()
  const record: PresentationTextChange = {
    version: 1,
    changeId: 'one',
    documentId: await binding.documentId(),
    projectId: 'project',
    requestId: 'request',
    artifactDigest: 'a'.repeat(64),
    pageId: 'page',
    hostSlideId: 'slide',
    shapeId: 'shape',
    before: 'one',
    after: 'two',
    state: 'pending',
  }
  return {
    values,
    save,
    settings,
    create,
    binding,
    record,
    move: () => {
      location = 'other'
    },
  }
}
it('retains same-kind records after reopening and selects older records exactly for CAS undo', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  const applied = { ...f.record, state: 'applied' as const }
  await f.binding.writeTextChange(applied, f.record)
  const second = { ...f.record, changeId: 'two', before: 'two', after: 'three' }
  await f.binding.writeTextChange(second, applied)
  const secondApplied = { ...second, state: 'applied' as const }
  await f.binding.writeTextChange(secondApplied, second)
  const reopened = f.create()
  expect(reopened.listChangeHistory()).toHaveLength(2)
  expect(reopened.readTextChange('one')).toEqual(applied)
  expect(reopened.readTextChange('missing')).toBeUndefined()
  await reopened.writeTextChange({ ...applied, state: 'undo_pending' }, applied)
  expect(reopened.readTextChange()?.changeId).toBe('one')
  expect(reopened.readTextChange('two')).toEqual(secondApplied)
  await expect(
    reopened.writeTextChange({ ...applied, state: 'undo_pending' }, applied),
  ).rejects.toThrow('stale')
})
it('saves legacy slot and history together and restores both after completion failure', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  const priorHistory = f.values.get(historyKey)
  const priorText = f.values.get(textKey)
  f.save.mockImplementationOnce(async () => {
    expect(JSON.parse(f.values.get(historyKey)!).entries[0].record.state).toBe('applied')
    expect(JSON.parse(f.values.get(textKey)!).state).toBe('applied')
    throw new Error('failed')
  })
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('failed')
  expect(f.values.get(historyKey)).toBe(priorHistory)
  expect(f.values.get(textKey)).toBe(priorText)
  expect(f.create().listChangeHistory()[0].record.state).toBe('pending')
})
it('imports old slots with unknown chronology and keeps cross-kind sequence stable', async () => {
  const f = await fixture()
  const legacy = { ...f.record, state: 'applied' as const }
  f.values.set(textKey, JSON.stringify(legacy))
  expect(f.binding.listChangeHistory()[0]).toMatchObject({
    kind: 'text',
    legacy: true,
    sequence: 1,
  })
  const geometry = {
    ...f.record,
    changeId: 'geometry',
    before: { left: 0, top: 0, width: 1, height: 1 },
    after: { left: 1, top: 0, width: 1, height: 1 },
  }
  await f.binding.writeGeometryChange(geometry, undefined)
  await f.binding.writeGeometryChange({ ...geometry, state: 'applied' }, geometry)
  expect(
    f
      .create()
      .listChangeHistory()
      .map((e) => [e.kind, e.sequence, e.legacy]),
  ).toEqual([
    ['text', 1, true],
    ['geometry', 2, false],
  ])
  const copied = f.binding.listChangeHistory()
  copied[0].record.state = 'undone'
  expect(f.binding.readTextChange()?.state).toBe('applied')
})
it('refuses another transaction while a pending record remains, then allows recovery', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  const geometry = {
    ...f.record,
    changeId: 'geometry',
    before: { left: 0, top: 0, width: 1, height: 1 },
    after: { left: 1, top: 0, width: 1, height: 1 },
  }
  await expect(f.binding.writeGeometryChange(geometry, undefined)).rejects.toThrow(
    'presentation_change_history_pending',
  )
  await f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record)
  await f.binding.writeGeometryChange(geometry, undefined)
})
it('rejects old-version head changes, malformed history, forged heads and duplicate ordering', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  const good = f.values.get(historyKey)!
  f.values.set(textKey, JSON.stringify({ ...f.record, state: 'applied' }))
  expect(() => f.create().readTextChange('one')).toThrow(
    'presentation_change_history_state_invalid',
  )
  f.values.set(textKey, JSON.stringify(f.record))
  for (const mutate of [
    (h: PresentationHistoryEnvelope & { extra?: boolean }) => {
      h.heads.text = 'text:missing'
    },
    (h: PresentationHistoryEnvelope & { extra?: boolean }) => {
      h.entries.push(h.entries[0])
    },
    (h: PresentationHistoryEnvelope & { extra?: boolean }) => {
      h.entries[0].sequence = 0
    },
    (h: PresentationHistoryEnvelope & { extra?: boolean }) => {
      h.entries[0].id = 'text:wrong'
    },
    (h: PresentationHistoryEnvelope & { extra?: boolean }) => {
      h.extra = true
    },
  ]) {
    const h = JSON.parse(good)
    mutate(h)
    f.values.set(historyKey, JSON.stringify(h))
    expect(() => f.create().listChangeHistory()).toThrow(
      'presentation_change_history_state_invalid',
    )
  }
  f.values.set(historyKey, '{')
  expect(() => f.create().listChangeHistory()).toThrow('presentation_change_history_state_invalid')
})
it('does not overwrite an earlier identity by reusing its ID as a new operation', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  const applied = { ...f.record, state: 'applied' as const }
  await f.binding.writeTextChange(applied, f.record)
  const second = { ...f.record, changeId: 'two' }
  await f.binding.writeTextChange(second, applied)
  const done = { ...second, state: 'applied' as const }
  await f.binding.writeTextChange(done, second)
  await expect(f.binding.writeTextChange({ ...f.record, after: 'forged' }, done)).rejects.toThrow(
    'stale',
  )
  expect(f.binding.readTextChange('one')).toEqual(applied)
})
it('blocks capacity before saving and never prunes completed records', async () => {
  const f = await fixture()
  let previous: PresentationTextChange | undefined
  for (let i = 0; i < 64; i++) {
    const pending = { ...f.record, changeId: `entry-${i}` }
    await f.binding.writeTextChange(pending, previous)
    previous = { ...pending, state: 'applied' }
    await f.binding.writeTextChange(previous, pending)
  }
  const saves = f.save.mock.calls.length
  await expect(
    f.binding.writeTextChange({ ...f.record, changeId: 'overflow' }, previous),
  ).rejects.toThrow('presentation_change_history_full')
  expect(f.save.mock.calls.length).toBe(saves)
  expect(f.create().listChangeHistory()).toHaveLength(64)
  await f.binding.writeTextChange({ ...previous!, state: 'undo_pending' }, previous)
})
it('reserves byte capacity before native writes and allows finalization at the boundary', async () => {
  const f = await fixture()
  let previous: PresentationTextChange | undefined
  let count = 0
  for (; count < 64; count++) {
    const pending = {
      ...f.record,
      changeId: `large-${count}`,
      before: '\u0000'.repeat(12000),
      after: '\u0001'.repeat(12000),
    }
    try {
      await f.binding.writeTextChange(pending, previous)
    } catch (error) {
      expect(String(error)).toContain('presentation_change_history_full')
      break
    }
    previous = { ...pending, state: 'applied' }
    await f.binding.writeTextChange(previous, pending)
  }
  expect(count).toBeGreaterThan(1)
  expect(count).toBeLessThan(64)
  expect(f.binding.listChangeHistory()).toHaveLength(count)
  await f.binding.writeTextChange({ ...previous!, state: 'undo_pending' }, previous)
})
it('restores both keys on set failure and locks all history if rollback cannot be proven', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  const before = new Map(f.values)
  const set = f.settings.set
  f.settings.set = (key, value) => {
    if (key === historyKey) throw new Error('set_failed')
    set(key, value)
  }
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('set_failed')
  expect(f.values.get(textKey)).toBe(before.get(textKey))
  expect(() => f.binding.listChangeHistory()).toThrow('presentation_change_history_state_invalid')
})
it('locks history after a Save As during save', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  f.save.mockImplementationOnce(async () => f.move())
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('presentation_document_changed')
  expect(() => f.binding.listChangeHistory()).toThrow('presentation_change_history_state_invalid')
})
it('preserves image history alongside text and rejects a legacy image-map change', async () => {
  const f = await fixture()
  const { imageReplacementKey } =
    await import('../src/skills/powerpoint/presentation-image-replacement-record')
  const key = await imageReplacementKey('project', 'request', 'page', 'old')
  const image = {
    version: 1 as const,
    documentId: f.record.documentId,
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    hostSlideId: 'slide',
    oldShapeId: 'old',
    assetDigest: 'a'.repeat(64),
    state: 'pending' as const,
  }
  await f.binding.writeImageReplacement(key, image)
  const candidate = { ...image, newShapeId: 'new' }
  await f.binding.writeImageReplacement(key, candidate)
  await f.binding.writeImageReplacement(key, { ...candidate, state: 'complete' })
  await f.binding.writeTextChange(f.record, undefined)
  expect(f.binding.listChangeHistory().map((e) => [e.kind, e.sequence])).toEqual([
    ['image', 1],
    ['text', 2],
  ])
  f.values.set('wiswork.presentation.image-replacements.v1', '{}')
  expect(() => f.create().listChangeHistory()).toThrow('presentation_change_history_state_invalid')
})
it('reserves whole-page growth to the record limit before starting another page transaction', async () => {
  const f = await fixture()
  const page = {
    version: 1 as const,
    changeId: 'page-change',
    documentId: f.record.documentId,
    projectId: 'project',
    parentRequestId: 'parent',
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
    state: 'pending' as const,
  }
  let prior: PresentationPageReplacement | undefined
  for (let i = 0; i < 5; i++) {
    const pending = { ...page, changeId: `page-${i}` }
    await f.binding.writePageReplacement(pending, prior)
    const inserted = { ...pending, newSlideId: `new-${i}`, state: 'inserted' as const }
    await f.binding.writePageReplacement(inserted, pending)
    const staged = { ...inserted, state: 'staged' as const }
    await f.binding.writePageReplacement(staged, inserted)
    const discarding = { ...staged, state: 'discard_pending' as const }
    await f.binding.writePageReplacement(discarding, staged)
    const discarded = { ...discarding, state: 'discarded' as const }
    await f.binding.writePageReplacement(discarded, discarding)
    prior = discarded
  }
  const saves = f.save.mock.calls.length
  await expect(
    f.binding.writePageReplacement({ ...page, changeId: 'overflow' }, prior),
  ).rejects.toThrow('presentation_change_history_full')
  expect(f.save.mock.calls.length).toBe(saves)
  expect(f.create().listChangeHistory()).toHaveLength(5)
})
it('blocks public image recovery readers when a downlevel image writer diverged', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  f.values.set(
    'wiswork.presentation.image-replacements.v1',
    JSON.stringify({
      ['a'.repeat(64)]: {
        version: 1,
        documentId: f.record.documentId,
        projectId: 'project',
        requestId: 'request',
        pageId: 'page',
        hostSlideId: 'slide',
        oldShapeId: 'old',
        assetDigest: 'a'.repeat(64),
        state: 'pending',
      },
    }),
  )
  expect(() => f.binding.readImageReplacement('a'.repeat(64))).toThrow(
    'presentation_change_history_state_invalid',
  )
  expect(() => f.binding.listImageReplacements()).toThrow(
    'presentation_change_history_state_invalid',
  )
})
it.each(['text', 'geometry'] as const)(
  'can undo %s at the exact accepted history byte boundary',
  async (kind) => {
    const f = await fixture()
    const { presentationHistoryBytes } =
      await import('../src/skills/powerpoint/presentation-change-history')
    const geometry = {
      ...f.record,
      before: { left: 0, top: 0, width: 1, height: 1 },
      after: { left: 1, top: 0, width: 1, height: 1 },
    }
    if (kind === 'text') {
      await f.binding.writeTextChange(f.record, undefined)
      await f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record)
    } else {
      await f.binding.writeGeometryChange(geometry, undefined)
      await f.binding.writeGeometryChange({ ...geometry, state: 'applied' }, geometry)
    }
    let previous = f.binding.readTextChange()
    for (let i = 0; i < 7; i++) {
      const pending = {
        ...f.record,
        changeId: `filler-${i}`,
        before: '\u0000'.repeat(12000),
        after: '\u0001'.repeat(12000),
      }
      await f.binding.writeTextChange(pending, previous)
      previous = { ...pending, state: 'applied' }
      await f.binding.writeTextChange(previous, pending)
    }
    const pending = { ...f.record, changeId: 'boundary', before: '', after: '' }
    const h: PresentationHistoryEnvelope = JSON.parse(f.values.get(historyKey)!)
    h.entries.push({
      id: 'text:boundary',
      sequence: h.entries.length + 1,
      legacy: false,
      kind: 'text',
      record: pending,
    })
    h.heads.text = 'text:boundary'
    const padding = 1024 * 1024 - presentationHistoryBytes(h)
    expect(padding).toBeGreaterThan(0)
    pending.after = '\u0000'.repeat(Math.floor(padding / 6)) + 'x'.repeat(padding % 6)
    expect(pending.after.length).toBeLessThanOrEqual(12000)
    await f.binding.writeTextChange(pending, previous)
    await f.binding.writeTextChange({ ...pending, state: 'applied' }, pending)
    expect(presentationHistoryBytes(JSON.parse(f.values.get(historyKey)!))).toBe(1024 * 1024)
    if (kind === 'text')
      await f.binding.writeTextChange(
        { ...f.record, state: 'undo_pending' },
        { ...f.record, state: 'applied' },
      )
    else
      await f.binding.writeGeometryChange(
        { ...geometry, state: 'undo_pending' },
        { ...geometry, state: 'applied' },
      )
  },
)
