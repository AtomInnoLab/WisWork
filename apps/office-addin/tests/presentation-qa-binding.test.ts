import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { type PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa'
import { PresentationQaCard } from '../src/agent/presentation-qa-card'

function fixture() {
  const values = new Map<string, string>()
  let location = 'file://deck.pptx',
    fail = false
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {
      if (fail) throw new Error('save_failed')
    },
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'document-qa')
  return {
    binding: create(),
    create,
    values,
    fail: () => {
      fail = true
    },
    move: () => {
      location = 'file://other.pptx'
    },
  }
}
function record(documentId: string, requestId = 'request-1'): PresentationQaRecord {
  return {
    version: 1,
    documentId,
    projectId: 'project',
    requestId,
    artifactDigest: 'a'.repeat(64),
    pages: [
      {
        pageId: 'page1',
        title: '第一页',
        hostSlideId: '256',
        capturedAt: '2026-09-23T08:00:00.000Z',
        screenshotDigest: 'b'.repeat(64),
        screenshotBytes: 128,
        structure: {
          status: 'passed',
          shapeCount: 1,
          overflowCount: 0,
          overlapCount: 0,
          shapesTruncated: false,
          overlapsTruncated: false,
        },
        visual: { status: 'needs_review' },
      },
    ],
  }
}
it('persists bounded QA metadata and reads it through a fresh binding', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  expect(f.create().readQa('project/request-1')).toEqual(value)
  expect([...f.values.values()].join('')).not.toContain('base64')
})
it('restores the prior review if settings save fails and rejects document mismatch', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  f.fail()
  const next = structuredClone(value)
  next.pages[0]!.visual = {
    status: 'pass',
    reviewer: 'agent',
    notes: '布局清晰',
    reviewedAt: '2026-09-23T08:01:00.000Z',
  }
  await expect(f.binding.writeQa('project/request-1', next)).rejects.toThrow('save_failed')
  expect(f.binding.readQa('project/request-1')).toEqual(value)
  f.move()
  await expect(f.binding.writeQa('project/request-1', value)).rejects.toThrow(
    'presentation_document_changed',
  )
})
it('rejects malformed persisted QA and limits retained request history', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  for (let i = 0; i < 8; i++)
    await f.binding.writeQa(`project/request-${i}`, record(doc, `request-${i}`))
  await expect(f.binding.writeQa('project/request-8', record(doc, 'request-8'))).rejects.toThrow(
    'presentation_qa_history_full',
  )
  f.values.set(
    'wiswork.presentation.qa.v1',
    JSON.stringify({ 'project/request-1': { ...record(doc), extra: 'untrusted' } }),
  )
  expect(() => f.binding.readQa('project/request-1')).toThrow('presentation_qa_state_invalid')
})
it('labels a restored Agent assessment as historical and escapes reviewer text', () => {
  const value = record('doc')
  value.pages[0]!.visual = {
    status: 'pass',
    reviewer: 'agent',
    notes: '<script>bad()</script>',
    reviewedAt: '2026-09-23T08:01:00.000Z',
  }
  const html = renderToStaticMarkup(
    createElement(PresentationQaCard, {
      controller: { read: () => value, revision: () => 0, subscribe: () => () => undefined },
    }),
  )
  expect(html).toContain('历史检查记录')
  expect(html).toContain('Agent 判断通过')
  expect(html).not.toContain('<script>')
})

it('marks every saved request for recheck before mutation and preserves historical judgments', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  const first = record(doc)
  first.pages[0]!.visual = {
    status: 'pass',
    reviewer: 'agent',
    notes: '旧截图通过',
    reviewedAt: '2026-09-23T08:01:00.000Z',
  }
  await f.binding.writeQa('project/request-1', first)
  await f.binding.writeQa('project/request-2', record(doc, 'request-2'))
  await f.binding.invalidateQa()
  for (const request of ['request-1', 'request-2']) {
    expect(f.create().readQa(`project/${request}`)?.pages[0]?.recheckRequired).toBe(true)
  }
  expect(f.binding.readQa('project/request-1')?.pages[0]?.visual).toEqual(first.pages[0]!.visual)
})
it('blocks invalidation on save failure and restores old metadata', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  f.fail()
  await expect(f.binding.invalidateQa()).rejects.toThrow('save_failed')
  expect(f.binding.readQa('project/request-1')).toEqual(value)
})
it('does not write empty QA and reads the latest queued capture before invalidating', async () => {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => values.get(key),
      set: (key, value) => {
        values.set(key, value)
      },
      save,
      location: () => 'test',
    },
    () => 'doc',
  )
  await binding.invalidateQa()
  expect(save).not.toHaveBeenCalled()
  const value = record(await binding.documentId())
  const pending = binding.writeQa('project/request-1', value)
  const invalidation = binding.invalidateQa()
  await Promise.all([pending, invalidation])
  expect(binding.readQa('project/request-1')?.pages[0]?.recheckRequired).toBe(true)
})
function nearLimitRecord(
  documentId: string,
  requestId: string,
  targetBytes: number,
): PresentationQaRecord {
  const value = record(documentId, requestId),
    page = value.pages[0]!
  value.pages = Array.from({ length: 32 }, (_, i) => ({
    ...structuredClone(page),
    pageId: `page${i}`,
    hostSlideId: `host${i}`,
    visual: {
      status: 'pass',
      reviewer: 'agent',
      notes: 'x',
      reviewedAt: '2026-09-23T08:01:00.000Z',
    },
  }))
  let remaining = targetBytes - new TextEncoder().encode(JSON.stringify(value)).byteLength
  for (const page of value.pages) {
    const added = Math.min(1999, remaining)
    page.visual.notes += 'x'.repeat(added)
    remaining -= added
  }
  expect(remaining).toBe(0)
  expect(new TextEncoder().encode(JSON.stringify(value)).byteLength).toBe(targetBytes)
  return value
}
it('invalidates a near-64KiB captured record without consuming its content budget', async () => {
  const f = fixture(),
    doc = await f.binding.documentId(),
    value = nearLimitRecord(doc, 'request-1', 64 * 1024 - 10)
  await f.binding.writeQa('project/request-1', value)
  await f.binding.invalidateQa()
  const restored = f.create().readQa('project/request-1')!
  expect(new TextEncoder().encode(JSON.stringify(restored)).byteLength).toBeGreaterThan(64 * 1024)
  expect(restored.pages.every((page) => page.recheckRequired === true)).toBe(true)
  expect(restored.pages.map((page) => page.visual)).toEqual(value.pages.map((page) => page.visual))
  // A fresh capture removes only its own bookkeeping flag; another invalidation remains safe.
  const recaptured = {
    ...restored,
    pages: restored.pages.map((page, i) => {
      if (i) return page
      const { recheckRequired: _flag, ...rest } = page
      return rest
    }),
  }
  await f.binding.writeQa('project/request-1', recaptured)
  await f.binding.invalidateQa()
  expect(f.create().readQa('project/request-1')).toEqual(restored)
})
it('invalidates near-256KiB history and retains the original aggregate content ceiling', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  for (let i = 0; i < 4; i++)
    await f.binding.writeQa(`project/request-${i}`, nearLimitRecord(doc, `request-${i}`, 65500))
  const key = 'wiswork.presentation.qa.v1',
    beforeBytes = new TextEncoder().encode(f.values.get(key)!).byteLength
  expect(beforeBytes).toBeGreaterThan(256 * 1024 - 256)
  expect(beforeBytes).toBeLessThanOrEqual(256 * 1024)
  await f.binding.invalidateQa()
  expect(new TextEncoder().encode(f.values.get(key)!).byteLength).toBeGreaterThan(256 * 1024)
  for (let i = 0; i < 4; i++)
    expect(
      f
        .create()
        .readQa(`project/request-${i}`)
        ?.pages.every((page) => page.recheckRequired),
    ).toBe(true)
  const over = nearLimitRecord(doc, 'request-0', 65537)
  over.pages = over.pages.map((page) => ({ ...page, recheckRequired: true }))
  await expect(f.binding.writeQa('project/request-0', over)).rejects.toThrow(
    'presentation_qa_state_invalid',
  )
  const increased = nearLimitRecord(doc, 'request-0', 65536)
  increased.pages = increased.pages.map((page) => ({ ...page, recheckRequired: true }))
  await f.binding.writeQa('project/request-0', increased)
  const last = nearLimitRecord(doc, 'request-1', 65536)
  last.pages = last.pages.map((page) => ({ ...page, recheckRequired: true }))
  await expect(f.binding.writeQa('project/request-1', last)).rejects.toThrow(
    'presentation_qa_history_full',
  )
  const corrupted = JSON.parse(f.values.get(key)!)
  corrupted['project/request-1'] = last
  f.values.set(key, JSON.stringify(corrupted))
  expect(() => f.create().readQa('project/request-1')).toThrow('presentation_qa_state_invalid')
})

it('invalidates matching host pages across requests while retaining unrelated QA exactly', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  for (const request of ['request-1', 'request-2']) {
    const value = record(doc, request)
    value.pages.push({ ...structuredClone(value.pages[0]!), pageId: 'other', hostSlideId: '257' })
    await f.binding.writeQa(`project/${request}`, value)
  }
  const unrelated = structuredClone(f.binding.readQa('project/request-1')!.pages[1])
  await f.binding.invalidateQa(['256'])
  for (const request of ['request-1', 'request-2']) {
    const saved = f.create().readQa(`project/${request}`)!
    expect(saved.pages[0]!.recheckRequired).toBe(true)
    expect(saved.pages[1]).toEqual(unrelated)
  }
})
it('does not save unmatched scopes and snapshots scope before joining the settings queue', async () => {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => values.get(key),
      set: (key, value) => {
        values.set(key, value)
      },
      save,
      location: () => 'test',
    },
    () => 'doc',
  )
  const value = record(await binding.documentId())
  value.pages.push({ ...structuredClone(value.pages[0]!), pageId: 'other', hostSlideId: '257' })
  await binding.writeQa('project/request-1', value)
  save.mockClear()
  await binding.invalidateQa(['unknown'])
  expect(save).not.toHaveBeenCalled()
  let release!: () => void
  const gate = new Promise<void>((resolve) => {
    release = resolve
  })
  save.mockImplementationOnce(async () => {
    await gate
  })
  const writing = binding.writeQa('project/request-1', value)
  const scope = ['256']
  const pending = binding.invalidateQa(scope)
  scope[0] = '257'
  release()
  await Promise.all([writing, pending])
  expect(binding.readQa('project/request-1')!.pages.map((page) => page.recheckRequired)).toEqual([
    true,
    undefined,
  ])
})
it('rejects invalid explicit scopes without turning them into full-deck invalidation', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  for (const scope of [
    ['256', '256'],
    [''],
    ['x'.repeat(257)],
    ['bad\n'],
    ['bad\x7f'],
    Array.from({ length: 101 }, (_, i) => String(i)),
    [null],
    null,
  ]) {
    await expect(
      Promise.resolve().then(() => f.binding.invalidateQa(scope as string[])),
    ).rejects.toThrow('invalid_tool_input')
    expect(f.binding.readQa('project/request-1')).toEqual(value)
  }
})
it('rolls back a failed scoped save without invalidating unrelated pages', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  value.pages.push({ ...structuredClone(value.pages[0]!), pageId: 'other', hostSlideId: '257' })
  await f.binding.writeQa('project/request-1', value)
  f.fail()
  await expect(f.binding.invalidateQa(['256'])).rejects.toThrow('save_failed')
  expect(f.binding.readQa('project/request-1')).toEqual(value)
})
it('accepts the maximum unique scope and ID length without broadening unmatched invalidation', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  await f.binding.invalidateQa(Array.from({ length: 100 }, (_, i) => String(i).padStart(256, 'x')))
  expect(f.binding.readQa('project/request-1')).toEqual(value)
})

it('isolates production QA records and invalidates shared host pages across namespaces', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  const production = { ...structuredClone(value), source: 'production' as const }
  await f.binding.writeQa('project/request-1', value)
  await f.binding.writeQa('production/project/request-1', production)
  expect(f.create().readQa('production/project/request-1')).toEqual(production)
  await expect(f.binding.writeQa('project/request-1', production)).rejects.toThrow(
    'presentation_qa_state_invalid',
  )
  await expect(f.binding.writeQa('production/project/request-1', value)).rejects.toThrow(
    'presentation_qa_state_invalid',
  )
  await f.binding.invalidateQa(['256'])
  for (const key of ['project/request-1', 'production/project/request-1']) {
    expect(f.create().readQa(key)?.pages[0]?.recheckRequired).toBe(true)
  }
  f.values.set('wiswork.presentation.qa.v1', JSON.stringify({ 'project/request-1': production }))
  expect(() => f.create().readQa('project/request-1')).toThrow('presentation_qa_state_invalid')
})
it('retains legacy QA project named production and rejects unsupported source values', async () => {
  const f = fixture(),
    value = { ...record(await f.binding.documentId()), projectId: 'production' }
  await f.binding.writeQa('production/request-1', value)
  expect(f.create().readQa('production/request-1')).toEqual(value)
  await expect(
    f.binding.writeQa('production/request-1', {
      ...value,
      source: 'legacy',
    } as unknown as PresentationQaRecord),
  ).rejects.toThrow('presentation_qa_state_invalid')
})
it('persists image source namespaces and refuses adding source to a prior image record', async () => {
  const { imageReplacementKey } =
    await import('../src/skills/powerpoint/presentation-image-replacement-record')
  const f = fixture(),
    documentId = await f.binding.documentId()
  const value = {
    version: 1 as const,
    documentId,
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    hostSlideId: 'host',
    oldShapeId: 'old',
    assetDigest: 'a'.repeat(64),
    state: 'pending' as const,
  }
  const oldKey = await imageReplacementKey('project', 'request', 'page', 'old')
  const productionKey = await imageReplacementKey('project', 'request', 'page', 'old', 'production')
  const production = { ...value, source: 'production' as const }
  await f.binding.writeImageReplacement(oldKey, value)
  await expect(f.binding.writeImageReplacement(oldKey, production)).rejects.toThrow(
    'presentation_image_replacement_state_invalid',
  )
  await expect(f.binding.writeImageReplacement(productionKey, production)).rejects.toThrow(
    'presentation_change_history_pending',
  )
  const candidate = { ...value, newShapeId: 'new' }
  const complete = { ...candidate, state: 'complete' as const }
  await f.binding.writeImageReplacement(oldKey, candidate)
  await f.binding.writeImageReplacement(oldKey, complete)
  await f.binding.writeImageReplacement(productionKey, production)
  expect(f.create().readImageReplacement(productionKey)).toEqual(production)
  expect(f.create().readImageReplacement(oldKey)).toEqual(complete)
  // Even a manually mis-keyed record cannot change its original source while saving a candidate.
  f.values.set(
    'wiswork.presentation.image-replacements.v1',
    JSON.stringify({ [productionKey]: value }),
  )
  await expect(
    f.binding.writeImageReplacement(productionKey, { ...production, newShapeId: 'candidate' }),
  ).rejects.toThrow('presentation_image_replacement_state_invalid')
})

it('treats known empty dependencies as no invalidation and no settings save', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  f.fail()
  await f.binding.invalidateQa([])
  expect(f.create().readQa('project/request-1')).toEqual(value)
})
