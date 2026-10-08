import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
const attemptKey = 'wiswork.presentation.qa-attempts.v1'
function fixture(initialLocation = 'file://synthetic.pptx') {
  const values = new Map<string, string>()
  let location = initialLocation,
    fail = false
  const save = vi.fn(async () => {
    if (fail) throw Error('save_failed')
  })
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'attempt-doc')
  return {
    values,
    settings,
    save,
    create,
    binding: create(),
    fail: () => {
      fail = true
    },
    move: () => {
      location = 'file://other.pptx'
    },
  }
}
const attempt = (documentId: string, index = 0) => ({
  version: 1 as const,
  id: `12345678-1234-4234-8234-${index.toString(16).padStart(12, '0')}`,
  documentId,
  projectId: 'project',
  requestId: 'request',
  artifactDigest: 'a'.repeat(64),
  pageId: 'page',
  hostSlideId: 'host',
  startedAt: '2026-09-29T00:00:00.000Z',
  status: 'started' as const,
})
it('persists original start then immutable terminal across reopen and lost ACK retries', async () => {
  const f = fixture(),
    a = attempt(await f.binding.documentId())
  expect(f.binding.readQaAttempts('project/request')).toEqual([])
  await f.binding.writeQaAttempt('project/request', a)
  expect(f.create().readQaAttempts('project/request')).toEqual([a])
  const terminal = { ...a, status: 'recorded' as const, finishedAt: a.startedAt }
  await f.binding.writeQaAttempt('project/request', terminal)
  const saves = f.save.mock.calls.length
  await f.binding.writeQaAttempt('project/request', {
    finishedAt: a.startedAt,
    ...a,
    status: 'recorded',
  })
  expect(f.save).toHaveBeenCalledTimes(saves)
  expect(f.create().readQaAttempts('project/request')).toEqual([terminal])
  await expect(
    f.binding.writeQaAttempt('project/request', {
      ...terminal,
      status: 'failed',
      errorCode: 'state_changed',
    }),
  ).rejects.toThrow()
  await expect(f.binding.writeQaAttempt('project/request', a)).rejects.toThrow()
})
it('rejects missing terminals, identity mutations, wrong key/doc and corrupted journal', async () => {
  const f = fixture(),
    a = attempt(await f.binding.documentId())
  await expect(
    f.binding.writeQaAttempt('project/request', {
      ...a,
      status: 'cancelled',
      finishedAt: a.startedAt,
      errorCode: 'cancelled',
    }),
  ).rejects.toThrow()
  await expect(f.binding.writeQaAttempt('production/project/request', a)).rejects.toThrow()
  await expect(
    f.binding.writeQaAttempt('project/request', { ...a, documentId: 'wrong' }),
  ).rejects.toThrow('presentation_document_changed')
  await f.binding.writeQaAttempt('project/request', a)
  await expect(
    f.binding.writeQaAttempt('project/request', {
      ...a,
      pageId: 'wrong',
      status: 'recorded',
      finishedAt: a.startedAt,
    }),
  ).rejects.toThrow()
  f.values.set(attemptKey, '{bad')
  expect(() => f.create().readQaAttempts('project/request')).toThrow()
})
it('never evicts started attempts and prunes only oldest finished when appending at capacity', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  for (let i = 0; i < 64; i++) await f.binding.writeQaAttempt('project/request', attempt(doc, i))
  await expect(f.binding.writeQaAttempt('project/request', attempt(doc, 64))).rejects.toThrow(
    'presentation_qa_attempt_history_full',
  )
  const a = attempt(doc, 20)
  await f.binding.writeQaAttempt('project/request', {
    ...a,
    status: 'waiting',
    finishedAt: a.startedAt,
    errorCode: 'screenshot_unavailable',
  })
  await f.binding.writeQaAttempt('project/request', attempt(doc, 64))
  const list = f.create().readQaAttempts('project/request')
  expect(list).toHaveLength(64)
  expect(list.some((v) => v.id === a.id)).toBe(false)
  expect(list.some((v) => v.id === attempt(doc, 0).id)).toBe(true)
})
it('rolls back save failures and poisons unsafe cross-document saves; read results are detached', async () => {
  const f = fixture(),
    a = attempt(await f.binding.documentId())
  await f.binding.writeQaAttempt('project/request', a)
  const before = f.values.get(attemptKey)
  f.fail()
  await expect(
    f.binding.writeQaAttempt('project/request', {
      ...a,
      status: 'recorded',
      finishedAt: a.startedAt,
    }),
  ).rejects.toThrow('save_failed')
  expect(f.values.get(attemptKey)).toBe(before)
  expect(f.create().readQaAttempts('project/request')).toEqual([a])
  const list = f.binding.readQaAttempts('project/request')
  list[0]!.pageId = 'mutated'
  expect(f.binding.readQaAttempts('project/request')).toEqual([a])
  f.move()
  expect(f.binding.readQaAttempts('project/request')).toEqual([])
})
it('enforces the UTF8 envelope quota across scopes and prunes a finished attempt only when a new start needs room', async () => {
  const f = fixture('file://' + '界'.repeat(1000)),
    doc = await f.binding.documentId()
  let count = 0
  while (count < 64) {
    try {
      await f.binding.writeQaAttempt('project/request', attempt(doc, count))
      count++
    } catch (error) {
      expect((error as Error).message).toBe('presentation_qa_attempt_history_full')
      break
    }
  }
  expect(count).toBeLessThan(64)
  expect(count).toBeGreaterThan(0)
  expect(new TextEncoder().encode(f.values.get(attemptKey)!).byteLength).toBeLessThanOrEqual(
    128 * 1024,
  )
  const first = attempt(doc, 0)
  await f.binding.writeQaAttempt('project/request', {
    ...first,
    status: 'recorded',
    finishedAt: first.startedAt,
  })
  const other = {
    ...attempt(doc, count),
    source: 'production' as const,
    projectId: 'other',
    requestId: 'other',
  }
  await f.binding.writeQaAttempt('production/other/other', other)
  expect(f.create().readQaAttempts('production/other/other')).toEqual([other])
  expect(f.create().readQaAttempts('project/request')).toHaveLength(count - 1)
})
it('serializes concurrent transitions and poisons a save that changes the current document', async () => {
  const f = fixture(),
    a = attempt(await f.binding.documentId())
  let release!: () => void, entered!: () => void
  const gate = new Promise<void>((resolve) => {
      release = resolve
    }),
    ready = new Promise<void>((resolve) => {
      entered = resolve
    })
  f.save.mockImplementationOnce(async () => {
    entered()
    await gate
  })
  const start = f.binding.writeQaAttempt('project/request', a)
  await ready
  const done = f.binding.writeQaAttempt('project/request', {
    ...a,
    status: 'recorded',
    finishedAt: a.startedAt,
  })
  release()
  await Promise.all([start, done])
  expect(f.binding.readQaAttempts('project/request')[0]!.status).toBe('recorded')
  const b = attempt(a.documentId, 2)
  f.save.mockImplementationOnce(async () => {
    f.move()
  })
  await expect(f.binding.writeQaAttempt('project/request', b)).rejects.toThrow(
    'presentation_document_changed',
  )
  expect(() => f.binding.readQaAttempts('project/request')).toThrow(
    'presentation_qa_attempt_state_invalid',
  )
  expect(f.create().readQaAttempts('project/request')).toEqual([])
})
it.each(['waiting', 'closed'] as const)(
  'reserves enough terminal bytes for every accepted start at a full journal: %s',
  async (status) => {
    const f = fixture('file://' + '界'.repeat(1000)),
      doc = await f.binding.documentId()
    const reserve =
      2 +
      ',"finishedAt":"0000-00-00T00:00:00.000Z"'.length +
      ',"errorCode":"screenshot_unavailable"'.length
    let count = 0
    while (count < 64) {
      try {
        await f.binding.writeQaAttempt('project/request', attempt(doc, count))
        count++
      } catch (error) {
        expect((error as Error).message).toBe('presentation_qa_attempt_history_full')
        break
      }
    }
    // Fill any remaining byte budget by updating the last accepted started identity in a synthetic journal.
    const map = JSON.parse(f.values.get(attemptKey)!)
    let raw = JSON.stringify(map)
    for (const value of Object.values(map) as ReturnType<typeof attempt>[]) {
      const room = 128 * 1024 - new TextEncoder().encode(raw).byteLength
      if (room <= reserve * count) break
      value.hostSlideId = 'h'.repeat(Math.min(256, 4 + room - reserve * count))
      raw = JSON.stringify(map)
    }
    f.values.set(attemptKey, raw)
    expect(new TextEncoder().encode(raw).byteLength + reserve * count).toBeLessThanOrEqual(
      128 * 1024,
    )
    // The parser must ensure writes had already budgeted all pending terminal metadata.
    const invalid = JSON.parse(raw)
    for (const value of Object.values(invalid) as ReturnType<typeof attempt>[]) {
      const room = 128 * 1024 - 10 - new TextEncoder().encode(JSON.stringify(invalid)).byteLength
      if (room <= 0) break
      value.hostSlideId = 'h'.repeat(Math.min(256, value.hostSlideId.length + room))
    }
    expect(new TextEncoder().encode(JSON.stringify(invalid)).byteLength).toBeLessThanOrEqual(
      128 * 1024,
    )
    expect(
      new TextEncoder().encode(JSON.stringify(invalid)).byteLength + reserve * count,
    ).toBeGreaterThan(128 * 1024)
    f.values.set(attemptKey, JSON.stringify(invalid))
    expect(() => f.create().readQaAttempts('project/request')).toThrow(
      'presentation_qa_attempt_state_invalid',
    )
    f.values.set(attemptKey, raw)
    for (const a of f.create().readQaAttempts('project/request'))
      await f.binding.writeQaAttempt('project/request', {
        ...a,
        status,
        finishedAt: a.startedAt,
        errorCode: status === 'closed' ? 'explicitly_closed' : 'screenshot_unavailable',
      })
    expect(f.create().readQaAttempts('project/request')).toHaveLength(count)
    expect(
      f
        .create()
        .readQaAttempts('project/request')
        .every((a) => a.status === status),
    ).toBe(true)
  },
)

it('explicit close releases one full-journal slot, preserves other starts and rejects late completion', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  for (let i = 0; i < 64; i++) await f.binding.writeQaAttempt('project/request', attempt(doc, i))
  const a = attempt(doc, 20)
  const closed = {
    ...a,
    status: 'closed' as const,
    finishedAt: a.startedAt,
    errorCode: 'explicitly_closed' as const,
  }
  await f.binding.writeQaAttempt('project/request', closed)
  const saves = f.save.mock.calls.length
  await f.create().writeQaAttempt('project/request', closed)
  expect(f.save).toHaveBeenCalledTimes(saves)
  await expect(
    f.binding.writeQaAttempt('project/request', {
      ...a,
      status: 'recorded',
      finishedAt: a.startedAt,
    }),
  ).rejects.toThrow('presentation_qa_attempt_state_invalid')
  await f.binding.writeQaAttempt('project/request', attempt(doc, 64))
  expect(f.create().readQaAttempts('project/request')).toEqual(
    Array.from({ length: 65 }, (_, i) => attempt(doc, i)).filter((a) => a.id !== closed.id),
  )
})
it('rolls back a failed explicit close without losing the original start on reopen', async () => {
  const f = fixture(),
    a = attempt(await f.binding.documentId())
  await f.binding.writeQaAttempt('project/request', a)
  f.fail()
  await expect(
    f.binding.writeQaAttempt('project/request', {
      ...a,
      status: 'closed',
      finishedAt: a.startedAt,
      errorCode: 'explicitly_closed',
    }),
  ).rejects.toThrow('save_failed')
  expect(f.create().readQaAttempts('project/request')).toEqual([a])
})
