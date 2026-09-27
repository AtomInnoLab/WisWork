import { describe, expect, it } from 'vitest'
import { parsePresentationProductionJob } from '../src/presentation-job.js'

describe('production job contract', () => {
  const job = {
    version: 1,
    projectId: 'p',
    documentId: 'doc',
    requestId: 'r',
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    planRevision: 1,
    revision: 1,
    state: 'running',
    events: [{ sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' }],
  }
  it('replays two interleaved in-flight pages while rejecting a third', () => {
    const at = '2026-09-24T00:00:00.000Z'
    const events = [
      { type: 'run.started' },
      { type: 'page.started', pageId: 'a', attempt: 1 },
      { type: 'page.started', pageId: 'b', attempt: 1 },
      { type: 'page.compiled', pageId: 'b', attempt: 1 },
      { type: 'page.compiled', pageId: 'a', attempt: 1 },
      { type: 'run.completed' },
    ].map((event, index) => ({ ...event, sequence: index + 1, createdAt: at }))
    expect(parsePresentationProductionJob({ ...job, revision: events.length, state: 'completed', events }).state).toBe('completed')
    const third = [events[0], events[1], events[2], { type: 'page.started', pageId: 'c', attempt: 1, sequence: 4, createdAt: at }]
    expect(() => parsePresentationProductionJob({ ...job, revision: 4, events: third })).toThrow('invalid_state')
    const premature = [events[0], events[1], { type: 'run.completed', sequence: 3, createdAt: at }]
    expect(() => parsePresentationProductionJob({ ...job, revision: 3, state: 'completed', events: premature })).toThrow('invalid_state')
  })
  it('accepts bounded strict jobs and rejects forged history and fields', () => {
    expect(parsePresentationProductionJob(job)).toEqual(job)
    expect(() =>
      parsePresentationProductionJob({
        ...job,
        revision: 2,
        events: [
          ...job.events,
          {
            sequence: 2,
            createdAt: job.events[0]!.createdAt,
            type: 'page.compiled',
            pageId: 's',
            attempt: 1,
          },
        ],
      }),
    ).toThrow('invalid_state')
    for (const value of [
      { ...job, extra: true },
      { ...job, state: 'completed' },
      { ...job, revision: 2 },
      { ...job, events: [{ ...job.events[0], createdAt: '2026-09-24' }] },
      { ...job, events: [{ ...job.events[0], error: 'arbitrary' }] },
    ])
      expect(() => parsePresentationProductionJob(value)).toThrow('invalid_state')
  })
})

import { afterEach } from 'vitest'
import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { PresentationStore } from '../src/presentation-store.js'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function setup() {
  const root = mkdtempSync(join(tmpdir(), 'jobs-'))
  roots.push(root)
  const store = new PresentationStore(root)
  store.beginProduction(
    'p',
    'doc',
    'r',
    { slides: [{ id: 's' }] },
    { revision: 1, plan: { slides: ['s'] } },
  )
  return { root, store }
}
describe('durable production jobs', () => {
  it('persists CAS transitions and page receipt proof across restarts', () => {
    const { store, root } = setup()
    expect(store.productionJob('p', 'doc', 'r')).toBeUndefined()
    let job = store.appendProductionJobEvent('p', 'doc', 'r', 0, { type: 'run.started' })
    expect(() =>
      store.appendProductionJobEvent('p', 'doc', 'r', 0, { type: 'run.pause_requested' }),
    ).toThrow('revision_conflict')
    expect(() =>
      store.appendProductionJobEvent('p', 'doc', 'r', 1, { type: 'run.completed' }),
    ).toThrow('invalid_state')
    expect(() =>
      store.appendProductionJobEvent('p', 'doc', 'r', 1, {
        type: 'page.started',
        pageId: 's',
        attempt: 1,
      }),
    ).toThrow('invalid_state')
    let production = store.production('p', 'doc', 'r')!
    production = store.updateProductionPage(production, 's', { state: 'building', attempt: 1 })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
      type: 'page.started',
      pageId: 's',
      attempt: 1,
    })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
      type: 'run.pause_requested',
    })
    store.updateProductionPage(production, 's', {
      state: 'compiled',
      attempt: 1,
      result: { pptxBase64: 'UEsDBAAAAAA=', sourceSlideId: '256#', report: {} },
    })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
      type: 'page.compiled',
      pageId: 's',
      attempt: 1,
    })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.paused' })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.started' })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.completed' })
    expect(new PresentationStore(root).productionJob('p', 'doc', 'r')).toEqual(job)
    expect(() => store.productionJob('p', 'other', 'r')).toThrow('document_mismatch')
    expect(() =>
      store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.started' }),
    ).toThrow('invalid_state')
  })
  it('replays failed attempts after retry and rejects duplicate receipts', () => {
    const { store, root } = setup()
    let job = store.appendProductionJobEvent('p', 'doc', 'r', 0, { type: 'run.started' })
    let production = store.production('p', 'doc', 'r')!
    for (let attempt = 1; attempt <= 2; attempt++) {
      production = store.updateProductionPage(production, 's', { state: 'building', attempt })
      job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
        type: 'page.started',
        pageId: 's',
        attempt,
      })
      production = store.updateProductionPage(production, 's', {
        state: 'failed',
        attempt,
        error: 'compile_failed',
      })
      job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
        type: 'page.failed',
        pageId: 's',
        attempt,
        error: 'compile_failed',
      })
      expect(() =>
        store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
          type: 'page.failed',
          pageId: 's',
          attempt,
          error: 'compile_failed',
        }),
      ).toThrow('invalid_state')
      job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
        type: 'run.failed',
        error: 'compile_failed',
      })
      expect(new PresentationStore(root).productionJob('p', 'doc', 'r')).toEqual(job)
      if (attempt === 1)
        job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.started' })
    }
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
      type: 'run.cancel_requested',
    })
    job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.cancelled' })
    expect(() =>
      store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.started' }),
    ).toThrow('invalid_state')
  })
  it('bounds continuous history and refuses corruption without overwriting', () => {
    const { store, root } = setup()
    let job = store.appendProductionJobEvent('p', 'doc', 'r', 0, { type: 'run.started' })
    for (let n = 0; n < 50; n++) {
      job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
        type: 'run.pause_requested',
      })
      job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.paused' })
      job = store.appendProductionJobEvent('p', 'doc', 'r', job.revision, { type: 'run.started' })
    }
    expect(job.events).toHaveLength(128)
    expect(job.events[0]!.sequence).toBe(24)
    expect(new PresentationStore(root).productionJob('p', 'doc', 'r')).toEqual(job)
    const hash = (s: string) => createHash('sha256').update(s).digest('hex')
    const path = join(
      root,
      'projects',
      'presentations',
      hash('p'),
      `production-job-${hash('r')}.json`,
    )
    const record = JSON.parse(readFileSync(path, 'utf8'))
    record.job.state = 'completed'
    writeFileSync(path, JSON.stringify(record))
    expect(() => store.productionJob('p', 'doc', 'r')).toThrow('invalid_state')
    expect(() =>
      store.appendProductionJobEvent('p', 'doc', 'r', job.revision, {
        type: 'run.pause_requested',
      }),
    ).toThrow('invalid_state')
    expect(JSON.parse(readFileSync(path, 'utf8'))).toEqual(record)
  })
})
