import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, rmSync, writeFileSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { PresentationStore } from '../src/index.js'

const hash = (value: string) => createHash('sha256').update(value).digest('hex')
const review = {
  pageId: 'page1',
  claimId: 'claim1',
  sourceId: 'source1',
  attachmentId: 'attachment1',
  offset: 0,
  maxChars: 1000,
  evidenceDigest: 'a'.repeat(64),
  outcome: 'supported',
  notes: 'The quoted evidence supports this claim.',
  reviewer: 'agent',
}
let root: string
let store: PresentationStore
let path: string
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'claim-review-'))
  store = new PresentationStore(root)
  store.beginProduction(
    'project',
    'document',
    'request',
    { slides: [{ id: 'page1' }] },
    { revision: 1, plan: { title: 'frozen' } },
  )
  path = join(
    root,
    'projects',
    'presentations',
    hash('project'),
    `claim-reviews-${hash('request')}.json`,
  )
})
afterEach(() => rmSync(root, { recursive: true, force: true }))
const save = (id = 'review1', value: unknown = review) =>
  store.saveClaimReview('project', 'document', 'request', id, value)
const read = () =>
  new PresentationStore(root).claimReview('project', 'document', 'request', 'review1')

describe('claim review persistence', () => {
  it('persists frozen binding and returns an immutable idempotent record after restart', () => {
    const production = store.production('project', 'document', 'request')!
    const record = save()
    expect(record).toMatchObject({
      version: 1,
      projectId: 'project',
      documentId: 'document',
      requestId: 'request',
      reviewId: 'review1',
      inputDigest: production.inputDigest,
      planDigest: production.planDigest,
      planRevision: 1,
      review,
    })
    expect(read()).toEqual(record)
    expect(save('review1', Object.fromEntries(Object.entries(review).reverse()))).toEqual(record)
    expect(() => save('review1', { ...review, notes: 'different' })).toThrow('request_conflict')
    expect(store.production('project', 'document', 'request')).toEqual(production)
    expect(store.history('project', 'document')).toEqual([])
    store.savePlan('project', 'document', 0, { title: 'later mutable plan' })
    expect(read()).toEqual(record)
  })
  it('requires an existing production and enforces document and request isolation', () => {
    expect(() => store.saveClaimReview('project', 'other', 'request', 'r', review)).toThrow(
      'document_mismatch',
    )
    expect(() => store.saveClaimReview('project', 'document', 'missing', 'r', review)).toThrow(
      'page_not_ready',
    )
    expect(store.claimReview('project', 'document', 'missing', 'r')).toBeUndefined()
    expect(() => save('../bad')).toThrow('invalid_request')
  })
  it('enforces 32 records and allows retry at capacity', () => {
    for (let index = 0; index < 32; index++) save(`r${index}`)
    expect(save('r0').reviewId).toBe('r0')
    expect(() => save('r32')).toThrow('quota_exceeded')
  })
  it.each([
    null,
    { ...review, extra: true },
    { ...review, notes: '' },
    { ...review, notes: '\u0000' },
    { ...review, notes: 'x'.repeat(2001) },
    { ...review, offset: -1 },
    { ...review, maxChars: 0 },
    { ...review, evidenceDigest: 'bad' },
    { ...review, reviewer: 'human' },
    { ...review, outcome: 'approved' },
    { ...review, maxChars: 8001 },
    { ...review, offset: 1000001 },
    { ...review, notes: '\ud800' },
    JSON.parse('{"__proto__":{}}'),
  ])('rejects invalid review input %#', (value) => {
    expect(() => save('r', value)).toThrow('invalid_request')
  })
  it('enforces the total byte quota without altering earlier records', () => {
    const documentId = 'd'.repeat(2048)
    const other = new PresentationStore(join(root, 'quota'))
    other.beginProduction(
      'project',
      documentId,
      'request',
      { slides: [{ id: 'page1' }] },
      { revision: 1, plan: {} },
    )
    const large = {
      ...review,
      notes: '文'.repeat(2000),
      claimId: 'c'.repeat(128),
      sourceId: 's'.repeat(128),
      attachmentId: 'a'.repeat(128),
    }
    let count = 0
    for (; count < 32; count++) {
      try {
        other.saveClaimReview('project', documentId, 'request', `r${count}`, large)
      } catch (error) {
        expect((error as Error).message).toBe('quota_exceeded')
        break
      }
    }
    expect(count).toBeGreaterThan(0)
    expect(count).toBeLessThan(32)
    expect(other.claimReview('project', documentId, 'request', 'r0')?.review).toEqual(large)
    expect(other.claimReview('project', documentId, 'request', `r${count}`)).toBeUndefined()
  })
  it.each([
    'notes',
    'createdAt',
    'reviewId',
    'inputDigest',
    'planDigest',
    'planRevision',
    'duplicate',
    'extra',
    'dangerous',
  ])('rejects tampering: %s', (field) => {
    save()
    const records = JSON.parse(readFileSync(path, 'utf8'))
    if (field === 'notes') records[0].review.notes = 'tampered'
    else if (field === 'duplicate') records.push(records[0])
    else if (field === 'extra') records[0].extra = true
    else if (field === 'dangerous') records[0].review = JSON.parse('{"constructor":{}}')
    else records[0][field] = field === 'planRevision' ? 2 : 'tampered'
    writeFileSync(path, JSON.stringify(records))
    expect(read).toThrow('invalid_state')
    expect(() => save('new')).toThrow('invalid_state')
  })
  it('rejects oversized files and symbolic links', () => {
    save()
    writeFileSync(path, ' '.repeat(256 * 1024 + 1))
    expect(read).toThrow('invalid_state')
    rmSync(path)
    const target = join(root, 'external.json')
    writeFileSync(target, '[]')
    symlinkSync(target, path)
    expect(read).toThrow('invalid_state')
  })
})

function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  if (value && typeof value === 'object')
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`)
      .join(',')}}`
  return JSON.stringify(value)
}
it.each([
  'createdAt',
  'reviewId',
  'inputDigest',
  'planDigest',
  'planRevision',
  'notes',
  'maxChars',
])('rejects invalid %s even with a recomputed checksum', (field) => {
  save()
  const records = JSON.parse(readFileSync(path, 'utf8'))
  if (field === 'notes') records[0].review.notes = '\u0000'
  else if (field === 'maxChars') records[0].review.maxChars = 8001
  else
    records[0][field] =
      field === 'planRevision'
        ? 2
        : field === 'createdAt'
          ? 'yesterday'
          : field === 'reviewId'
            ? '../bad'
            : 'b'.repeat(64)
  const content = { ...records[0] }
  delete content.reviewDigest
  records[0].reviewDigest = hash(canonical(content))
  writeFileSync(path, JSON.stringify(records))
  expect(read).toThrow('invalid_state')
})

it('lists all validated immutable history across restart and isolates documents', () => {
  const one = save(),
    two = save('review2', { ...review, outcome: 'contradicted' })
  const restarted = new PresentationStore(root)
  const listed = restarted.listClaimReviews('project', 'document', 'request')
  expect(listed).toEqual([one, two])
  listed[0]!.review = {}
  expect(restarted.listClaimReviews('project', 'document', 'request')).toEqual([one, two])
  expect(restarted.listClaimReviews('project', 'document', 'missing')).toEqual([])
  expect(() => restarted.listClaimReviews('project', 'another-doc', 'request')).toThrow()
  const damaged = JSON.parse(readFileSync(path, 'utf8'))
  damaged[1].review.outcome = 'supported'
  writeFileSync(path, JSON.stringify(damaged))
  expect(() => restarted.listClaimReviews('project', 'document', 'request')).toThrow(
    'invalid_state',
  )
})

it('persists complete optional assessment and rejects changed assessment same ID after restart', () => {
  const sourceAssessment = {
    scope: '此主张',
    authority: { outcome: 'uncertain', sourceTier: 'unverified', reason: '无认证' },
    timeliness: { outcome: 'uncertain', referenceDate: '2026-09-29', reason: '未验证' },
    basis: [],
  }
  const saved = save('review1', { ...review, sourceAssessment })
  expect(read()).toEqual(saved)
  expect(save('review1', { ...review, sourceAssessment })).toEqual(saved)
  expect(() =>
    save('review1', { ...review, sourceAssessment: { ...sourceAssessment, scope: '变化' } }),
  ).toThrow('request_conflict')
  expect(() =>
    save('bad', { ...review, sourceAssessment: { ...sourceAssessment, unknown: true } }),
  ).toThrow('invalid_request')
})
it('enforces unchanged total quota and validates assessed corruption even with recomputed digest', () => {
  const sourceAssessment = {
    scope: '界'.repeat(400),
    authority: { outcome: 'uncertain', sourceTier: 'unverified', reason: '界'.repeat(600) },
    timeliness: {
      outcome: 'uncertain',
      referenceDate: '2026-09-29',
      claimAsOf: '界'.repeat(100),
      sourceAsOf: '界'.repeat(100),
      reason: '界'.repeat(600),
    },
    jurisdiction: {
      claimJurisdiction: '界'.repeat(400),
      outcome: 'uncertain',
      reason: '界'.repeat(600),
    },
    basis: Array.from({ length: 4 }, (_, offset) => ({ offset, text: '界'.repeat(600) })),
  }
  const value = { ...review, notes: '界'.repeat(2000), sourceAssessment }
  expect(Buffer.byteLength(JSON.stringify(value))).toBeGreaterThan(8 * 1024)
  const first = save('review1', value)
  expect(read()).toEqual(first)
  let written = 1
  for (; written < 32; written++) {
    try {
      save(`assessed${written}`, value)
    } catch (error) {
      expect((error as Error).message).toBe('quota_exceeded')
      break
    }
  }
  expect(written).toBeLessThan(32)
  expect(read()).toEqual(first)
  const records = JSON.parse(readFileSync(path, 'utf8'))
  records[0].review.sourceAssessment.authority.outcome = 'certified'
  const { reviewDigest: _old, ...content } = records[0]
  const canonical = (v: unknown): unknown =>
    v && typeof v === 'object' && !Array.isArray(v)
      ? Object.fromEntries(
          Object.entries(v)
            .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
            .map(([k, v]) => [k, canonical(v)]),
        )
      : Array.isArray(v)
        ? v.map(canonical)
        : v
  records[0].reviewDigest = hash(JSON.stringify(canonical(content)))
  writeFileSync(path, JSON.stringify(records))
  expect(() => read()).toThrow('invalid_state')
})
