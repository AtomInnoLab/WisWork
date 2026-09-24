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
