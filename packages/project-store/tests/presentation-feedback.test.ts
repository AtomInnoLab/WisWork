// Exercise the portable checks with the POSIX-only open flag unavailable.
vi.mock('node:fs', async (original) => {
  const fs = await original<typeof import('node:fs')>()
  return { ...fs, constants: { ...fs.constants, O_NOFOLLOW: 0 } }
})
import { canonicalPresentationValue } from '../src/presentation-canonical.js'
import { expect, it, vi } from 'vitest'
import {
  parsePresentationProductionFeedbackLedger,
  parsePresentationProductionFeedbackPages,
} from '../src/presentation-feedback.js'
const ledger = {
  version: 1,
  source: 'user_reported',
  projectId: 'p',
  documentId: 'doc',
  requestId: 'r',
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  planRevision: 1,
  pageIds: ['a', 'b'],
  revision: 1,
  snapshots: [
    {
      revision: 1,
      recordedAt: '2026-09-29T00:00:00.000Z',
      pages: [
        { pageId: 'a', status: 'needs_correction', note: '原文字面说明' },
        { pageId: 'b', status: 'not_evaluated' },
      ],
    },
  ],
}
it('accepts complete strict frozen user feedback and clones metadata', () => {
  const parsed = parsePresentationProductionFeedbackLedger(ledger)
  expect(parsed).toEqual(ledger)
  parsed.snapshots[0]!.pages[0]!.note = 'changed'
  expect(ledger.snapshots[0]!.pages[0]!.note).toBe('原文字面说明')
  expect(
    parsePresentationProductionFeedbackPages([
      { pageId: 'a', status: 'no_correction', note: ' 保留 ' },
    ]),
  ).toEqual([{ pageId: 'a', status: 'no_correction', note: ' 保留 ' }])
})
it('rejects unknown fields/status, duplicate/missing pages and forged history', () => {
  for (const v of [
    { ...ledger, extra: true },
    { ...ledger, source: 'qa' },
    { ...ledger, revision: 2 },
    { ...ledger, documentId: '' },
    { ...ledger, pageIds: ['a', 'a'] },
    { ...ledger, snapshots: [{ ...ledger.snapshots[0], pages: [ledger.snapshots[0]!.pages[0]] }] },
    { ...ledger, snapshots: [{ ...ledger.snapshots[0], recordedAt: '2026-09-29' }] },
  ])
    expect(() => parsePresentationProductionFeedbackLedger(v)).toThrow('invalid_state')
  for (const pages of [
    [],
    [{ pageId: 'a', status: 'passed' }],
    [{ pageId: 'a', status: 'no_correction', extra: true }],
    [
      { pageId: 'a', status: 'no_correction' },
      { pageId: 'a', status: 'needs_correction' },
    ],
    [{ pageId: 'a', status: 'no_correction', note: '中'.repeat(667) }],
  ])
    expect(() => parsePresentationProductionFeedbackPages(pages)).toThrow('invalid_request')
})
import { afterEach } from 'vitest'
import { mkdtempSync, rmSync, readFileSync, writeFileSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { PresentationStore } from '../src/presentation-store.js'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function setup(pageCount = 2) {
  const root = mkdtempSync(join(tmpdir(), 'feedback-'))
  roots.push(root)
  const store = new PresentationStore(root)
  const pageIds = Array.from({ length: pageCount }, (_, i) =>
    i === 0 ? 'a' : i === 1 ? 'b' : 'page' + i,
  )
  let production = store.beginProduction(
    'p',
    'doc',
    'r',
    { slides: pageIds.map((id) => ({ id })) },
    { revision: 1, plan: { slides: pageIds } },
  )
  for (const page of pageIds) {
    production = store.updateProductionPage(production, page, { state: 'building', attempt: 1 })
    production = store.updateProductionPage(production, page, {
      state: 'compiled',
      attempt: 1,
      result: { pptxBase64: 'UEsDBAAAAAA=', sourceSlideId: '256#', report: {} },
    })
  }
  return { root, store, production }
}
it('persists immutable bound partial feedback, defaults unknown and repairs lost ACK without new revision', () => {
  const { store, root, production } = setup()
  expect(typeof store.recordProductionFeedback).toBe('function')
  expect(store.productionFeedback('p', 'doc', 'r')).toBeUndefined()
  const first = store.recordProductionFeedback('p', 'doc', 'r', 0, [
    { pageId: 'a', status: 'needs_correction', note: ' 原说明 ' },
  ])
  expect(first).toMatchObject({
    source: 'user_reported',
    revision: 1,
    inputDigest: production.inputDigest,
    planDigest: production.planDigest,
    pageIds: ['a', 'b'],
  })
  expect(first.snapshots[0].pages[1]).toEqual({ pageId: 'b', status: 'not_evaluated' })
  expect(
    store.recordProductionFeedback('p', 'doc', 'r', 0, [
      { pageId: 'a', status: 'needs_correction', note: ' 原说明 ' },
    ]),
  ).toEqual(first)
  const next = store.recordProductionFeedback('p', 'doc', 'r', 1, [
    { pageId: 'b', status: 'no_correction' },
  ])
  expect(next.snapshots[0]).toEqual(first.snapshots[0])
  expect(next.snapshots[1].pages[0]).toEqual(first.snapshots[0].pages[0])
  expect(new PresentationStore(root).productionFeedback('p', 'doc', 'r')).toEqual(next)
  expect(store.production('p', 'doc', 'r')).toEqual(production)
  expect(() =>
    store.recordProductionFeedback('p', 'doc', 'r', 0, [{ pageId: 'a', status: 'no_correction' }]),
  ).toThrow('revision_conflict')
  expect(() => store.productionFeedback('p', 'other', 'r')).toThrow('document_mismatch')
})
it('never discards 64 immutable revisions and still allows the final identical ACK retry', () => {
  const { root, store } = setup()
  let feedback: ReturnType<PresentationStore['recordProductionFeedback']> | undefined
  for (let revision = 0; revision < 64; revision++)
    feedback = store.recordProductionFeedback('p', 'doc', 'r', revision, [
      {
        pageId: 'a',
        status: revision % 2 ? 'no_correction' : 'needs_correction',
        note: String(revision),
      },
    ])
  expect(feedback!.snapshots).toHaveLength(64)
  expect(() =>
    store.recordProductionFeedback('p', 'doc', 'r', 64, [{ pageId: 'a', status: 'not_evaluated' }]),
  ).toThrow('output_too_large')
  expect(
    store.recordProductionFeedback('p', 'doc', 'r', 63, [
      { pageId: 'a', status: 'no_correction', note: '63' },
    ]),
  ).toEqual(feedback)
  expect(new PresentationStore(root).productionFeedback('p', 'doc', 'r')).toEqual(feedback)
})
it('rejects incomplete production and unknown page patches before durable writes', () => {
  const { store } = setup()
  store.beginProduction(
    'p',
    'doc',
    'pending',
    { slides: [{ id: 'a' }] },
    { revision: 1, plan: { slides: ['a'] } },
  )
  expect(store.productionFeedback('p', 'doc', 'pending')).toBeUndefined()
  expect(() =>
    store.recordProductionFeedback('p', 'doc', 'pending', 0, [
      { pageId: 'a', status: 'needs_correction' },
    ]),
  ).toThrow('page_not_ready')
  expect(() =>
    store.recordProductionFeedback('p', 'doc', 'r', 0, [
      { pageId: 'foreign', status: 'needs_correction' },
    ]),
  ).toThrow('invalid_request')
  expect(store.productionFeedback('p', 'doc', 'r')).toBeUndefined()
})
it('rejects corrupt or substituted frozen binding and symlink feedback files on reopen', () => {
  const { store, root } = setup()
  const saved = store.recordProductionFeedback('p', 'doc', 'r', 0, [
    { pageId: 'a', status: 'needs_correction' },
  ])
  const hash = (v: string) => createHash('sha256').update(v).digest('hex')
  const path = join(
    root,
    'projects',
    'presentations',
    hash('p'),
    `production-feedback-${hash('r')}.json`,
  )
  const original = readFileSync(path, 'utf8')
  const corrupt = JSON.parse(original)
  corrupt.feedback.snapshots[0].pages[0].status = 'no_correction'
  writeFileSync(path, JSON.stringify(corrupt))
  expect(() => new PresentationStore(root).productionFeedback('p', 'doc', 'r')).toThrow(
    'invalid_state',
  )
  corrupt.feedback = { ...saved, inputDigest: 'e'.repeat(64) }
  corrupt.checksum = hash(canonicalPresentationValue(corrupt.feedback))
  writeFileSync(path, JSON.stringify(corrupt))
  expect(() => new PresentationStore(root).productionFeedback('p', 'doc', 'r')).toThrow(
    'invalid_state',
  )
  writeFileSync(path, original)
  const target = join(root, 'unrelated.json')
  writeFileSync(target, original)
  rmSync(path)
  symlinkSync(target, path)
  expect(() => new PresentationStore(root).productionFeedback('p', 'doc', 'r')).toThrow(
    'invalid_state',
  )
  expect(() =>
    store.recordProductionFeedback('p', 'doc', 'r', 1, [{ pageId: 'a', status: 'no_correction' }]),
  ).toThrow('invalid_state')
  expect(readFileSync(target, 'utf8')).toBe(original)
})
it('bounds escaped JSON history bytes without discarding the prior durable snapshots', () => {
  const { store, root, production } = setup(32)
  const patch = production.pages.map((page) => ({
    pageId: page.pageId,
    status: 'needs_correction' as const,
    note: '\u0000'.repeat(2000),
  }))
  let previous: ReturnType<PresentationStore['recordProductionFeedback']> | undefined
  let capped = false
  for (let revision = 0; revision < 64; revision++) {
    try {
      previous = store.recordProductionFeedback('p', 'doc', 'r', revision, patch)
    } catch (error) {
      expect((error as Error).message).toBe('output_too_large')
      capped = true
      break
    }
  }
  expect(capped).toBe(true)
  expect(previous!.revision).toBeGreaterThan(0)
  expect(previous!.revision).toBeLessThan(64)
  expect(new PresentationStore(root).productionFeedback('p', 'doc', 'r')).toEqual(previous)
})
it('preserves literal UTF8 note limits and rejects forged chronology and ordered page identity', () => {
  expect(
    parsePresentationProductionFeedbackPages([
      { pageId: 'a', status: 'not_evaluated', note: '中'.repeat(666) + 'ab' },
    ])[0]!.note,
  ).toHaveLength(668)
  const first = ledger.snapshots[0]!
  for (const altered of [
    { ...ledger, revision: 0 },
    { ...ledger, revision: 65 },
    { ...ledger, snapshots: [{ ...first, extra: true }] },
    { ...ledger, snapshots: [{ ...first, pages: [...first.pages].reverse() }] },
    { ...ledger, snapshots: [{ ...first, recordedAt: '2026-02-30T00:00:00.000Z' }] },
    {
      ...ledger,
      revision: 2,
      snapshots: [first, { ...first, revision: 2, recordedAt: '2026-09-28T00:00:00.000Z' }],
    },
  ])
    expect(() => parsePresentationProductionFeedbackLedger(altered)).toThrow('invalid_state')
})
