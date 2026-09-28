import { expect, it } from 'vitest'
import {
  parsePresentationClaimEvidence,
  matchPresentationClaimExcerpt,
} from '../src/presentation-claim-evidence'
const hash = 'a'.repeat(64)
const report = () => ({
  version: 1,
  projectId: 'project',
  requestId: 'run',
  planRevision: 1,
  inputDigest: hash,
  planDigest: hash,
  pageId: 'page',
  claimId: 'claim',
  statement: 'claim',
  source: { id: 'source', uri: `attachment:${hash}`, excerpt: '原文' },
  attachment: {
    id: hash,
    name: 'file.txt',
    offset: 2,
    totalChars: 8,
    text: '😀原文',
    offsetUnit: 'utf16_code_unit',
  },
  excerptMatch: { status: 'found', offset: 4 },
  checks: {
    support: 'not_verified',
    sourceAuthority: 'not_verified',
    timeliness: 'not_verified',
    host: 'not_checked',
  },
})
it('parses a report and checks UTF16 literal offsets', () => {
  expect(parsePresentationClaimEvidence(report())).toEqual(report())
  expect(matchPresentationClaimExcerpt('原文', '😀原文', 2)).toEqual({ status: 'found', offset: 4 })
  expect(matchPresentationClaimExcerpt(' 原文 ', '原文', 0)).toEqual({
    status: 'not_found_in_window',
  })
  expect(matchPresentationClaimExcerpt(' \n', '原文', 0)).toEqual({ status: 'empty_excerpt' })
})
it('rejects forged matches, identity, bounds, checks and extra fields', () => {
  for (const patch of [
    { extra: true },
    { excerptMatch: { status: 'found', offset: 3 } },
    { excerptMatch: { status: 'not_found_in_window' } },
    { attachment: { ...report().attachment, id: 'b'.repeat(64) } },
    { attachment: { ...report().attachment, offset: 2.5 } },
    { attachment: { ...report().attachment, totalChars: 3 } },
    { checks: { ...report().checks, support: 'verified' } },
    { source: { ...report().source, uri: 'https://example.com' } },
  ])
    expect(() => parsePresentationClaimEvidence({ ...report(), ...patch })).toThrow()
})

it('accepts persisted 128-character request IDs', () => {
  expect(
    parsePresentationClaimEvidence({ ...report(), requestId: 'r'.repeat(128) }).requestId,
  ).toHaveLength(128)
})
it('rejects missing required fields, nested extras and oversized windows', () => {
  for (const key of Object.keys(report())) {
    const value: Record<string, unknown> = report()
    delete value[key]
    expect(() => parsePresentationClaimEvidence(value)).toThrow()
  }
  for (const patch of [
    { source: { ...report().source, extra: true } },
    { excerptMatch: { status: 'not_found_in_window', offset: 0 } },
    { attachment: { ...report().attachment, text: 'x'.repeat(8001), totalChars: 10000 } },
    { attachment: { ...report().attachment, totalChars: 1000001 } },
    { planRevision: 1.5 },
  ])
    expect(() => parsePresentationClaimEvidence({ ...report(), ...patch })).toThrow()
})
it('accepts empty excerpts and empty end windows without claiming global absence', () => {
  expect(
    parsePresentationClaimEvidence({
      ...report(),
      source: { ...report().source, excerpt: ' \n' },
      excerptMatch: { status: 'empty_excerpt' },
    }).excerptMatch.status,
  ).toBe('empty_excerpt')
  expect(
    parsePresentationClaimEvidence({
      ...report(),
      attachment: { ...report().attachment, offset: 8, text: '' },
      excerptMatch: { status: 'not_found_in_window' },
    }).excerptMatch.status,
  ).toBe('not_found_in_window')
})

it('preserves form-feed parsed text and literal offsets', () => {
  const value = {
    ...report(),
    attachment: { ...report().attachment, text: '\f原文' },
    excerptMatch: { status: 'found', offset: 3 },
  }
  expect(parsePresentationClaimEvidence(value)).toEqual(value)
})

it('binds external original URI evidence to an explicit snapshot', () => {
  const value = report()
  const source = {
    ...value.source,
    uri: 'https://example.com/original',
    snapshotAttachmentId: hash,
  }
  expect(parsePresentationClaimEvidence({ ...value, source }).source).toEqual(source)
  for (const changed of [
    { ...source, snapshotAttachmentId: 'b'.repeat(64) },
    { ...source, snapshotAttachmentId: '' },
    { ...value.source, snapshotAttachmentId: 'b'.repeat(64) },
    { ...value.source, uri: source.uri },
  ])
    expect(() => parsePresentationClaimEvidence({ ...value, source: changed })).toThrow(
      'presentation_claim_evidence_invalid:',
    )
})
