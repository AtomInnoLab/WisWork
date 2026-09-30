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
it('accepts unbound professional full claim and rejects partial omissions and malformed context', () => {
  const claim = {
    id: 'claim',
    statement: 'claim',
    type: 'fact',
    sourceIds: ['source'],
    confidence: 'high',
    reviewStatus: 'needs_review',
    professionalContext: { domain: 'science' },
  }
  const value = {
    ...report(),
    documentId: 'doc',
    claim,
    source: { ...report().source, asOf: 'Declared date' },
  }
  expect(parsePresentationClaimEvidence(value)).toEqual(value)
  for (const patch of [
    { documentId: undefined },
    { claim: { ...claim, professionalContext: undefined } },
    { claim: { ...claim, professionalContext: { domain: 'science', currency: 'USD' } } },
    { claim: { ...claim, professionalContext: { domain: 'law', effectiveFrom: '2026-02-30' } } },
  ]) {
    expect(() => parsePresentationClaimEvidence({ ...value, ...patch })).toThrow()
  }
  expect(parsePresentationClaimEvidence(report())).toEqual(report())
})
it('binds a repeated excerpt to the preferred indexed page and rejects forged locators', () => {
  const value = {
    ...report(),
    source: { ...report().source, locator: '第2页' },
    attachment: {
      ...report().attachment,
      name: 'study.pdf',
      offset: 0,
      totalChars: 8,
      text: '原文\n\n原文',
      locatorSpans: [
        { locator: '第 1 页', start: 0, end: 2 },
        { locator: '第 2 页', start: 4, end: 6 },
      ],
    },
    excerptMatch: { status: 'found', offset: 4, locator: '第 2 页' },
  }
  expect(parsePresentationClaimEvidence(value).excerptMatch).toEqual(value.excerptMatch)
  expect(() =>
    parsePresentationClaimEvidence({
      ...value,
      excerptMatch: { status: 'found', offset: 0, locator: '第 2 页' },
    }),
  ).toThrow()
  expect(() =>
    parsePresentationClaimEvidence({
      ...value,
      attachment: { ...value.attachment, name: 'study.html' },
    }),
  ).toThrow()
})
it('preserves image-backed PDF spans and a scanned blank page without claiming OCR verification', () => {
  const value = {
    ...report(),
    source: { ...report().source, excerpt: 'BCDE', locator: '第 3 页' },
    attachment: {
      ...report().attachment,
      name: 'scan.pdf',
      offset: 0,
      totalChars: 9,
      text: 'A\n\n\n\nBCDE',
      locatorSpans: [
        { locator: '第 1 页', start: 0, end: 1, imageBacked: true },
        { locator: '第 2 页', start: 3, end: 3, imageBacked: true },
        { locator: '第 3 页', start: 5, end: 9, imageBacked: true },
      ],
    },
    excerptMatch: { status: 'found', offset: 5, locator: '第 3 页' },
  }
  expect(parsePresentationClaimEvidence(value).attachment.locatorSpans).toEqual(
    value.attachment.locatorSpans,
  )
  expect(() =>
    parsePresentationClaimEvidence({
      ...value,
      attachment: {
        ...value.attachment,
        name: 'scan.html',
      },
    }),
  ).toThrow()
})
it('prefers a contained excerpt when an earlier match crosses page boundaries', () => {
  expect(
    matchPresentationClaimExcerpt('X\n\nY', 'X\n\nY\n\nX\n\nY', 0, [
      { locator: '第 1 页', start: 0, end: 1 },
      { locator: '第 2 页', start: 3, end: 4 },
      { locator: '第 3 页', start: 6, end: 10 },
    ]),
  ).toEqual({ status: 'found', offset: 6, locator: '第 3 页' })
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
    { attachment: { ...report().attachment, totalChars: 8_000_001 } },
    { planRevision: 1.5 },
  ])
    expect(() => parsePresentationClaimEvidence({ ...report(), ...patch })).toThrow()
})
it('accepts source offsets beyond one million characters', () => {
  const value = report()
  value.attachment.offset = 4_000_000
  value.attachment.totalChars = 4_338_831
  value.excerptMatch.offset = 4_000_002
  expect(parsePresentationClaimEvidence(value).attachment.offset).toBe(4_000_000)
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
it('requires fetched URL provenance to carry a bounded retrieval time and snapshot', () => {
  const value = report()
  const source = {
    ...value.source,
    uri: 'https://example.com/original',
    snapshotAttachmentId: hash,
  }
  const attachment = {
    ...value.attachment,
    provenance: { binding: 'fetched_url_matched', retrievedAt: 1790590000000 },
  }
  expect(
    parsePresentationClaimEvidence({ ...value, source, attachment }).attachment.provenance,
  ).toEqual(attachment.provenance)
  for (const changed of [
    { ...attachment, provenance: { binding: 'fetched_url_matched' } },
    { ...attachment, provenance: { binding: 'fetched_url_matched', retrievedAt: 1.5 } },
    { ...attachment, provenance: { binding: 'user_supplied', retrievedAt: 1790590000000 } },
  ])
    expect(() =>
      parsePresentationClaimEvidence({ ...value, source, attachment: changed }),
    ).toThrow()
  expect(() => parsePresentationClaimEvidence({ ...value, attachment })).toThrow()
})
import { researchFixture } from './fixtures/presentation-research'
import { presentationResearchBindingFindings } from '../src/presentation-research-binding'
it('retains complete bound frozen claim and original research with professional qualifiers', () => {
  const { plan, record } = researchFixture()
  const source = plan.sources[0]!,
    claim = plan.claims[0]!
  const evidence = {
    ...report(),
    projectId: plan.projectId,
    claimId: claim.id,
    statement: claim.statement,
    documentId: record.documentId,
    claim,
    source: {
      id: source.id,
      uri: source.uri,
      snapshotAttachmentId: source.snapshotAttachmentId,
      excerpt: source.excerpt,
      locator: source.locator,
      asOf: source.asOf,
    },
    research: {
      binding: plan.research!,
      record,
      findings: presentationResearchBindingFindings(plan, record).filter(
        (f) => f.claimId === claim.id,
      ),
    },
  }
  evidence.source.uri = `attachment:${hash}`
  delete evidence.source.snapshotAttachmentId
  evidence.source.excerpt = '原文'
  delete evidence.source.locator
  delete evidence.source.asOf
  // Bind the same actual source literal used by the original text window.
  record.draft.sources[0]!.uri = evidence.source.uri
  record.draft.sources[0]!.excerpt = evidence.source.excerpt
  delete record.draft.sources[0]!.locator
  delete record.draft.sources[0]!.snapshotAttachmentId
  delete record.draft.sources[0]!.asOf
  expect(parsePresentationClaimEvidence(evidence).research?.record).toEqual(record)
})
