import { expect, it } from 'vitest'
import {
  parsePresentationClaimReview,
  presentationClaimEvidenceContent,
} from '../src/presentation-claim-review'
const hash = 'a'.repeat(64)
const report = () => ({
  version: 1,
  projectId: 'project',
  requestId: 'run',
  reviewId: 'review',
  planRevision: 1,
  inputDigest: hash,
  planDigest: hash,
  pageId: 'page',
  claimId: 'claim',
  sourceId: 'source',
  attachmentId: hash,
  offset: 0,
  maxChars: 8000,
  evidenceDigest: hash,
  outcome: 'supported',
  notes: 'Agent judgment',
  reviewer: 'agent',
  createdAt: '2026-09-24T00:00:00.000Z',
  checks: {
    support: 'agent_reviewed',
    sourceAuthority: 'not_verified',
    timeliness: 'not_verified',
    host: 'not_checked',
  },
})
it('validates complete immutable agent reports and rejects forged verification and invalid metadata', () => {
  expect(parsePresentationClaimReview(report())).toEqual(report())
  for (const patch of [
    { extra: true },
    { notes: ' ' },
    { notes: '\u0000' },
    { notes: 'x'.repeat(2001) },
    { outcome: 'verified' },
    { reviewer: 'human' },
    { offset: 1.5 },
    { maxChars: 8001 },
    { createdAt: 'yesterday' },
    { createdAt: '2026-02-30T00:00:00.000Z' },
    { checks: { ...report().checks, support: 'verified' } },
  ])
    expect(() => parsePresentationClaimReview({ ...report(), ...patch })).toThrow()
  for (const key of Object.keys(report())) {
    const value: Record<string, unknown> = report()
    delete value[key]
    expect(() => parsePresentationClaimReview(value)).toThrow()
  }
})
it('canonicalizes all validated evidence fields independently of property order', () => {
  const evidence = {
    version: 1,
    projectId: 'project',
    requestId: 'run',
    planRevision: 1,
    inputDigest: hash,
    planDigest: hash,
    pageId: 'page',
    claimId: 'claim',
    statement: 'claim',
    source: { id: 'source', uri: `attachment:${hash}`, excerpt: '' },
    attachment: {
      id: hash,
      name: 'file',
      offset: 0,
      totalChars: 2,
      text: '\fX',
      offsetUnit: 'utf16_code_unit',
    },
    excerptMatch: { status: 'empty_excerpt' },
    checks: {
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  }
  const content = presentationClaimEvidenceContent(evidence)
  expect(JSON.parse(content)).toEqual(evidence)
  expect(
    presentationClaimEvidenceContent(Object.fromEntries(Object.entries(evidence).reverse())),
  ).toBe(content)
  expect(presentationClaimEvidenceContent({ ...evidence, statement: 'changed' })).not.toBe(content)
  expect(
    presentationClaimEvidenceContent({
      ...evidence,
      attachment: { ...evidence.attachment, text: '\fY' },
    }),
  ).not.toBe(content)
})
