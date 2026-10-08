import { expect, it } from 'vitest'
import {
  parsePresentationQaAttempt,
  presentationQaAttemptIdentity,
} from '../src/skills/powerpoint/presentation-qa-attempts'
const base = {
  version: 1,
  id: '12345678-1234-4234-8234-123456789abc',
  documentId: 'doc',
  projectId: 'project',
  requestId: 'request',
  artifactDigest: 'a'.repeat(64),
  pageId: 'page',
  hostSlideId: 'host',
  startedAt: '2026-09-29T00:00:00.000Z',
  status: 'started',
}
it('clones strict start and bounded terminal records without raw errors or content', () => {
  expect(parsePresentationQaAttempt(base)).toEqual(base)
  expect(parsePresentationQaAttempt(base)).not.toBe(base)
  for (const [status, errorCode] of [
    ['recorded', undefined],
    ['waiting', 'screenshot_unavailable'],
    ['failed', 'inspection_failed'],
    ['failed', 'publication_failed'],
    ['failed', 'state_changed'],
    ['cancelled', 'cancelled'],
  ]) {
    const value = {
      ...base,
      status,
      finishedAt: base.startedAt,
      ...(errorCode ? { errorCode } : {}),
    }
    expect(parsePresentationQaAttempt(value)).toEqual(value)
  }
})
it.each([
  { ...base, id: base.id.toUpperCase() },
  { ...base, extra: true },
  { ...base, startedAt: '2026-09-29T00:00:00Z' },
  { ...base, finishedAt: base.startedAt },
  { ...base, errorCode: 'cancelled' },
  { ...base, status: 'recorded' },
  { ...base, status: 'recorded', finishedAt: base.startedAt, errorCode: 'state_changed' },
  { ...base, status: 'waiting', finishedAt: base.startedAt, errorCode: 'inspection_failed' },
  { ...base, status: 'failed', finishedAt: base.startedAt, errorCode: 'cancelled' },
  { ...base, status: 'cancelled', finishedAt: '2000-01-01T00:00:00.000Z', errorCode: 'cancelled' },
  { ...base, source: undefined },
  { ...base, documentId: 'x'.repeat(4097) },
  { ...base, rawError: 'private' },
])('rejects malformed attempt %#', (value) =>
  expect(() => parsePresentationQaAttempt(value)).toThrow('presentation_qa_attempt_state_invalid'),
)
it('preserves the existing text request identity including punctuation', () => {
  const value = { ...base, requestId: 'request/one:1' }
  expect(parsePresentationQaAttempt(value)).toEqual(value)
})

it('accepts only explicit close with canonical terminal time and preserves original identity', () => {
  const start = parsePresentationQaAttempt(base)
  const closed = {
    ...start,
    status: 'closed',
    finishedAt: start.startedAt,
    errorCode: 'explicitly_closed',
  }
  const parsed = parsePresentationQaAttempt(closed)
  expect(parsed).toEqual(closed)
  expect(presentationQaAttemptIdentity(parsed)).toBe(presentationQaAttemptIdentity(start))
  expect(presentationQaAttemptIdentity({ ...parsed, pageId: 'other' })).not.toBe(
    presentationQaAttemptIdentity(start),
  )
  for (const patch of [
    { errorCode: undefined },
    { errorCode: 'cancelled' },
    { status: 'cancelled' },
    { status: 'failed' },
    { status: 'unknown' },
    { finishedAt: '2026-09-29T00:00:00Z' },
    { finishedAt: '2000-01-01T00:00:00.000Z' },
  ])
    expect(() => parsePresentationQaAttempt({ ...closed, ...patch })).toThrow(
      'presentation_qa_attempt_state_invalid',
    )
})
